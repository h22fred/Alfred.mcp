import { connectWithRetry } from "../auth/tokenExtractor.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";
import { loadCachedAuth, saveCachedAuth, clearCachedAuthFile } from "../auth/authFileCache.js";
import { stripHtml, urlHostMatches } from "../shared.js";
import { sanitizeODataSearch } from "./dynamicsClient.js";

const CDP_PORT = 9222;
const OUTLOOK_ORIGIN = "https://outlook.office.com";
const TOKEN_CACHE_FALLBACK_MS = 45 * 60 * 1000; // 45 min — fallback when MSAL doesn't report expiry
const TOKEN_REFRESH_MARGIN_MS = 5 * 60 * 1000;  // refresh 5 min before expiry

interface TokenCache {
  token: string;
  expiresAt: number;
}
let tokenCache: TokenCache | null = null;

export function clearGraphTokenCache(): void {
  tokenCache = null;
  clearCachedAuthFile("graphToken");
}

/** @internal — test-only helper to pre-seed the Graph token cache */
export function _seedGraphTokenCache(token: string, ttlMs = TOKEN_CACHE_FALLBACK_MS): void {
  tokenCache = { token, expiresAt: Date.now() + ttlMs };
}

// ---------------------------------------------------------------------------
// Acquire a Graph Bearer token via raw CDP WebSocket — no Playwright needed.
// Enables Network tracking on the Outlook tab, triggers a lightweight OWA
// service call, and captures the outgoing Authorization header.
// ---------------------------------------------------------------------------

// JS snippet injected into the Outlook page to extract a Bearer token + expiry from MSAL cache
// Returns { token, expiresAt } or null
const MSAL_EXTRACT_JS = `(function() {
  const tryStorage = (s) => {
    try {
      for (const key of Object.keys(s)) {
        try {
          const val = JSON.parse(s.getItem(key));
          if (val && val.credentialType === 'AccessToken' && val.secret &&
              (val.target || '').toLowerCase().includes('mail')) {
            const exp = Number(val.expiresOn || val.extended_expires_on || 0);
            if (!exp || exp * 1000 > Date.now()) {
              return { token: val.secret, expiresAt: exp ? exp * 1000 : 0 };
            }
          }
        } catch {}
      }
    } catch {}
    return null;
  };
  return tryStorage(sessionStorage) || tryStorage(localStorage);
})()`;

async function acquireGraphTokenRawCDP(progress: ProgressFn): Promise<string> {
  if (tokenCache && Date.now() < tokenCache.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    const mins = Math.round((tokenCache.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Graph token (~${mins} min remaining)`);
    return tokenCache.token;
  }

  // Check file cache before hitting CDP — apply refresh margin
  const fileCached = loadCachedAuth("graphToken");
  if (fileCached && Date.now() < fileCached.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    tokenCache = { token: fileCached.value, expiresAt: fileCached.expiresAt };
    const mins = Math.round((fileCached.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Graph token (~${mins} min remaining)`);
    return fileCached.value;
  }

  progress("🔑 Acquiring Graph Bearer token via CDP...");

  const listRes = await fetch(`http://localhost:${CDP_PORT}/json/list`);
  const targets = await listRes.json() as Array<{ webSocketDebuggerUrl?: string; type?: string; url?: string }>;

  // Prefer Outlook tab, but also try Teams or any tab that might have a mail-scoped MSAL token
  const outlookTarget = targets.find(t => t.type === "page" && t.url && urlHostMatches(t.url, "outlook.office.com") && t.webSocketDebuggerUrl);
  const teamsTarget   = targets.find(t => t.type === "page" && t.url && urlHostMatches(t.url, "teams.microsoft.com") && t.webSocketDebuggerUrl);
  const anyTarget     = targets.find(t => t.type === "page" && t.webSocketDebuggerUrl);
  const target        = outlookTarget ?? teamsTarget ?? anyTarget;

  if (!target?.webSocketDebuggerUrl) {
    throw new Error("No page targets found in Alfred. Open Alfred.app and log into Outlook.");
  }

  if (!outlookTarget) {
    process.stderr.write(`[alfred] CDP: no Outlook tab found — trying ${teamsTarget ? "Teams" : "other"} tab for MSAL cache\n`);
  }

  // Detect if the Outlook tab is stuck on a login/redirect page
  if (outlookTarget?.url) {
    const tabUrl = outlookTarget.url;
    if (tabUrl.includes("login.microsoftonline.com") || tabUrl.includes("login.live.com") || tabUrl.includes("/oauth2/")) {
      throw new Error(
        "Outlook tab is on the Microsoft login page — your session has expired.\n" +
        "Please log back into Outlook in the Alfred window (make sure the inbox is fully loaded), then retry."
      );
    }
  }

  // Step 1: try reading token directly from MSAL storage — fast, no network interception
  const msalResult = await new Promise<{ token: string; expiresAt: number } | null>((resolve) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    const timer = setTimeout(() => { try { ws.close(); } catch {} resolve(null); }, 5_000);

    ws.addEventListener("open", () => {
      ws.send(JSON.stringify({ id: 1, method: "Runtime.evaluate", params: { expression: MSAL_EXTRACT_JS, returnByValue: true } }));
    });
    ws.addEventListener("message", (event: MessageEvent) => {
      clearTimeout(timer);
      try { ws.close(); } catch {}
      try {
        const msg = JSON.parse(event.data as string) as { result?: { result?: { value?: { token: string; expiresAt: number } | null } } };
        resolve(msg.result?.result?.value ?? null);
      } catch { resolve(null); }
    });
    ws.addEventListener("error", () => { clearTimeout(timer); try { ws.close(); } catch {} resolve(null); });
  });

  if (msalResult) {
    // Use real MSAL expiry when available, otherwise fall back to 45 min
    const expiresAt = msalResult.expiresAt > 0 ? msalResult.expiresAt : Date.now() + TOKEN_CACHE_FALLBACK_MS;
    tokenCache = { token: msalResult.token, expiresAt };
    saveCachedAuth("graphToken", msalResult.token, expiresAt);
    const mins = Math.round((expiresAt - Date.now()) / 60_000);
    progress(`✅ Graph token acquired from MSAL cache (~${mins} min valid)`);
    return msalResult.token;
  }

  // Step 2: fallback — enable Network tracking and trigger an OWA API fetch
  progress("🔑 MSAL cache miss — capturing token via network interception...");

  return new Promise((resolve, reject) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    let capturedToken: string | null = null;

    const timer = setTimeout(() => {
      try { ws.close(); } catch {}
      process.stderr.write("[alfred] CDP: Graph token capture timed out from Outlook tab\n");
      reject(new Error(
        "Could not capture Graph token from Outlook.\n" +
        "Make sure you are logged into Outlook in the Alfred window."
      ));
    }, 10_000);

    const done = (token: string) => {
      clearTimeout(timer);
      try { ws.close(); } catch {}
      const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
      tokenCache = { token, expiresAt };
      saveCachedAuth("graphToken", token, expiresAt);
      progress("✅ Graph token acquired");
      resolve(token);
    };

    let msgId = 0;
    const send = (method: string, params?: Record<string, unknown>) =>
      ws.send(JSON.stringify({ id: ++msgId, method, params }));

    ws.addEventListener("open", () => send("Network.enable"));

    ws.addEventListener("message", (event: MessageEvent) => {
      try {
        const msg = JSON.parse(event.data as string) as { id?: number; method?: string; params?: Record<string, unknown> };
        if (msg.id === 1) {
          // Network enabled — trigger an OWA mail fetch which adds Authorization via JS interceptor
          send("Runtime.evaluate", {
            expression: `fetch('/api/v2.0/me/calendarview?$top=1&$select=Id&startDateTime=${new Date().toISOString()}&endDateTime=${new Date(Date.now()+86400000).toISOString()}', { credentials: 'include' }).catch(()=>{})`,
            awaitPromise: false,
          });
        }
        if (msg.method === "Network.requestWillBeSent") {
          const headers = ((msg.params?.request as Record<string, unknown>)?.headers ?? {}) as Record<string, string>;
          const auth = headers["Authorization"] ?? headers["authorization"] ?? "";
          if (!capturedToken && auth.startsWith("Bearer ")) {
            capturedToken = auth.slice(7);
            done(capturedToken);
          }
        }
      } catch (e) { process.stderr.write(`[alfred:warn] CDP message parse error: ${e instanceof Error ? e.message : String(e)}\n`); }
    });

    ws.addEventListener("error", () => {
      clearTimeout(timer);
      reject(new Error("CDP WebSocket error — is Alfred running?"));
    });
  });
}

// acquireGraphToken — tries raw CDP first (fast), falls back to Playwright page load
async function acquireGraphToken(progress: ProgressFn): Promise<string> {
  try {
    return await acquireGraphTokenRawCDP(progress);
  } catch (e) {
    process.stderr.write(`[alfred:warn] raw CDP token acquisition failed, falling back to Playwright: ${e instanceof Error ? e.message : String(e)}\n`);
    // Raw CDP failed — use Playwright to capture token from an existing Outlook tab
    // (or navigate an existing tab to Outlook — avoids opening a new window)
    progress("📡 Falling back to Playwright token capture via Outlook...");
    const browser = await connectWithRetry();
    try {
      const ctx = browser.contexts()[0];
      if (!ctx) throw new Error("No browser context found in Alfred");

      // Try to find an existing Outlook tab instead of opening a new one
      const existingPages = ctx.pages();
      let page = existingPages.find(p => urlHostMatches(p.url(), "outlook.office.com"));

      let isNewPage = false;
      if (!page) {
        // No Outlook tab — reuse any existing tab or open one as last resort
        page = existingPages.find(p => p.url().startsWith("http")) ?? await ctx.newPage();
        isNewPage = !existingPages.includes(page);
      }

      let capturedToken: string | null = null;
      await page.route("**/*", async (route) => {
        const auth = route.request().headers()["authorization"] ?? "";
        if (!capturedToken && auth.startsWith("Bearer ")) capturedToken = auth.slice(7);
        await route.continue();
      });

      // Navigate/reload to trigger Graph API calls
      if (!urlHostMatches(page.url(), "outlook.office.com")) {
        await page.goto("https://outlook.office.com/mail/", { waitUntil: "domcontentloaded", timeout: 20_000 }).catch((e) => { process.stderr.write(`[alfred:warn] Outlook page.goto failed: ${e instanceof Error ? e.message : String(e)}\n`); });
      } else {
        await page.reload({ waitUntil: "domcontentloaded", timeout: 20_000 }).catch((e) => { process.stderr.write(`[alfred:warn] Outlook page.reload failed: ${e instanceof Error ? e.message : String(e)}\n`); });
      }

      const deadline = Date.now() + 10_000;
      while (!capturedToken && Date.now() < deadline) await page.waitForTimeout(500);

      // Clean up route interceptor
      await page.unroute("**/*").catch((e) => { process.stderr.write(`[alfred:warn] unroute failed: ${e instanceof Error ? e.message : String(e)}\n`); });
      // Only close pages we created
      if (isNewPage) await page.close();

      if (!capturedToken) throw new Error("Could not capture Graph token from Outlook.\nMake sure you are logged into Outlook in the Alfred window.");
      const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
      tokenCache = { token: capturedToken, expiresAt };
      saveCachedAuth("graphToken", capturedToken, expiresAt);
      progress("✅ Graph token acquired via Playwright");
      return capturedToken;
    } finally {
      await browser.close();
    }
  }
}

const OUTLOOK_API = "https://outlook.office.com/api/v2.0/me";

// ---------------------------------------------------------------------------
// Outlook REST v2 fetch using Bearer token (kept for Teams Graph API usage)
// ---------------------------------------------------------------------------

async function outlookApiFetch(path: string, token: string, progress?: ProgressFn, _retryCount = 0): Promise<Record<string, unknown>> {
  const res = await fetch(`${OUTLOOK_API}${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
    signal: AbortSignal.timeout(30_000),
  });

  if (res.status === 429 && _retryCount < 3) {
    const retryAfter = parseInt(res.headers.get("Retry-After") ?? "", 10);
    const delayMs = retryAfter > 0 ? retryAfter * 1000 : 1000 * Math.pow(2, _retryCount);
    progress?.(`⏳ Outlook API throttled (429) — retrying in ${(delayMs / 1000).toFixed(0)}s...`);
    await new Promise(r => setTimeout(r, delayMs));
    return outlookApiFetch(path, token, progress, _retryCount + 1);
  }

  if (res.status === 401) {
    tokenCache = null;
    clearCachedAuthFile("graphToken");
    progress?.("🔄 Graph token expired — re-acquiring...");
    const freshToken = await acquireGraphToken(progress ?? (() => {}));
    const retry = await fetch(`${OUTLOOK_API}${path}`, {
      headers: { Authorization: `Bearer ${freshToken}`, Accept: "application/json" },
      signal: AbortSignal.timeout(30_000),
    });
    if (!retry.ok) {
      const body = await retry.text().catch(() => "");
      throw new Error(`Outlook API ${retry.status} ${retry.statusText}${body ? `: ${body.slice(0, 200)}` : ""}`);
    }
    return retry.json() as Promise<Record<string, unknown>>;
  }

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Outlook API ${res.status} ${res.statusText}${body ? `: ${body.slice(0, 200)}` : ""}`);
  }

  return res.json() as Promise<Record<string, unknown>>;
}


// ---------------------------------------------------------------------------
// Calendar events
// ---------------------------------------------------------------------------

export interface CalendarEvent {
  id: string;
  subject: string;
  start: string;
  end: string;
  location?: string;
  organizer?: string;
  organizerEmail?: string;
  attendees?: { name: string; email: string }[];
  isOnlineMeeting?: boolean;
  bodyPreview?: string;
  webLink?: string;
}

export async function getCalendarEvents(
  startDate: string,
  endDate: string,
  search?: string,
  progress: ProgressFn = () => {},
  top: number = 100,
): Promise<CalendarEvent[]> {
  progress(`📅 Fetching calendar events ${startDate} → ${endDate}${search ? ` (filter: "${search}")` : ""}...`);

  // NOTE: /calendarview does NOT support $search — we filter client-side below.
  // Also skip the massive Graph "id" field to reduce payload size.
  const params = new URLSearchParams({
    startDateTime: `${startDate}T00:00:00Z`,
    endDateTime:   `${endDate}T23:59:59Z`,
    $select: "subject,start,end,location,organizer,attendees,isOnlineMeeting,webLink",
    $top: String(top),
    $orderby: "start/dateTime",
  });

  // If search filter is present, use $filter contains(subject,...) where supported,
  // but calendarView may not support it — so we always apply client-side filter as fallback.
  if (search) {
    try {
      const safe = sanitizeODataSearch(search);
      params.set("$filter", `contains(subject,'${safe}')`);
    } catch (e) {
      process.stderr.write(`[alfred:warn] OData search sanitize failed, using client-side filter: ${e instanceof Error ? e.message : String(e)}\n`);
    }
  }

  const token = await acquireGraphToken(progress);
  let res = await fetch(`https://graph.microsoft.com/v1.0/me/calendarview?${params}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
    signal: AbortSignal.timeout(30_000),
  });

  // If $filter is not supported, retry without it
  if (!res.ok && search && params.has("$filter")) {
    progress("⚠️ Server-side filter not supported on calendarView — falling back to client-side filter...");
    params.delete("$filter");
    res = await fetch(`https://graph.microsoft.com/v1.0/me/calendarview?${params}`, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
      signal: AbortSignal.timeout(30_000),
    });
  }

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Calendar API ${res.status} ${res.statusText}${body ? `: ${body.slice(0, 200)}` : ""}`);
  }
  const data = await res.json() as { value?: Record<string, unknown>[] };

  let events = (data.value ?? []).map(e => {
    const org = (e.organizer as { emailAddress: { name: string; address: string } } | undefined)?.emailAddress;
    const rawAttendees = e.attendees as Array<{ emailAddress: { name: string; address: string } }> ?? [];
    const attendees = rawAttendees.map(a => ({
      name:  a.emailAddress?.name  ?? "",
      email: a.emailAddress?.address ?? "",
    }));
    return {
      id:              "", // intentionally omitted — Graph IDs are ~200 char base64, not needed downstream
      subject:         e.subject as string,
      start:           (e.start as { dateTime: string })?.dateTime,
      end:             (e.end   as { dateTime: string })?.dateTime,
      location:        (e.location as { displayName: string })?.displayName || undefined,
      organizer:       org?.name || undefined,
      organizerEmail:  org?.address || undefined,
      attendees,
      isOnlineMeeting: e.isOnlineMeeting as boolean,
      bodyPreview:     undefined, // dropped to save context — use search_emails for full content
      webLink:         (e.webLink as string) || undefined,
    };
  });

  // Client-side search filter (always applied when search is set, in case server-side filter was not used)
  if (search) {
    const needle = search.toLowerCase();
    const beforeCount = events.length;
    events = events.filter(e =>
      e.subject?.toLowerCase().includes(needle) ||
      e.organizer?.toLowerCase().includes(needle) ||
      e.attendees.some(a => a.name.toLowerCase().includes(needle) || a.email.toLowerCase().includes(needle))
    );
    if (events.length < beforeCount) {
      progress(`🔍 Filtered ${beforeCount} → ${events.length} events matching "${search}"`);
    }
  }

  // Warn if we hit the $top limit — results may be truncated
  const rawCount = (data.value ?? []).length;
  if (rawCount >= top) {
    progress(`⚠️ Returned ${rawCount} events (hit limit) — narrow the date range or add a search filter for complete results`);
  }

  progress(`✅ Found ${events.length} calendar event(s)`);
  return events;
}

// ---------------------------------------------------------------------------
// Email / messages
// ---------------------------------------------------------------------------

export interface EmailMessage {
  id: string;
  subject: string;
  from: string;
  fromAddress: string;
  receivedDateTime: string;
  bodyPreview: string;
  body?: string;       // full body (plain text, stripped from HTML) — only when full_body requested
  isRead: boolean;
  hasAttachments: boolean;
}

export async function getEmails(opts: {
  search?: string;
  folder?: string;
  top?: number;
  unreadOnly?: boolean;
  fullBody?: boolean;
}, progress: ProgressFn = () => {}): Promise<EmailMessage[]> {
  const { search, folder: rawFolder = "inbox", top = 25, unreadOnly, fullBody } = opts;
  progress("📧 Fetching emails...");

  // Resolve custom folder names (e.g. "SITA", "PMI") to Graph folder IDs
  const folder = await resolveMailFolder(rawFolder, progress);

  const selectFields = fullBody
    ? "Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments,Body"
    : "Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments";

  let path: string;
  if (search) {
    const p = new URLSearchParams({
      $search: `"${search.replace(/"/g, "")}"`,
      $select: selectFields,
      $top: String(top),
    });
    // Search within the specified folder, not across all mail
    path = `/mailfolders/${folder}/messages?${p}`;
  } else {
    const filters: string[] = [];
    if (unreadOnly) filters.push("IsRead eq false");
    const p = new URLSearchParams({
      $select: selectFields,
      $top: String(top),
      $orderby: "ReceivedDateTime desc",
      ...(filters.length ? { $filter: filters.join(" and ") } : {}),
    });
    path = `/mailfolders/${folder}/messages?${p}`;
  }

  const token = await acquireGraphToken(progress);
  const data = await outlookApiFetch(path, token, progress);
  const messages = (data.value as Record<string, unknown>[] ?? []).map(m => {
    const fromEA = (m.From as { EmailAddress: { Name: string; Address: string } })?.EmailAddress;
    const bodyContent = (m.Body as { Content?: string; ContentType?: string } | undefined);
    const bodyText = bodyContent?.Content
      ? (bodyContent.ContentType === "html" || bodyContent.ContentType === "HTML"
          ? stripHtml(bodyContent.Content)
          : bodyContent.Content)
      : undefined;
    return {
      id:               m.Id as string,
      subject:          m.Subject as string,
      from:             fromEA?.Name || "",
      fromAddress:      fromEA?.Address || "",
      receivedDateTime: m.ReceivedDateTime as string,
      bodyPreview:      m.BodyPreview as string,
      ...(bodyText !== undefined ? { body: bodyText } : {}),
      isRead:           m.IsRead as boolean,
      hasAttachments:   m.HasAttachments as boolean,
    };
  });

  progress(`✅ Found ${messages.length} message(s)`);
  return messages;
}

// ---------------------------------------------------------------------------
// Mail folders — list all folders so users can browse custom client folders
// ---------------------------------------------------------------------------

export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId: string;
  childFolderCount: number;
  totalItemCount: number;
  unreadItemCount: number;
}

export async function listMailFolders(progress: ProgressFn = () => {}): Promise<MailFolder[]> {
  progress("📁 Fetching mail folders...");
  const token = await acquireGraphToken(progress);
  const data = await outlookApiFetch(
    "/mailfolders?$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount&$top=100",
    token, progress
  );
  const folders = (data.value as Record<string, unknown>[] ?? []).map(f => ({
    id:               f.id as string,
    displayName:      f.displayName as string,
    parentFolderId:   f.parentFolderId as string,
    childFolderCount: f.childFolderCount as number,
    totalItemCount:   f.totalItemCount as number,
    unreadItemCount:  f.unreadItemCount as number,
  }));

  // Also fetch child folders for any folder that has children
  const withChildren = folders.filter(f => f.childFolderCount > 0);
  for (const parent of withChildren) {
    try {
      const childData = await outlookApiFetch(
        `/mailfolders/${parent.id}/childfolders?$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount&$top=100`,
        token, progress
      );
      const children = (childData.value as Record<string, unknown>[] ?? []).map(f => ({
        id:               f.id as string,
        displayName:      f.displayName as string,
        parentFolderId:   f.parentFolderId as string,
        childFolderCount: f.childFolderCount as number,
        totalItemCount:   f.totalItemCount as number,
        unreadItemCount:  f.unreadItemCount as number,
      }));
      folders.push(...children);
    } catch (e) { process.stderr.write(`[alfred:warn] child folder fetch failed for ${parent.id}: ${e instanceof Error ? e.message : String(e)}\n`); }
  }

  progress(`✅ Found ${folders.length} mail folder(s)`);
  return folders;
}

/** Resolve a folder name to its Graph ID. Tries well-known names first, then searches user folders. */
export async function resolveMailFolder(folder: string, progress: ProgressFn = () => {}): Promise<string> {
  const wellKnown = ["inbox", "sentitems", "drafts", "deleteditems", "junkemail", "archive"];
  if (wellKnown.includes(folder.toLowerCase())) return folder.toLowerCase();

  // Search by display name (case-insensitive)
  const folders = await listMailFolders(progress);
  const match = folders.find(f => f.displayName.toLowerCase() === folder.toLowerCase());
  if (match) return match.id;

  // Partial match fallback
  const partial = folders.find(f => f.displayName.toLowerCase().includes(folder.toLowerCase()));
  if (partial) {
    progress(`📁 Matched folder: "${partial.displayName}"`);
    return partial.id;
  }

  // Return as-is — let Graph API error if invalid
  return folder;
}

