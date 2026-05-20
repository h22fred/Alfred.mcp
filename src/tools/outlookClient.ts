import { execFileSync } from "child_process";
import { dirname, join } from "path";
import { fileURLToPath } from "url";
import { getAlfredPages, scheduleIdleClose } from "../auth/tokenExtractor.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";
import { loadCachedAuth, saveCachedAuth, clearCachedAuthFile } from "../auth/authFileCache.js";
import { stripHtml, urlHostMatches } from "../shared.js";
import { sanitizeODataSearch } from "./dynamicsClient.js";

/** All known Outlook web domains — new ones can be added here. */
const OUTLOOK_DOMAINS = ["outlook.cloud.microsoft", "outlook.cloud.microsoft.com", "outlook.office.com", "outlook.office365.com", "outlook.live.com"] as const;

/** Match any known Outlook domain (legacy + new). */
function isOutlookUrl(url: string): boolean {
  return OUTLOOK_DOMAINS.some(d => urlHostMatches(url, d));
}

let _updateHintCache: string | null = null;
function checkForAlfredUpdate(): string {
  if (_updateHintCache !== null) return _updateHintCache;
  try {
    const __fn = fileURLToPath(import.meta.url);
    const installDir = join(dirname(__fn), "..", "..");
    const localSha = execFileSync("git", ["-C", installDir, "rev-parse", "--short", "HEAD"], { encoding: "utf8", timeout: 5_000 }).trim();
    const remoteSha = execFileSync("git", ["-C", installDir, "rev-parse", "--short", "origin/main"], { encoding: "utf8", timeout: 3_000 }).trim();
    if (localSha !== remoteSha) {
      const behind = execFileSync("git", ["-C", installDir, "rev-list", "--count", `${localSha}..origin/main`], { encoding: "utf8", timeout: 3_000 }).trim();
      _updateHintCache = `\n\n💡 An Alfred update is available (${behind} commit(s) behind). Ask Claude to run "update_alfred" — this may fix the issue.`;
    } else {
      _updateHintCache = "";
    }
  } catch { _updateHintCache = ""; }
  return _updateHintCache;
}

/**
 * Detected Outlook origin from the browser tab (set during token acquisition).
 * E.g. "https://outlook.cloud.microsoft.com" or "https://outlook.office.com".
 */
let detectedOutlookOrigin: string | null = null;
/** Decode the audience (aud) claim from a JWT without verifying signature. */
function decodeJwtAudience(jwt: string): string {
  try {
    const payload = JSON.parse(Buffer.from(jwt.split(".")[1]!, "base64url").toString());
    return (typeof payload.aud === "string" ? payload.aud : "").toLowerCase();
  } catch { return ""; }
}

const TOKEN_CACHE_FALLBACK_MS = 45 * 60 * 1000; // 45 min — fallback when MSAL doesn't report expiry
const TOKEN_REFRESH_MARGIN_MS = 5 * 60 * 1000;  // refresh 5 min before expiry

interface TokenCache {
  token: string;
  expiresAt: number;
  aud?: string;  // JWT audience — determines which API endpoint to use
}
let tokenCache: TokenCache | null = null;
let outlookRestTokenCache: TokenCache | null = null;

export function clearGraphTokenCache(): void {
  tokenCache = null;
  clearCachedAuthFile("graphToken");
}

export function clearOutlookRestTokenCache(): void {
  outlookRestTokenCache = null;
  clearCachedAuthFile("outlookRestToken");
}

/** @internal — test-only helper to pre-seed the Graph token cache */
export function _seedGraphTokenCache(token: string, ttlMs = TOKEN_CACHE_FALLBACK_MS): void {
  tokenCache = { token, expiresAt: Date.now() + ttlMs };
}

/** @internal — test-only helper to pre-seed the Outlook REST token cache */
export function _seedOutlookRestTokenCache(token: string, ttlMs = TOKEN_CACHE_FALLBACK_MS): void {
  outlookRestTokenCache = { token, expiresAt: Date.now() + ttlMs };
}

// ---------------------------------------------------------------------------
// Outlook REST token — for email/folder operations against outlook.office.com
// Separate from Graph token (which is for calendar via graph.microsoft.com)
// ---------------------------------------------------------------------------

const OUTLOOK_REST_MSAL_JS = `(function() {
  const decodeAud = (jwt) => {
    try {
      const payload = JSON.parse(atob(jwt.split('.')[1].replace(/-/g,'+').replace(/_/g,'/')));
      return (payload.aud || '').toLowerCase();
    } catch { return ''; }
  };
  const isMailAud = (aud) =>
    aud.includes('outlook.office') || aud.includes('outlook.cloud.microsoft') ||
    aud.includes('outlook.microsoft') || aud.includes('substrate.office') ||
    aud.includes('graph.microsoft.com');
  const isJwt = (s) => typeof s === 'string' && s.length > 100 && s.split('.').length === 3;
  const tryStorage = (s) => {
    try {
      for (const key of Object.keys(s)) {
        try {
          const raw = s.getItem(key);
          if (!raw || raw.length < 50) continue;
          const val = JSON.parse(raw);
          if (!val || typeof val !== 'object') continue;
          const secret = val.secret || val.access_token || val.accessToken;
          const credType = (val.credentialType || val.credential_type || '').toLowerCase();
          if (credType.includes('accesstoken') && secret && isJwt(secret)) {
            const aud = decodeAud(secret);
            if (!isMailAud(aud)) continue;
            const exp = Number(val.expiresOn || val.expires_on || val.extended_expires_on || 0);
            if (exp && exp * 1000 < Date.now()) continue;
            return { token: secret, expiresAt: exp ? exp * 1000 : 0, aud };
          }
        } catch {}
      }
    } catch {}
    return null;
  };
  return tryStorage(sessionStorage) || tryStorage(localStorage);
})()`;

// ---------------------------------------------------------------------------
// Pre-flight connection check — called before every Outlook tool operation
// ---------------------------------------------------------------------------

/**
 * Verify that the Alfred browser is running and an Outlook tab is active.
 * Throws a clear, actionable error if the connection is not healthy.
 */
async function verifyOutlookConnection(progress: ProgressFn): Promise<void> {
  if (outlookRestTokenCache && Date.now() < outlookRestTokenCache.expiresAt - TOKEN_REFRESH_MARGIN_MS) return;

  const updateHint = checkForAlfredUpdate();

  let pages: Awaited<ReturnType<typeof getAlfredPages>>;
  try {
    pages = await getAlfredPages();
  } catch {
    throw new Error(
      "Cannot connect to Alfred — the Alfred browser is not running.\n" +
      "Please launch Alfred from your Desktop and make sure it's running." + updateHint
    );
  }

  const pageSummary = pages.map(p => p.url().slice(0, 80)).join("; ");
  process.stderr.write(`[alfred] Pages (${pages.length}): ${pageSummary}\n`);

  const outlookPage = pages.find(p => isOutlookUrl(p.url()));
  if (!outlookPage) {
    throw new Error(
      "Alfred is running but no Outlook tab was found.\n" +
      `Pages: ${pageSummary || "(none)"}\n` +
      "Please open Outlook (outlook.cloud.microsoft.com) in the Alfred window and log in." + updateHint
    );
  }

  const tabUrl = outlookPage.url();
  if (urlHostMatches(tabUrl, "login.microsoftonline.com") || urlHostMatches(tabUrl, "login.live.com") || tabUrl.includes("/oauth2/")) {
    throw new Error(
      "Outlook session has expired — the Alfred browser is on the Microsoft login page.\n" +
      "Please log back into Outlook in the Alfred window (make sure the inbox loads fully), then retry." + updateHint
    );
  }

  progress("✅ Alfred connection verified — Outlook tab is active");
}

async function acquireOutlookRestToken(progress: ProgressFn): Promise<string> {
  // Cache checks (unchanged from original — keep exactly)
  if (outlookRestTokenCache && Date.now() < outlookRestTokenCache.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    const mins = Math.round((outlookRestTokenCache.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Outlook REST token (~${mins} min remaining)`);
    return outlookRestTokenCache.token;
  }
  const fileCached = loadCachedAuth("outlookRestToken");
  if (fileCached && Date.now() < fileCached.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    const aud = decodeJwtAudience(fileCached.value);
    outlookRestTokenCache = { token: fileCached.value, expiresAt: fileCached.expiresAt, aud };
    const mins = Math.round((fileCached.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Outlook REST token (~${mins} min remaining)`);
    return fileCached.value;
  }

  progress("🔑 Acquiring Outlook REST token via Playwright...");

  const pages = await getAlfredPages();

  // Detect which Outlook domain is open
  if (!detectedOutlookOrigin) {
    const outlookPageUrl = pages.find(p => isOutlookUrl(p.url()))?.url();
    if (outlookPageUrl) {
      try {
        detectedOutlookOrigin = new URL(outlookPageUrl).origin;
        process.stderr.write(`[alfred] Detected Outlook origin: ${detectedOutlookOrigin}\n`);
      } catch (e) {
        process.stderr.write(`[alfred:warn] Could not parse Outlook page URL: ${e instanceof Error ? e.message : String(e)}\n`);
      }
    }
  }

  const outlookPage = pages.find(p => isOutlookUrl(p.url()) && !urlHostMatches(p.url(), "login.microsoftonline.com"))
    ?? pages.find(p => !p.url().startsWith("about:"));

  if (!outlookPage) {
    throw new Error("No browser tabs found. Launch Alfred from your Desktop and log into Outlook.");
  }

  // Login page check
  const tabUrl = outlookPage.url();
  if (urlHostMatches(tabUrl, "login.microsoftonline.com") || urlHostMatches(tabUrl, "login.live.com") || tabUrl.includes("/oauth2/")) {
    throw new Error(
      "Outlook tab is on the Microsoft login page — your session has expired.\n" +
      "Please log back into Outlook in the Alfred window (make sure the inbox is fully loaded), then retry."
    );
  }

  // Strategy 1: MSAL cache extraction via page.evaluate()
  const msalResult = await outlookPage.evaluate(OUTLOOK_REST_MSAL_JS)
    .then(v => v as { token: string; expiresAt: number; aud?: string } | null)
    .catch((e) => {
      process.stderr.write(`[alfred:warn] Outlook REST MSAL extraction failed: ${e instanceof Error ? e.message : String(e)}\n`);
      return null;
    });

  if (msalResult?.token) {
    const expiresAt = msalResult.expiresAt > 0 ? msalResult.expiresAt : Date.now() + TOKEN_CACHE_FALLBACK_MS;
    outlookRestTokenCache = { token: msalResult.token, expiresAt, aud: msalResult.aud };
    saveCachedAuth("outlookRestToken", msalResult.token, expiresAt);
    const mins = Math.round((expiresAt - Date.now()) / 60_000);
    progress(`✅ Outlook REST token acquired from MSAL cache (~${mins} min valid, audience: ${msalResult.aud ?? "unknown"})`);
    scheduleIdleClose(3_000);
    return msalResult.token;
  }

  // Strategy 2: network interception — reload page and capture Bearer token
  progress("🔑 MSAL cache unreadable — capturing token via network interception...");

  // Use a ref object — TypeScript 5.9 narrows local variables assigned in async closures
  // to `never` at the outer scope; object properties escape that narrowing.
  const capture: { result: { token: string; requestUrl: string } | null } = { result: null };
  let found = false;

  const routeHandler = async (route: import("playwright").Route) => {
    try {
      if (!found) {
        const auth = route.request().headers()["authorization"] ?? "";
        const reqUrl = route.request().url();
        if (auth.startsWith("Bearer ")) {
          const candidateToken = auth.slice(7);
          const candidateAud = decodeJwtAudience(candidateToken);
          if (candidateAud.includes("outlook.office") || candidateAud.includes("outlook.cloud.microsoft") || candidateAud.includes("outlook.microsoft")) {
            let scp = "";
            try { scp = JSON.parse(Buffer.from(candidateToken.split(".")[1]!, "base64url").toString()).scp ?? ""; } catch {}
            if (scp.includes("Mail") || scp.includes("Calendar")) {
              found = true;
              capture.result = { token: candidateToken, requestUrl: reqUrl };
            } else {
              process.stderr.write(`[alfred:cdp] Skipping Outlook token with limited scopes "${scp.slice(0, 40)}" from ${reqUrl.slice(0, 60)}\n`);
            }
          } else {
            process.stderr.write(`[alfred:cdp] Skipping token with audience "${candidateAud}" (not Outlook) from ${reqUrl.slice(0, 60)}\n`);
          }
        }
      }
    } catch { /* swallow */ }
    await route.continue().catch(() => {});
  };

  await outlookPage.route("**", routeHandler);

  await outlookPage.reload({ waitUntil: "domcontentloaded", timeout: 15_000 }).catch(() => {});

  if (!found) {
    const now = new Date().toISOString();
    const tomorrow = new Date(Date.now() + 86400000).toISOString();
    await outlookPage.evaluate(
      `fetch('/owa/service.svc?action=GetFolder',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'}}).catch(()=>{});` +
      `fetch('/api/v2.0/me/calendarview?$top=1&$select=Id&startDateTime=${now}&endDateTime=${tomorrow}',{credentials:'include'}).catch(()=>{});` +
      `fetch('https://graph.microsoft.com/v1.0/me/messages?$top=1',{credentials:'include'}).catch(()=>{});`
    ).catch(() => {});
  }

  const deadline = Date.now() + 7_000;
  while (!capture.result && Date.now() < deadline) {
    await new Promise(r => setTimeout(r, 300));
  }

  await outlookPage.unroute("**", routeHandler).catch(() => {});

  if (capture.result) {
    const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
    const capturedAud = decodeJwtAudience(capture.result.token) || (() => { try { return new URL(capture.result!.requestUrl).origin.toLowerCase(); } catch { return ""; } })();
    outlookRestTokenCache = { token: capture.result.token, expiresAt, aud: capturedAud };
    saveCachedAuth("outlookRestToken", capture.result.token, expiresAt);
    progress(`✅ Outlook REST token captured via network interception (${capturedAud || "unknown origin"})`);
    scheduleIdleClose(3_000);
    return capture.result.token;
  }

  throw new Error(
    "Could not capture Outlook token.\n" +
    "Make sure Outlook is fully loaded (inbox visible) in the Alfred window, then retry." + checkForAlfredUpdate()
  );
}

// Resolve the correct mail API base URL based on token audience.
// New Outlook uses Graph API (graph.microsoft.com), old uses REST v2.0 (outlook.office.com).
function getOutlookApiBase(): string {
  // `aud` is the JWT audience claim (e.g. "https://graph.microsoft.com"), NOT a user-supplied URL.
  // The substring checks below select which Microsoft API base URL to use — not URL validation.
  // lgtm[js/incomplete-url-substring-sanitization]
  const aud = outlookRestTokenCache?.aud ?? "";
  // Graph API token — new Outlook uses this
  if (aud.includes("graph.microsoft.com")) return "https://graph.microsoft.com/v1.0/me";
  // Outlook REST API variants
  if (aud.includes("outlook.office365")) return "https://outlook.office365.com/api/v2.0/me";
  if (aud.includes("outlook.cloud.microsoft")) return "https://outlook.cloud.microsoft.com/api/v2.0/me";
  if (aud.includes("outlook.microsoft")) return "https://outlook.microsoft.com/api/v2.0/me";
  if (aud.includes("outlook.office.com")) return "https://outlook.office.com/api/v2.0/me";
  // If we detected the browser origin, use that
  if (detectedOutlookOrigin) return `${detectedOutlookOrigin}/api/v2.0/me`;
  // Fallback — try Graph API first (new Outlook default)
  return "https://graph.microsoft.com/v1.0/me";
}

// ---------------------------------------------------------------------------
// Outlook REST v2 fetch using Outlook REST Bearer token
// ---------------------------------------------------------------------------

async function outlookApiFetch(path: string, token: string, progress?: ProgressFn, _retryCount = 0): Promise<Record<string, unknown>> {
  const res = await fetch(`${getOutlookApiBase()}${path}`, {
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
    outlookRestTokenCache = null;
    clearCachedAuthFile("outlookRestToken");
    progress?.("🔄 Outlook REST token expired — re-acquiring...");
    const freshToken = await acquireOutlookRestToken(progress ?? (() => {}));
    const retry = await fetch(`${getOutlookApiBase()}${path}`, {
      headers: { Authorization: `Bearer ${freshToken}`, Accept: "application/json" },
      signal: AbortSignal.timeout(30_000),
    });
    if (!retry.ok) {
      // Clear caches again so next call starts fresh
      outlookRestTokenCache = null;
      clearCachedAuthFile("outlookRestToken");
      if (retry.status === 401) {
        const updateHint = checkForAlfredUpdate();
        throw new Error(
          "Outlook session has expired — re-acquired token was also rejected (401).\n" +
          "Please go to the Alfred window and log back into Outlook (make sure the inbox loads fully), then retry." +
          updateHint
        );
      }
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
  categories?: string[];
}

export async function getCalendarEvents(
  startDate: string,
  endDate: string,
  search?: string,
  progress: ProgressFn = () => {},
  top: number = 100,
  categories?: string[],
): Promise<CalendarEvent[]> {
  await verifyOutlookConnection(progress);
  progress(`📅 Fetching calendar events ${startDate} → ${endDate}${search ? ` (filter: "${search}")` : ""}...`);

  // NOTE: /calendarview does NOT support $search — we filter client-side below.
  // Use same proven token path as emails — routes to Graph API or Outlook REST depending on token audience.
  // Use camelCase field names — works with both Graph API (canonical) and Outlook REST (case-insensitive).
  const params = new URLSearchParams({
    startDateTime: `${startDate}T00:00:00Z`,
    endDateTime:   `${endDate}T23:59:59Z`,
    $select: "subject,start,end,location,organizer,attendees,isOnlineMeeting,webLink,categories",
    $top: String(top),
    $orderby: "start/dateTime",
  });

  if (search) {
    try {
      const safe = sanitizeODataSearch(search);
      params.set("$filter", `contains(subject,'${safe}')`);
    } catch (e) {
      process.stderr.write(`[alfred:warn] OData search sanitize failed, using client-side filter: ${e instanceof Error ? e.message : String(e)}\n`);
    }
  }

  const token = await acquireOutlookRestToken(progress);

  let data: { value?: Record<string, unknown>[] };
  try {
    data = await outlookApiFetch(`/calendarview?${params}`, token, progress) as { value?: Record<string, unknown>[] };
  } catch (e) {
    // If $filter is not supported (400), retry without it
    if (search && params.has("$filter") && e instanceof Error && (e.message.includes("400") || e.message.includes("UnsupportedQuery"))) {
      progress("⚠️ Server-side filter not supported on calendarView — falling back to client-side filter...");
      params.delete("$filter");
      data = await outlookApiFetch(`/calendarview?${params}`, token, progress) as { value?: Record<string, unknown>[] };
    } else {
      throw e;
    }
  }

  // Handle both PascalCase (Outlook REST) and camelCase (Graph) response fields
  let events = (data.value ?? []).map(e => {
    const orgRaw = (e.Organizer ?? e.organizer) as { EmailAddress?: { Name: string; Address: string }; emailAddress?: { name: string; address: string } } | undefined;
    const org = orgRaw?.EmailAddress
      ? { name: orgRaw.EmailAddress.Name, address: orgRaw.EmailAddress.Address }
      : orgRaw?.emailAddress
      ? { name: orgRaw.emailAddress.name, address: orgRaw.emailAddress.address }
      : undefined;
    const rawAttendees = ((e.Attendees ?? e.attendees) as Array<{ EmailAddress?: { Name: string; Address: string }; emailAddress?: { name: string; address: string } }>) ?? [];
    const attendees = rawAttendees.map(a => ({
      name:  a.EmailAddress?.Name  ?? a.emailAddress?.name  ?? "",
      email: a.EmailAddress?.Address ?? a.emailAddress?.address ?? "",
    }));
    const startField = (e.Start ?? e.start) as { DateTime?: string; dateTime?: string } | undefined;
    const endField   = (e.End ?? e.end) as { DateTime?: string; dateTime?: string } | undefined;
    const locationField = (e.Location ?? e.location) as { DisplayName?: string; displayName?: string } | undefined;
    const rawCategories = (e.Categories ?? e.categories) as string[] | undefined;
    return {
      id:              "",
      subject:         ((e.Subject ?? e.subject) as string) || "",
      start:           startField?.DateTime ?? startField?.dateTime ?? "",
      end:             endField?.DateTime ?? endField?.dateTime ?? "",
      location:        locationField?.DisplayName ?? locationField?.displayName ?? undefined,
      organizer:       org?.name || undefined,
      organizerEmail:  org?.address || undefined,
      attendees,
      isOnlineMeeting: (e.IsOnlineMeeting ?? e.isOnlineMeeting) as boolean,
      bodyPreview:     undefined,
      webLink:         ((e.WebLink ?? e.webLink) as string) || undefined,
      categories:      Array.isArray(rawCategories) && rawCategories.length > 0 ? rawCategories : undefined,
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

  // Client-side categories filter (Graph calendarView doesn't support $filter on categories)
  if (categories && categories.length > 0) {
    const needles = categories.map(c => c.toLowerCase());
    const beforeCount = events.length;
    events = events.filter(e =>
      e.categories && e.categories.some(c => needles.includes(c.toLowerCase()))
    );
    progress(`🏷️ Filtered by categories [${categories.join(", ")}]: ${beforeCount} → ${events.length} events`);
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
  const { search, folder: rawFolder, top = 25, unreadOnly, fullBody } = opts;
  await verifyOutlookConnection(progress);
  progress("📧 Fetching emails...");

  // When searching with no folder specified, search ALL mail (not just inbox).
  // When browsing (no search), default to inbox.
  const hasExplicitFolder = rawFolder !== undefined && rawFolder !== "";
  const folder = hasExplicitFolder
    ? await resolveMailFolder(rawFolder, progress)
    : null;

  // Use camelCase — works with both Graph API and Outlook REST (case-insensitive)
  const selectFields = fullBody
    ? "id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments,body"
    : "id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments";

  // Build folder path prefix: empty string = all mail, "/mailfolders/{id}" = specific folder
  // Outlook REST folder IDs are API-generated (base64-like) — do NOT encodeURIComponent
  // as it turns '=' into '%3D' which causes ErrorInvalidIdMalformed.
  const folderPrefix = folder ? `/mailfolders/${folder}` : "";

  let path: string;
  if (search) {
    const p = new URLSearchParams({
      $search: `"${search.replace(/"/g, "")}"`,
      $select: selectFields,
      $top: String(top),
    });
    // No folder = search across ALL folders (inbox, sent, custom client folders, etc.)
    path = `${folderPrefix}/messages?${p}`;
  } else {
    // Browsing without search — default to inbox if no folder specified
    const browsePrefix = folderPrefix || "/mailfolders/inbox";
    const filters: string[] = [];
    if (unreadOnly) filters.push("isRead eq false");
    const p = new URLSearchParams({
      $select: selectFields,
      $top: String(top),
      $orderby: "receivedDateTime desc",
      ...(filters.length ? { $filter: filters.join(" and ") } : {}),
    });
    path = `${browsePrefix}/messages?${p}`;
  }

  const token = await acquireOutlookRestToken(progress);
  const data = await outlookApiFetch(path, token, progress);
  // Handle both PascalCase (Outlook REST) and camelCase (Graph API) response fields
  const messages = (data.value as Record<string, unknown>[] ?? []).map(m => {
    // From field: REST = { EmailAddress: { Name, Address } }, Graph = { emailAddress: { name, address } }
    const fromRaw = (m.From ?? m.from) as { EmailAddress?: { Name: string; Address: string }; emailAddress?: { name: string; address: string } } | undefined;
    const fromEA = fromRaw?.EmailAddress
      ? { name: fromRaw.EmailAddress.Name, address: fromRaw.EmailAddress.Address }
      : fromRaw?.emailAddress
      ? { name: fromRaw.emailAddress.name, address: fromRaw.emailAddress.address }
      : undefined;
    // Body: REST = { Content, ContentType }, Graph = { content, contentType }
    const bodyRaw = (m.Body ?? m.body) as { Content?: string; ContentType?: string; content?: string; contentType?: string } | undefined;
    const bodyHtml = bodyRaw?.Content ?? bodyRaw?.content;
    const bodyType = (bodyRaw?.ContentType ?? bodyRaw?.contentType ?? "").toLowerCase();
    const bodyText = bodyHtml
      ? (bodyType === "html" ? stripHtml(bodyHtml) : bodyHtml)
      : undefined;
    return {
      id:               ((m.Id ?? m.id) as string) || "",
      subject:          ((m.Subject ?? m.subject) as string) || "",
      from:             fromEA?.name || "",
      fromAddress:      fromEA?.address || "",
      receivedDateTime: ((m.ReceivedDateTime ?? m.receivedDateTime) as string) || "",
      bodyPreview:      ((m.BodyPreview ?? m.bodyPreview) as string) || "",
      ...(bodyText !== undefined ? { body: bodyText } : {}),
      isRead:           ((m.IsRead ?? m.isRead) as boolean) ?? false,
      hasAttachments:   ((m.HasAttachments ?? m.hasAttachments) as boolean) ?? false,
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
  await verifyOutlookConnection(progress);
  progress("📁 Fetching mail folders...");
  const token = await acquireOutlookRestToken(progress);
  const SELECT = "$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount&$top=100";
  const data = await outlookApiFetch(`/mailfolders?${SELECT}`, token, progress);

  const mapFolder = (f: Record<string, unknown>): MailFolder => ({
    id:               (f.Id ?? f.id) as string,
    displayName:      (f.DisplayName ?? f.displayName) as string,
    parentFolderId:   (f.ParentFolderId ?? f.parentFolderId) as string,
    childFolderCount: (f.ChildFolderCount ?? f.childFolderCount ?? 0) as number,
    totalItemCount:   (f.TotalItemCount ?? f.totalItemCount ?? 0) as number,
    unreadItemCount:  (f.UnreadItemCount ?? f.unreadItemCount ?? 0) as number,
  });

  const folders = (data.value as Record<string, unknown>[] ?? []).map(mapFolder);

  const MAX_DEPTH = 4;
  const CHILD_BATCH = 5;
  const fetchChildren = async (parents: MailFolder[], depth: number) => {
    if (depth >= MAX_DEPTH) return;
    const withChildren = parents.filter(f => f.childFolderCount > 0);
    for (let i = 0; i < withChildren.length; i += CHILD_BATCH) {
      const batch = withChildren.slice(i, i + CHILD_BATCH);
      const results = await Promise.allSettled(
        batch.map(parent =>
          outlookApiFetch(`/mailfolders/${parent.id}/childfolders?${SELECT}`, token, progress)
            .then(data => (data.value as Record<string, unknown>[] ?? []).map(mapFolder))
        )
      );
      const allChildren: MailFolder[] = [];
      for (let j = 0; j < results.length; j++) {
        const result = results[j]!;
        if (result.status === "fulfilled") {
          folders.push(...result.value);
          allChildren.push(...result.value);
        } else {
          process.stderr.write(`[alfred:warn] child folder fetch failed for "${batch[j]!.displayName}" (${batch[j]!.id}): ${result.reason instanceof Error ? result.reason.message : String(result.reason)}\n`);
        }
      }
      await fetchChildren(allChildren, depth + 1);
    }
  };
  await fetchChildren(folders, 0);

  progress(`✅ Found ${folders.length} mail folder(s)`);
  return folders;
}

/** Normalize folder name for matching: collapse whitespace, normalize dashes (em/en dash → hyphen), lowercase */
function normFolder(s: string): string {
  return s.toLowerCase().replace(/[\u2013\u2014]/g, "-").replace(/\s+/g, " ").trim();
}

/** Resolve a folder name to its Outlook REST folder ID. Tries well-known names first, then searches user folders. */
export async function resolveMailFolder(folder: string, progress: ProgressFn = () => {}): Promise<string> {
  const wellKnown = ["inbox", "sentitems", "drafts", "deleteditems", "junkemail", "archive"];
  if (wellKnown.includes(folder.toLowerCase())) return folder.toLowerCase();

  progress(`📁 Looking up folder "${folder}"...`);
  const folders = await listMailFolders(progress);
  const needle = normFolder(folder);

  // 1. Exact match (normalized)
  const exact = folders.find(f => normFolder(f.displayName) === needle);
  if (exact) {
    progress(`📁 Exact match: "${exact.displayName}" (${exact.totalItemCount} items)`);
    return exact.id;
  }

  // 2. Partial: folder name contained in display name (e.g. "Fabrikam" matches "Fabrikam - Philip Morris")
  const partial = folders.find(f => normFolder(f.displayName).includes(needle));
  if (partial) {
    progress(`📁 Partial match: "${partial.displayName}" (${partial.totalItemCount} items)`);
    return partial.id;
  }

  // 3. Reverse partial: display name contained in folder name (e.g. "Fabrikam - Philip Morris International"
  //    matches a folder named "Fabrikam - Philip Morris")
  const reverse = folders.find(f => needle.includes(normFolder(f.displayName)));
  if (reverse) {
    progress(`📁 Reverse match: "${reverse.displayName}" (${reverse.totalItemCount} items)`);
    return reverse.id;
  }

  // 4. Token match: split on " - " and match first segment (e.g. "Fabrikam - Philip Morris" → "Fabrikam")
  const firstToken = needle.split(" - ")[0]?.trim();
  if (firstToken && firstToken !== needle) {
    const tokenMatch = folders.find(f => normFolder(f.displayName).includes(firstToken));
    if (tokenMatch) {
      progress(`📁 Token match on "${firstToken}": "${tokenMatch.displayName}" (${tokenMatch.totalItemCount} items)`);
      return tokenMatch.id;
    }
  }

  progress(`⚠️ No folder match for "${folder}" — available: ${folders.map(f => f.displayName).slice(0, 10).join(", ")}`);
  throw new Error(
    `Mail folder "${folder}" not found.\n` +
    `Available folders: ${folders.map(f => f.displayName).join(", ")}\n` +
    `Use list_mail_folders to see all available folders.`
  );
}

