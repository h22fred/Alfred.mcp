import { connectWithRetry, getOutlookCookies, clearAuthCache } from "../auth/tokenExtractor.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";
import { execFileSync } from "child_process";

const CDP_PORT = 9222;
const OUTLOOK_ORIGIN = "https://outlook.office.com";
const TOKEN_CACHE_MS  = 45 * 60 * 1000; // 45 min

interface TokenCache {
  token: string;
  expiresAt: number;
}
let tokenCache: TokenCache | null = null;

export function clearGraphTokenCache(): void {
  tokenCache = null;
}

// ---------------------------------------------------------------------------
// Acquire a Graph Bearer token via raw CDP WebSocket — no Playwright needed.
// Enables Network tracking on the Outlook tab, triggers a lightweight OWA
// service call, and captures the outgoing Authorization header.
// ---------------------------------------------------------------------------

// JS snippet injected into the Outlook page to extract a Bearer token from MSAL cache
const MSAL_EXTRACT_JS = `(function() {
  const tryStorage = (s) => {
    try {
      for (const key of Object.keys(s)) {
        try {
          const val = JSON.parse(s.getItem(key));
          if (val && val.credentialType === 'AccessToken' && val.secret &&
              (val.target || '').toLowerCase().includes('mail')) {
            const exp = Number(val.expiresOn || val.extended_expires_on || 0);
            if (!exp || exp * 1000 > Date.now()) return val.secret;
          }
        } catch {}
      }
    } catch {}
    return null;
  };
  return tryStorage(sessionStorage) || tryStorage(localStorage);
})()`;

async function acquireGraphTokenRawCDP(progress: ProgressFn): Promise<string> {
  if (tokenCache && Date.now() < tokenCache.expiresAt) {
    const mins = Math.round((tokenCache.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Graph token (~${mins} min remaining)`);
    return tokenCache.token;
  }

  progress("🔑 Acquiring Graph Bearer token via CDP...");

  const listRes = await fetch(`http://localhost:${CDP_PORT}/json/list`);
  const targets = await listRes.json() as Array<{ webSocketDebuggerUrl?: string; type?: string; url?: string }>;

  // Prefer Outlook tab, fall back to any page
  const outlookTarget = targets.find(t => t.type === "page" && t.url?.includes("outlook.office.com") && t.webSocketDebuggerUrl);
  const anyTarget     = targets.find(t => t.type === "page" && t.webSocketDebuggerUrl);
  const target        = outlookTarget ?? anyTarget;

  if (!target?.webSocketDebuggerUrl) {
    throw new Error("No page targets found in Alfred. Open Alfred.app and log into Outlook.");
  }

  if (!outlookTarget) {
    throw new Error(
      "No Outlook tab found in Alfred.\n" +
      "Please open https://outlook.office.com in the Alfred window and log in, then retry."
    );
  }

  // Step 1: try reading token directly from MSAL storage — fast, no network interception
  const msalToken = await new Promise<string | null>((resolve) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    const timer = setTimeout(() => { try { ws.close(); } catch {} resolve(null); }, 5_000);

    ws.addEventListener("open", () => {
      ws.send(JSON.stringify({ id: 1, method: "Runtime.evaluate", params: { expression: MSAL_EXTRACT_JS, returnByValue: true } }));
    });
    ws.addEventListener("message", (event: MessageEvent) => {
      clearTimeout(timer);
      try { ws.close(); } catch {}
      try {
        const msg = JSON.parse(event.data as string) as { result?: { result?: { value?: string } } };
        resolve(msg.result?.result?.value ?? null);
      } catch { resolve(null); }
    });
    ws.addEventListener("error", () => { clearTimeout(timer); try { ws.close(); } catch {} resolve(null); });
  });

  if (msalToken) {
    tokenCache = { token: msalToken, expiresAt: Date.now() + TOKEN_CACHE_MS };
    progress("✅ Graph token acquired from MSAL cache");
    return msalToken;
  }

  // Step 2: fallback — enable Network tracking and trigger an OWA API fetch
  progress("🔑 MSAL cache miss — capturing token via network interception...");

  return new Promise((resolve, reject) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    let capturedToken: string | null = null;

    const timer = setTimeout(() => {
      try { ws.close(); } catch {}
      reject(new Error(
        "Could not capture Graph token from Outlook.\n" +
        "Make sure you are logged into https://outlook.office.com in Alfred."
      ));
    }, 10_000);

    const done = (token: string) => {
      clearTimeout(timer);
      try { ws.close(); } catch {}
      tokenCache = { token, expiresAt: Date.now() + TOKEN_CACHE_MS };
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
            expression: `fetch('/api/v2.0/me/messages?$top=1&$select=Id', { credentials: 'include' }).catch(()=>{})`,
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
      } catch { /* ignore */ }
    });

    ws.addEventListener("error", () => {
      clearTimeout(timer);
      reject(new Error("CDP WebSocket error — is Alfred running?"));
    });
  });
}

// acquireGraphToken — uses raw CDP (no Playwright) for calendar/email.
// Playwright-based token capture is kept in teamsClient.ts for Graph API calls.
async function acquireGraphToken(progress: ProgressFn): Promise<string> {
  return acquireGraphTokenRawCDP(progress);
}

const OUTLOOK_API = "https://outlook.office.com/api/v2.0/me";

// ---------------------------------------------------------------------------
// Outlook REST v2 fetch using Bearer token (kept for Teams Graph API usage)
// ---------------------------------------------------------------------------

async function outlookApiFetch(path: string, token: string, progress?: ProgressFn): Promise<Record<string, unknown>> {
  const res = await fetch(`${OUTLOOK_API}${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
  });

  if (res.status === 401) {
    tokenCache = null;
    progress?.("🔄 Graph token expired — re-acquiring...");
    const freshToken = await acquireGraphToken(progress ?? (() => {}));
    const retry = await fetch(`${OUTLOOK_API}${path}`, {
      headers: { Authorization: `Bearer ${freshToken}`, Accept: "application/json" },
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
}

export async function getCalendarEvents(
  startDate: string,
  endDate: string,
  search?: string,
  progress: ProgressFn = () => {}
): Promise<CalendarEvent[]> {
  progress(`📅 Fetching calendar events ${startDate} → ${endDate}...`);

  const params = new URLSearchParams({
    startDateTime: `${startDate}T00:00:00Z`,
    endDateTime:   `${endDate}T23:59:59Z`,
    $select: "Id,Subject,Start,End,Location,Organizer,Attendees,IsOnlineMeeting,BodyPreview",
    $top: "200",
    $orderby: "Start/DateTime",
  });
  if (search) params.set("$search", `"${search}"`);

  const token = await acquireGraphToken(progress);
  const data = await outlookApiFetch(`/calendarview?${params}`, token, progress);

  const events = (data.value as Record<string, unknown>[] ?? []).map(e => {
    const org = (e.Organizer as { EmailAddress: { Name: string; Address: string } } | undefined)?.EmailAddress;
    const rawAttendees = e.Attendees as Array<{ EmailAddress: { Name: string; Address: string }; Type: string }> ?? [];
    const attendees = rawAttendees.map(a => ({
      name:  a.EmailAddress?.Name  ?? "",
      email: a.EmailAddress?.Address ?? "",
    }));
    return {
      id:              e.Id as string,
      subject:         e.Subject as string,
      start:           (e.Start as { DateTime: string })?.DateTime,
      end:             (e.End   as { DateTime: string })?.DateTime,
      location:        (e.Location as { DisplayName: string })?.DisplayName || undefined,
      organizer:       org?.Name || undefined,
      organizerEmail:  org?.Address || undefined,
      attendees,
      isOnlineMeeting: e.IsOnlineMeeting as boolean,
      bodyPreview:     e.BodyPreview as string | undefined,
    };
  });

  progress(`✅ Found ${events.length} calendar event(s)`);
  return events;
}

// ---------------------------------------------------------------------------
// Email / messages
// ---------------------------------------------------------------------------

// Strip HTML to readable plain text
function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<\/th>/gi, " | ")
    .replace(/<\/td>/gi, " | ")
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&nbsp;/g, " ")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

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
  const { search, folder = "inbox", top = 25, unreadOnly, fullBody } = opts;
  progress("📧 Fetching emails...");

  const selectFields = fullBody
    ? "Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments,Body"
    : "Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments";

  let path: string;
  if (search) {
    const p = new URLSearchParams({
      $search: `"${search}"`,
      $select: selectFields,
      $top: String(top),
    });
    path = `/messages?${p}`;
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

// no-op — kept so index.ts import compiles
export function setOutlookCookies(_cookie: string): void { /* unused */ }
