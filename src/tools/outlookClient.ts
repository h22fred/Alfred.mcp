import { connectWithRetry } from "../auth/tokenExtractor.js";
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
// Acquire a Graph Bearer token by loading a lightweight Outlook page
// and intercepting its outgoing Graph API requests.
// ---------------------------------------------------------------------------

async function acquireGraphToken(progress: ProgressFn): Promise<string> {
  if (tokenCache && Date.now() < tokenCache.expiresAt) {
    const mins = Math.round((tokenCache.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Graph token (~${mins} min remaining)`);
    return tokenCache.token;
  }

  // Check Chrome is reachable
  try {
    execFileSync("curl", ["-s", "--max-time", "1", `http://localhost:${CDP_PORT}/json/version`], { timeout: 2_000 });
  } catch {
    throw new Error("ChromeLink is not running. Open ChromeLink.app first.");
  }

  const browser = await connectWithRetry();

  try {
    const ctx = browser.contexts()[0];
    if (!ctx) throw new Error("No browser context found");

    let capturedToken: string | null = null;

    // Prefer the existing Outlook tab — avoids a full page load
    const existingPage = ctx.pages().find(p => p.url().includes("outlook.office.com"));
    const page = existingPage ?? await ctx.newPage();
    const isNewPage = !existingPage;

    await page.route("**/outlook.office.com/**", async (route) => {
      const auth = route.request().headers()["authorization"] ?? "";
      if (!capturedToken && auth.startsWith("Bearer ")) {
        capturedToken = auth.slice(7);
      }
      await route.continue();
    });

    if (isNewPage) {
      // No existing tab — load Outlook fresh (slow path, only happens once per session)
      progress("📧 Loading Outlook to capture auth token...");
      await page.goto(`${OUTLOOK_ORIGIN}/mail/`, { waitUntil: "domcontentloaded", timeout: 30_000 });
    } else {
      // Existing tab — trigger a lightweight API call to fire auth headers immediately
      progress("📧 Capturing auth token from existing Outlook tab...");
      await page.evaluate(() =>
        fetch("/owa/service.svc?action=GetAccessTokenforResource", { method: "POST", body: "{}", credentials: "include" }).catch(() => {})
      );
    }

    // Wait up to 5 s for token (existing tab fires immediately; new tab may take longer)
    const deadline = Date.now() + (isNewPage ? 10_000 : 5_000);
    while (!capturedToken && Date.now() < deadline) {
      await page.waitForTimeout(200);
    }

    await page.unroute("**/outlook.office.com/**");
    if (isNewPage) await page.close();

    if (!capturedToken) {
      throw new Error(
        "Could not capture Graph token from Outlook.\n" +
        "Make sure you are logged into https://outlook.office.com in ChromeLink and the page is fully loaded."
      );
    }

    tokenCache = { token: capturedToken, expiresAt: Date.now() + TOKEN_CACHE_MS };
    progress("✅ Graph token acquired");
    return capturedToken;
  } finally {
    await browser.close();
  }
}

const OUTLOOK_API = "https://outlook.office.com/api/v2.0/me";

// ---------------------------------------------------------------------------
// Outlook REST v2 fetch using the captured Bearer token
// ---------------------------------------------------------------------------

async function outlookApiFetch(path: string, token: string, progress?: ProgressFn): Promise<Record<string, unknown>> {
  const res = await fetch(`${OUTLOOK_API}${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
  });

  if (res.status === 401) {
    // Token expired — clear cache and retry once with a fresh token
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
  const token = await acquireGraphToken(progress);

  const params = new URLSearchParams({
    startDateTime: `${startDate}T00:00:00Z`,
    endDateTime:   `${endDate}T23:59:59Z`,
    $select: "Id,Subject,Start,End,Location,Organizer,IsOnlineMeeting,BodyPreview",
    $top: "200",
    $orderby: "Start/DateTime",
  });
  if (search) params.set("$search", `"${search}"`);

  const data = await outlookApiFetch(`/calendarview?${params}`, token, progress);

  const events = (data.value as Record<string, unknown>[] ?? []).map(e => ({
    id:              e.Id as string,
    subject:         e.Subject as string,
    start:           (e.Start as { DateTime: string })?.DateTime,
    end:             (e.End   as { DateTime: string })?.DateTime,
    location:        (e.Location as { DisplayName: string })?.DisplayName || undefined,
    organizer:       (e.Organizer as { EmailAddress: { Name: string } })?.EmailAddress?.Name || undefined,
    isOnlineMeeting: e.IsOnlineMeeting as boolean,
    bodyPreview:     e.BodyPreview as string | undefined,
  }));

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
  isRead: boolean;
  hasAttachments: boolean;
}

export async function getEmails(opts: {
  search?: string;
  folder?: string;
  top?: number;
  unreadOnly?: boolean;
}, progress: ProgressFn = () => {}): Promise<EmailMessage[]> {
  const { search, folder = "inbox", top = 25, unreadOnly } = opts;
  progress("📧 Fetching emails...");
  const token = await acquireGraphToken(progress);

  let url: string;
  let path: string;
  if (search) {
    const p = new URLSearchParams({
      $search: `"${search}"`,
      $select: "Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments",
      $top: String(top),
    });
    path = `/messages?${p}`;
  } else {
    const filters: string[] = [];
    if (unreadOnly) filters.push("IsRead eq false");
    const p = new URLSearchParams({
      $select: "Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments",
      $top: String(top),
      $orderby: "ReceivedDateTime desc",
      ...(filters.length ? { $filter: filters.join(" and ") } : {}),
    });
    path = `/mailfolders/${folder}/messages?${p}`;
  }

  const data = await outlookApiFetch(path, token, progress);
  const messages = (data.value as Record<string, unknown>[] ?? []).map(m => {
    const fromEA = (m.From as { EmailAddress: { Name: string; Address: string } })?.EmailAddress;
    return {
      id:               m.Id as string,
      subject:          m.Subject as string,
      from:             fromEA?.Name || "",
      fromAddress:      fromEA?.Address || "",
      receivedDateTime: m.ReceivedDateTime as string,
      bodyPreview:      m.BodyPreview as string,
      isRead:           m.IsRead as boolean,
      hasAttachments:   m.HasAttachments as boolean,
    };
  });

  progress(`✅ Found ${messages.length} message(s)`);
  return messages;
}

// no-op — kept so index.ts import compiles
export function setOutlookCookies(_cookie: string): void { /* unused */ }
