import type { ProgressFn } from "../auth/tokenExtractor.js";
import { isAlfredgable, getAlfredContext, getAlfredPages } from "../auth/tokenExtractor.js";
import type { Page } from "playwright";
import { loadCachedAuth, saveCachedAuth, clearCachedAuthFile } from "../auth/authFileCache.js";
import { stripHtml, urlHostMatches } from "../shared.js";

/** Match all known Outlook domains (legacy + new). */
function isOutlookUrl(url: string): boolean {
  return urlHostMatches(url, "outlook.cloud.microsoft") || urlHostMatches(url, "outlook.cloud.microsoft.com") || urlHostMatches(url, "outlook.office.com") || urlHostMatches(url, "outlook.office365.com");
}
const TOKEN_CACHE_FALLBACK_MS = 45 * 60 * 1000; // fallback when MSAL doesn't report expiry
const TOKEN_REFRESH_MARGIN_MS = 5 * 60 * 1000;  // refresh 5 min before expiry

// ---------------------------------------------------------------------------
// Teams webhook config — loaded from ~/.alfred-config.json on startup,
// can be overridden at runtime via configure_teams_webhook tool
// ---------------------------------------------------------------------------

let webhookUrl: string | null = null;

function isValidWebhookUrl(url: string): boolean {
  try {
    const parsed = new URL(url);
    return parsed.protocol === "https:" &&
      parsed.hostname.endsWith(".webhook.office.com");
  } catch { return false; }
}

// Auto-load webhook from persistent config
try {
  const fs = await import("fs");
  const os = await import("os");
  const cfgPath = `${os.default.homedir()}/.alfred-config.json`;
  if (fs.default.existsSync(cfgPath)) {
    const cfg = JSON.parse(fs.default.readFileSync(cfgPath, "utf-8"));
    if (cfg.teamsWebhook) {
      if (!isValidWebhookUrl(cfg.teamsWebhook)) {
        console.error("[teams] Webhook URL in config rejected — not a valid Microsoft/Office webhook URL");
      } else {
        webhookUrl = cfg.teamsWebhook;
        console.error("[teams] Webhook URL loaded from config");
      }
    }
  }
} catch (e) { process.stderr.write(`[alfred:warn] failed to load webhook config: ${e instanceof Error ? e.message : String(e)}\n`); }

export function setTeamsWebhook(url: string): void {
  if (!isValidWebhookUrl(url)) {
    throw new Error("Invalid webhook URL. Must be an HTTPS URL on *.webhook.office.com");
  }
  webhookUrl = url;
  console.error("[teams] Webhook URL configured");
}

// ---------------------------------------------------------------------------
// Post notification to Teams channel via incoming webhook
// ---------------------------------------------------------------------------

export async function postAdaptiveCard(
  card: Record<string, unknown>,
  progress: ProgressFn = () => {}
): Promise<void> {
  if (!webhookUrl) throw new Error("Teams webhook not configured. Use configure_teams_webhook first.");
  progress("📣 Posting to Teams...");
  const payload = JSON.stringify({
    type: "message",
    attachments: [{
      contentType: "application/vnd.microsoft.card.adaptive",
      content: card,
    }],
  });
  const sizeKb = payload.length / 1024;
  progress(`📦 Card payload: ${sizeKb.toFixed(1)}KB`);
  if (sizeKb > 27) {
    progress(`⚠️  Card payload is ${sizeKb.toFixed(1)}KB — Teams limit is ~28KB, card may be silently dropped`);
  }
  const res = await fetch(webhookUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: payload,
    signal: AbortSignal.timeout(15_000),
  });
  const responseText = await res.text().catch(() => "");
  if (!res.ok) {
    throw new Error(`Teams webhook error: ${res.status} ${res.statusText}${responseText ? ` — ${responseText}` : ""}`);
  }
  // Teams returns "1" on success, anything else is a silent error
  if (responseText && responseText !== "1") {
    progress(`⚠️  Teams response: ${responseText}`);
    if (responseText.toLowerCase().includes("failed") || responseText.toLowerCase().includes("error")) {
      throw new Error(`Teams rejected the card: ${responseText}`);
    }
  }
  progress("✅ Posted to Teams");
}

export async function postTeamsNotification(
  title: string,
  body: string,
  progress: ProgressFn = () => {}
): Promise<void> {
  if (!webhookUrl) {
    throw new Error(
      "Teams webhook not configured.\n" +
      "Use the configure_teams_webhook tool to set your Teams incoming webhook URL."
    );
  }

  progress(`📣 Posting Teams notification: "${title}"...`);

  // Adaptive Card payload
  const card = {
    type: "message",
    attachments: [{
      contentType: "application/vnd.microsoft.card.adaptive",
      content: {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.4",
        body: [
          { type: "TextBlock", text: title, weight: "Bolder", size: "Large", wrap: true },
          { type: "TextBlock", text: body, wrap: true, spacing: "Medium" },
        ],
      },
    }],
  };

  const res = await fetch(webhookUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(card),
    signal: AbortSignal.timeout(15_000),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Teams webhook error: ${res.status} ${res.statusText}${text ? ` — ${text}` : ""}`);
  }

  progress("✅ Teams notification sent");
}

// ---------------------------------------------------------------------------
// Graph Bearer token acquisition
// ---------------------------------------------------------------------------

// Reads a graph.microsoft.com-scoped token from MSAL localStorage/sessionStorage
// Decodes the JWT aud claim and checks for Chat.Read scope — Teams endpoints need
// a token with chat scopes, not just any Graph token (Outlook tokens lack these).
const GRAPH_MSAL_EXTRACT_JS = `(function() {
  const decodeToken = (secret) => {
    try {
      return JSON.parse(atob(secret.split('.')[1].replace(/-/g,'+').replace(/_/g,'/')));
    } catch { return null; }
  };
  const isGraphToken = (payload) => String(payload.aud || '').includes('graph.microsoft.com');
  const hasScope = (payload, scope) => String(payload.scp || '').toLowerCase().includes(scope.toLowerCase());
  const tryStorage = (s) => {
    try {
      // Two passes: first look for tokens with Chat.Read (broad Teams scopes),
      // then fall back to any Graph token (still useful for calendar/mail).
      let bestToken = null;
      for (const key of Object.keys(s)) {
        try {
          const val = JSON.parse(s.getItem(key));
          if (val && val.credentialType === 'AccessToken' && val.secret) {
            const payload = decodeToken(val.secret);
            if (!payload || !isGraphToken(payload)) continue;
            const exp = Number(val.expiresOn || val.extended_expires_on || 0);
            if (exp && exp * 1000 < Date.now()) continue;
            if (hasScope(payload, 'Chat.Read')) return val.secret;  // best: has chat scopes
            if (!bestToken) bestToken = val.secret;                  // fallback: any graph token
          }
        } catch {}
      }
      return bestToken;
    } catch {}
    return null;
  };
  return tryStorage(sessionStorage) || tryStorage(localStorage);
})()`;

async function extractGraphTokenFromPage(page: Page): Promise<string | null> {
  return page.evaluate(GRAPH_MSAL_EXTRACT_JS)
    .then(v => v as string | null)
    .catch((e) => {
      process.stderr.write(`[alfred:warn] Teams MSAL extraction failed: ${e instanceof Error ? e.message : String(e)}\n`);
      return null;
    });
}

interface TokenCache { token: string; expiresAt: number; }
let teamsTokenCache: TokenCache | null = null;

// Skype messaging API token cache (separate from Graph token)
interface SkypeTokenCache { token: string; region: string; expiresAt: number; }
let skypeTokenCache: SkypeTokenCache | null = null;

// Teams client app ID — registered by Microsoft with Graph permissions
const TEAMS_CLIENT_ID = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
const TEAMS_REDIRECT_URI = "https://teams.microsoft.com/go";

async function acquireTokenViaSilentAuth(progress: ProgressFn): Promise<string | null> {
  const ctx = await getAlfredContext().catch(() => null);
  if (!ctx) return null;

  // Get tenant ID from Teams cookies (replaces raw CDP Network.getAllCookies)
  const allCookies = await ctx.cookies();
  const tenantCookie = allCookies.find(c =>
    c.name === "tenantId" && (c.domain === "teams.microsoft.com" || c.domain.endsWith(".teams.microsoft.com"))
  );
  const tenantId = tenantCookie?.value ?? null;

  if (!tenantId) {
    process.stderr.write("[alfred:warn] Could not extract tenant ID from Teams cookies — silent auth unavailable\n");
    return null;
  }

  progress("🔐 Opening silent Azure AD auth flow...");

  const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?` +
    `client_id=${TEAMS_CLIENT_ID}` +
    `&response_type=token` +
    `&redirect_uri=${encodeURIComponent(TEAMS_REDIRECT_URI)}` +
    `&scope=${encodeURIComponent("https://graph.microsoft.com/.default")}` +
    `&prompt=none` +
    `&response_mode=fragment`;

  const CAPTURE_HTML = `<html><body><script>var h=window.location.hash;document.title=h.includes("access_token=")?"TOKEN:"+h:"ERROR:"+h;</script></body></html>`;

  const page = await ctx.newPage();
  try {
    await page.route("https://teams.microsoft.com/go**", async (route) => {
      await route.fulfill({ status: 200, contentType: "text/html", body: CAPTURE_HTML });
    });

    await page.goto(authUrl, { waitUntil: "domcontentloaded", timeout: 15_000 }).catch(() => {});

    await page.waitForFunction(
      () => document.title.startsWith("TOKEN:") || document.title.startsWith("ERROR:"),
      { timeout: 5_000 }
    ).catch(() => {});

    let token: string | null = null;

    const currentUrl = page.url();
    if (currentUrl.includes("access_token=")) {
      const hashPart = currentUrl.includes("#") ? currentUrl.split("#")[1]! : "";
      token = new URLSearchParams(hashPart).get("access_token");
    }

    if (!token) {
      const title = await page.title();
      if (title.startsWith("TOKEN:")) {
        const hashPart = title.replace(/^TOKEN:#?/, "");
        token = new URLSearchParams(hashPart).get("access_token");
      } else if (title.startsWith("ERROR:")) {
        process.stderr.write(`[alfred:warn] Silent auth: error in redirect: ${title.slice(0, 100)}\n`);
      } else {
        process.stderr.write(`[alfred:warn] Silent auth: no token (title=${title.slice(0, 100)}, url=${currentUrl.slice(0, 100)})\n`);
      }
    }

    return token;
  } catch (e) {
    process.stderr.write(`[alfred:warn] Silent auth failed: ${e instanceof Error ? e.message : String(e)}\n`);
    return null;
  } finally {
    await page.close().catch(() => {});
  }
}

export async function acquireTeamsGraphToken(progress: ProgressFn): Promise<string> {
  if (teamsTokenCache && Date.now() < teamsTokenCache.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    const mins = Math.round((teamsTokenCache.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Teams Graph token (~${mins} min remaining)`);
    return teamsTokenCache.token;
  }

  // Check file cache before hitting CDP — apply refresh margin
  const fileCached = loadCachedAuth("teamsGraphToken");
  if (fileCached && Date.now() < fileCached.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    teamsTokenCache = { token: fileCached.value, expiresAt: fileCached.expiresAt };
    const mins = Math.round((fileCached.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Teams Graph token (~${mins} min remaining)`);
    return fileCached.value;
  }

  if (!await isAlfredgable()) {
    throw new Error("The Alfred browser is not running. Launch Alfred from your Desktop first.");
  }

  progress("🔐 Acquiring Graph token via Teams/Outlook in Alfred...");

  // Step 1: try to read Graph token directly from MSAL cache in open tabs (fast, no new page)
  const pages = await getAlfredPages();

  // Prefer Teams tab — it has broader Graph scopes (Chat.Read etc.)
  const teamsPage = pages.find(p => urlHostMatches(p.url(), "teams.microsoft.com"));
  const outlookPage = pages.find(p => isOutlookUrl(p.url()));
  const candidates = [teamsPage, outlookPage].filter(Boolean) as Page[];

  if (!teamsPage) {
    process.stderr.write("[alfred] No Teams tab found in Alfred — Graph token capture may fail. Open teams.microsoft.com/v2/ in Alfred.\n");
  }

  // Try MSAL extraction from all candidate tabs in parallel — first success wins
  for (let attempt = 0; attempt < 2; attempt++) {
    const results = await Promise.allSettled(
      candidates.map(p => extractGraphTokenFromPage(p))
    );
    const token = results
      .filter((r): r is PromiseFulfilledResult<string | null> => r.status === "fulfilled")
      .map(r => r.value)
      .find(v => v != null);
    if (token) {
      const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
      teamsTokenCache = { token, expiresAt };
      saveCachedAuth("teamsGraphToken", token, expiresAt);
      progress("✅ Graph token acquired from MSAL cache");
      return token;
    }
    if (attempt < 1) await new Promise(r => setTimeout(r, 1_000));
  }

  // Step 2: Silent Azure AD auth — open a temporary Playwright page, navigate to the
  // Azure AD authorize endpoint with prompt=none (uses existing ESTSAUTH session
  // cookie), and capture the Graph token from the redirect URL fragment.
  // This works even when Teams v2 encrypts its MSAL cache.
  progress("🔑 MSAL cache miss — acquiring Graph token via silent Azure AD auth...");
  const silentToken = await acquireTokenViaSilentAuth(progress);
  if (silentToken) {
    // Decode JWT to get real expiry
    let expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
    try {
      const payload = JSON.parse(atob(silentToken.split(".")[1]!.replace(/-/g, "+").replace(/_/g, "/")));
      if (payload.exp) expiresAt = payload.exp * 1000;
    } catch (e) {
      process.stderr.write(`[alfred:warn] Graph token JWT decode failed, using fallback expiry: ${e instanceof Error ? e.message : String(e)}\n`);
    }
    teamsTokenCache = { token: silentToken, expiresAt };
    saveCachedAuth("teamsGraphToken", silentToken, expiresAt);
    progress("✅ Graph token acquired via silent Azure AD auth");
    return silentToken;
  }

  throw new Error(
    "Could not acquire a Microsoft Graph token.\n" +
    "Make sure you are logged into Teams (teams.microsoft.com/v2/) in the Alfred browser.\n" +
    "If the problem persists, try: restart Alfred, log into Teams, then retry."
  );
}

// ---------------------------------------------------------------------------
// Skype messaging API token — extracted from Teams cookies via CDP
// Used for chats (Graph Chat.Read is not available on the Teams client app)
// ---------------------------------------------------------------------------

/**
 * Extract the skypetoken_asm cookie from the Alfred browser via CDP.
 * This token authenticates against the Teams Skype messaging API
 * (emea/amer.ng.msg.teams.microsoft.com) for chat operations.
 * The JWT payload contains `rgn` (region) and `exp` (expiry).
 */
async function acquireSkypeToken(progress: ProgressFn): Promise<{ token: string; region: string }> {
  // Check in-memory cache
  if (skypeTokenCache && Date.now() < skypeTokenCache.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    return { token: skypeTokenCache.token, region: skypeTokenCache.region };
  }

  // Check file cache
  const fileCached = loadCachedAuth("teamsSkypeToken");
  if (fileCached && Date.now() < fileCached.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    // Decode region from cached token
    let region = "amer";
    try {
      const payload = JSON.parse(Buffer.from(fileCached.value.split(".")[1]!, "base64url").toString());
      if (payload.rgn) region = payload.rgn;
    } catch (e) {
      process.stderr.write(`[alfred:warn] Skype token JWT decode failed, using default region: ${e instanceof Error ? e.message : String(e)}\n`);
    }
    skypeTokenCache = { token: fileCached.value, region, expiresAt: fileCached.expiresAt };
    return { token: fileCached.value, region };
  }

  if (!await isAlfredgable()) {
    throw new Error("The Alfred browser is not running. Launch Alfred from your Desktop first.");
  }

  progress("🔑 Extracting Teams chat token from browser cookies...");

  const ctx = await getAlfredContext();
  const allCookies = await ctx.cookies();
  const skypeCookie = allCookies.find(c =>
    c.name === "skypetoken_asm" && (c.domain === "teams.microsoft.com" || c.domain.endsWith(".teams.microsoft.com"))
  );

  if (!skypeCookie?.value) {
    throw new Error(
      "Could not extract Teams chat token from browser cookies.\n" +
      "Make sure you are logged into Teams (teams.microsoft.com/v2/) in the Alfred browser."
    );
  }

  let region = "amer";
  let expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
  try {
    const payload = JSON.parse(Buffer.from(skypeCookie.value.split(".")[1]!, "base64url").toString());
    if (payload.rgn) region = payload.rgn;
    if (payload.exp) expiresAt = payload.exp * 1000;
  } catch (e) {
    process.stderr.write(`[alfred:warn] Skype cookie JWT decode failed: ${e instanceof Error ? e.message : String(e)}\n`);
  }

  if (Date.now() >= expiresAt) {
    throw new Error(
      "Could not extract Teams chat token from browser cookies.\n" +
      "Make sure you are logged into Teams (teams.microsoft.com/v2/) in the Alfred browser."
    );
  }

  skypeTokenCache = { token: skypeCookie.value, region, expiresAt };
  saveCachedAuth("teamsSkypeToken", skypeCookie.value, expiresAt);
  progress(`✅ Teams chat token acquired (region: ${region})`);
  return { token: skypeCookie.value, region };
}

/** Known Teams Skype messaging API regions, ordered by likelihood. */
const SKYPE_REGIONS = ["amer", "emea", "apac"];

/** Make an authenticated request to the Teams Skype messaging API.
 *  If the primary region fails with a non-auth error, tries other regions
 *  (handles cases where the JWT rgn field doesn't match the data region). */
async function skypeFetch(
  path: string,
  token: string,
  region: string
): Promise<Record<string, unknown>> {
  const regionsToTry = [region, ...SKYPE_REGIONS.filter(r => r !== region)];

  for (let i = 0; i < regionsToTry.length; i++) {
    const r = regionsToTry[i]!;
    const url = `https://${r}.ng.msg.teams.microsoft.com${path}`;
    const res = await fetch(url, {
      headers: { Authentication: `skypetoken=${token}`, Accept: "application/json" },
      signal: AbortSignal.timeout(30_000),
    });

    if (res.ok) {
      // If a fallback region worked, update the cache so future calls use it directly
      if (i > 0 && skypeTokenCache) {
        process.stderr.write(`[alfred] Skype API: region ${region} failed, ${r} works — updating cache\n`);
        skypeTokenCache = { ...skypeTokenCache, region: r };
      }
      return res.json() as Promise<Record<string, unknown>>;
    }

    if (res.status === 401 || res.status === 403) {
      // Auth failure — don't try other regions, the token itself is bad
      skypeTokenCache = null;
      clearCachedAuthFile("teamsSkypeToken");
      const body = await res.text().catch(() => "");
      throw new Error(`Teams messaging API ${res.status}: ${body.slice(0, 300)}`);
    }

    // Non-auth error on non-last region — try next region
    if (i < regionsToTry.length - 1) {
      process.stderr.write(`[alfred:warn] Skype API ${r} returned ${res.status}, trying next region...\n`);
      continue;
    }

    // Last region also failed
    const body = await res.text().catch(() => "");
    throw new Error(`Teams messaging API ${res.status}: ${body.slice(0, 300)}`);
  }

  // Should never reach here, but TypeScript needs it
  throw new Error("Teams messaging API: no regions available");
}

async function graphFetch(path: string, token: string, _retryCount = 0): Promise<Record<string, unknown>> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
    signal: AbortSignal.timeout(30_000),
  });

  // Handle 429 throttling with exponential backoff
  if (res.status === 429 && _retryCount < 3) {
    const retryAfter = parseInt(res.headers.get("Retry-After") ?? "", 10);
    const delayMs = (retryAfter > 0 ? retryAfter * 1000 : 1000 * Math.pow(2, _retryCount));
    await new Promise(r => setTimeout(r, delayMs));
    return graphFetch(path, token, _retryCount + 1);
  }

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    // Only clear cache for auth failures on non-transcript endpoints (transcript 401 = scope issue, not bad token)
    if (res.status === 401 && !path.includes("transcript")) { teamsTokenCache = null; clearCachedAuthFile("teamsGraphToken"); }
    // Detect scope permission errors and give actionable guidance
    if (res.status === 403 && body.includes("Missing scope permissions")) {
      const match = body.match(/API requires one of '([^']+)'/);
      const needed = match?.[1] ?? "unknown scopes";
      teamsTokenCache = null; clearCachedAuthFile("teamsGraphToken");
      throw new Error(
        `Graph 403: Token is missing required scope (${needed}).\n` +
        `The Teams client app does not have this Graph permission configured.\n` +
        `This scope requires an Azure AD app registration with admin consent.`
      );
    }
    throw new Error(`Graph ${res.status} ${res.statusText}${body ? `: ${body.slice(0, 300)}` : ""}`);
  }
  return res.json() as Promise<Record<string, unknown>>;
}

// ---------------------------------------------------------------------------
// Teams transcript
// ---------------------------------------------------------------------------

export interface MeetingTranscript {
  meetingId: string;
  subject: string;
  start: string;
  end?: string;
  attendees: string[];
  transcriptId?: string;
  transcript?: string;
}

export async function getTeamsTranscript(opts: {
  search?: string;
  startDate?: string;
  endDate?: string;
  meetingId?: string;
}, progress: ProgressFn = () => {}): Promise<MeetingTranscript[]> {
  const token = await acquireTeamsGraphToken(progress);

  // Search calendar for matching meetings
  const start = opts.startDate ?? new Date(Date.now() - 30 * 86400_000).toISOString().slice(0, 10);
  const end   = opts.endDate   ?? new Date().toISOString().slice(0, 10);

  progress(`🔍 Searching calendar ${start} → ${end} for online meetings...`);

  const params = new URLSearchParams({
    startDateTime: `${start}T00:00:00Z`,
    endDateTime:   `${end}T23:59:59Z`,
    $top: "50",
  });

  const calData = await graphFetch(`/me/calendarView?${params}`, token);
  const allEvents = calData.value as Record<string, unknown>[] ?? [];
  const events = allEvents.filter(e =>
    (e.isOnlineMeeting === true) &&
    (e.onlineMeeting as Record<string, unknown> | null)?.joinUrl
  );

  progress(`📅 Found ${events.length} online meeting(s) in range`);

  // Search OneDrive for transcript files — Teams stores transcripts as
  // "{subject}-{datetime}-Meeting Transcript.mp4" in the Recordings folder.
  // We have Files.ReadWrite.All scope (no OnlineMeetings.Read needed).
  progress("📂 Searching OneDrive for transcript files...");
  let transcriptFiles: Array<{ name: string; id: string; modified: string; size: number; webUrl?: string }> = [];
  try {
    const searchData = await graphFetch(
      `/me/drive/root/search(q='Meeting Transcript')?$top=50&$select=name,id,lastModifiedDateTime,size,webUrl`,
      token
    );
    transcriptFiles = ((searchData.value as Record<string, unknown>[]) ?? [])
      .filter(f => (f.name as string)?.includes("Transcript"))
      .map(f => ({
        name: f.name as string,
        id: f.id as string,
        modified: f.lastModifiedDateTime as string,
        size: f.size as number,
        webUrl: f.webUrl as string | undefined,
      }));
    progress(`📄 Found ${transcriptFiles.length} transcript file(s) in OneDrive`);
  } catch (e) {
    process.stderr.write(`[alfred:warn] OneDrive transcript search failed: ${e instanceof Error ? e.message : String(e)}\n`);
  }

  const results: MeetingTranscript[] = [];

  for (const event of events) {
    const subject = (event.subject as string) ?? "";
    const startTime = (event.start as { dateTime: string })?.dateTime ?? "";
    const endTime = (event.end as { dateTime: string })?.dateTime;

    // Apply search filter if specified
    if (opts.search && !subject.toLowerCase().includes(opts.search.toLowerCase())) continue;

    // Match this meeting to a transcript file by subject + date.
    // Filename format: "{subject}-{YYYYMMDD}_{HHMMSS}-Meeting Transcript.mp4"
    // Calendar subjects may use en-dashes (–) while filenames use hyphens (-),
    // so we normalize both to alphanumeric-only for comparison.
    const subjectNorm = subject.toLowerCase().replace(/[^a-z0-9]/g, "");
    const eventDate = startTime.slice(0, 10).replace(/-/g, ""); // YYYYMMDD
    const matchingFile = transcriptFiles.find(f => {
      const parsed = f.name.match(/^(.+)-(\d{8})_\d{6}-Meeting Transcript\.mp4$/i);
      if (!parsed) return false;
      const fileSubjectNorm = parsed[1]!.toLowerCase().replace(/[^a-z0-9]/g, "");
      const fileDate = parsed[2]!;
      // Match by subject (exact normalized match OR one contains the other for partial matches)
      const subjectMatch = fileSubjectNorm === subjectNorm ||
        (subjectNorm.length > 5 && (fileSubjectNorm.includes(subjectNorm) || subjectNorm.includes(fileSubjectNorm)));
      // Date must match (same day)
      return subjectMatch && fileDate === eventDate;
    }) ?? transcriptFiles.find(f => {
      // Fallback: subject-only match (for cases where date extraction fails)
      const fileSubject = f.name.replace(/-\d{8}_\d{6}-Meeting Transcript\.mp4$/i, "");
      return fileSubject.toLowerCase().replace(/[^a-z0-9]/g, "") === subjectNorm;
    });

    const attendees = ((event.attendees as Array<{ emailAddress: { address: string } }>) ?? [])
      .map(a => a.emailAddress?.address).filter(Boolean);

    const meetingId = (event.id as string) ?? "";

    if (matchingFile) {
      results.push({
        meetingId,
        subject,
        start: startTime,
        end: endTime,
        attendees,
        transcriptId: matchingFile.id,
        transcript: `[Transcript available in OneDrive: ${matchingFile.name}]` +
          (matchingFile.webUrl ? `\nView: ${matchingFile.webUrl}` : ""),
      });
    } else {
      // No transcript file found — still include the meeting with a note
      results.push({
        meetingId,
        subject,
        start: startTime,
        end: endTime,
        attendees,
        transcript: "(No transcript file found in OneDrive for this meeting)",
      });
    }
  }

  // Also include any transcript files that didn't match a calendar event
  // (e.g. meetings outside the date range but with transcripts available)
  const matchedIds = new Set(results.filter(r => r.transcriptId).map(r => r.transcriptId));
  for (const f of transcriptFiles) {
    if (matchedIds.has(f.id)) continue;
    // Extract subject and date from filename: "Subject-YYYYMMDD_HHMMSS-Meeting Transcript.mp4"
    const match = f.name.match(/^(.+)-(\d{8})_(\d{6})-Meeting Transcript\.mp4$/i);
    if (!match) continue;
    const [, fileSubject, dateStr] = match;
    const isoDate = `${dateStr!.slice(0, 4)}-${dateStr!.slice(4, 6)}-${dateStr!.slice(6, 8)}`;
    // Only include if the search matches (or no search specified)
    if (opts.search && !fileSubject!.toLowerCase().includes(opts.search.toLowerCase())) continue;
    results.push({
      meetingId: f.id,
      subject: fileSubject!.replace(/ - /g, " – "),
      start: `${isoDate}T00:00:00`,
      attendees: [],
      transcriptId: f.id,
      transcript: `[Transcript available in OneDrive: ${f.name}]` +
        (f.webUrl ? `\nView: ${f.webUrl}` : ""),
    });
  }

  progress(`✅ Found ${results.filter(r => r.transcriptId).length} meeting(s) with transcripts`);
  return results;
}

// ---------------------------------------------------------------------------
// Teams chats
// ---------------------------------------------------------------------------

export interface ChatMessage {
  id: string;
  from: string;
  body: string;
  createdDateTime: string;
}

export interface TeamsChat {
  chatId: string;
  topic?: string;
  chatType: string;
  members: string[];
  lastMessageAt?: string;
  messages?: ChatMessage[];
}

export async function getTeamsChats(opts: {
  search?: string;
  chatId?: string;
  includeMessages?: boolean;
  top?: number;
}, progress: ProgressFn = () => {}): Promise<TeamsChat[]> {
  // Use the Skype messaging API — the Graph Chat.Read scope is not available
  // on the Teams first-party app, but the Skype API uses the browser session cookie.
  const { token, region } = await acquireSkypeToken(progress);

  // Direct lookup by chat ID — fetch messages for that conversation
  if (opts.chatId) {
    progress(`💬 Fetching chat ${opts.chatId}...`);
    const messages = await fetchSkypeChatMessages(opts.chatId, token, region, opts.top ?? 50, progress);
    return [{
      chatId: opts.chatId,
      chatType: opts.chatId.includes("meeting_") ? "meeting" : "chat",
      members: [],
      messages,
    }];
  }

  progress("💬 Fetching Teams chats via messaging API...");
  const pageSize = Math.min(opts.top ?? 50, 50);
  const data = await skypeFetch(
    `/v1/users/ME/conversations?view=mychats&pageSize=${pageSize}`,
    token, region
  );

  const conversations = (data.conversations as Record<string, unknown>[]) ?? [];
  let chats: TeamsChat[] = conversations.map(c => mapSkypeConversation(c));

  // Filter by search term (matches topic or member names)
  if (opts.search) {
    const q = opts.search.toLowerCase();
    chats = chats.filter(c =>
      c.topic?.toLowerCase().includes(q) ||
      c.members.some(m => m.toLowerCase().includes(q))
    );
  }

  progress(`✅ Found ${chats.length} chat(s)`);

  if (opts.includeMessages) {
    const MSG_BATCH = 5;
    for (let i = 0; i < chats.length; i += MSG_BATCH) {
      const batch = chats.slice(i, i + MSG_BATCH);
      const results = await Promise.allSettled(
        batch.map(chat => fetchSkypeChatMessages(chat.chatId, token, region, 25, progress))
      );
      for (let j = 0; j < results.length; j++) {
        const result = results[j]!;
        batch[j]!.messages = result.status === "fulfilled" ? result.value : [];
      }
    }
  }

  return chats;
}

/** Map a Skype conversation object to our TeamsChat interface. */
function mapSkypeConversation(raw: Record<string, unknown>): TeamsChat {
  const tp = (raw.threadProperties ?? {}) as Record<string, unknown>;
  const topic = (tp.topicThreadTopic ?? tp.spaceThreadTopic ?? tp.topic ?? null) as string | null;
  const id = raw.id as string;
  const lastMsg = raw.lastMessage as Record<string, unknown> | undefined;

  // Determine chat type from the conversation ID format
  let chatType = "chat";
  if (id.includes("meeting_")) chatType = "meeting";
  else if (tp.spaceType === "standard" || tp.productThreadType === "TeamsTeam") chatType = "channel";
  else if (id.startsWith("19:") && tp.chatModalityType === "Conversational") chatType = "group";

  return {
    chatId: id,
    topic: topic ?? undefined,
    chatType,
    members: [], // Skype API doesn't include member list in conversation listing
    lastMessageAt: (lastMsg?.composetime as string) ?? undefined,
  };
}

async function fetchSkypeChatMessages(
  chatId: string,
  token: string,
  region: string,
  top: number,
  progress: ProgressFn
): Promise<ChatMessage[]> {
  progress(`📨 Fetching messages for chat ${chatId}...`);
  try {
    const data = await skypeFetch(
      `/v1/users/ME/conversations/${encodeURIComponent(chatId)}/messages?pageSize=${top}`,
      token, region
    );
    const messages = (data.messages as Record<string, unknown>[]) ?? [];
    return messages
      .filter(m => {
        const type = m.messagetype as string;
        // Only include actual user messages, not system events
        return type === "RichText/Html" || type === "Text" || type === "RichText";
      })
      .map(m => ({
        id: (m.id ?? m.skypeeditedid ?? "") as string,
        from: (m.imdisplayname as string) || ((m.from as string) ?? "").split("/").pop()?.split(":").pop() || "Unknown",
        body: stripHtml((m.content as string) ?? ""),
        createdDateTime: (m.composetime as string) ?? "",
      }))
      .filter(m => m.body);
  } catch (e) {
    process.stderr.write(`[alfred:warn] fetchSkypeChatMessages failed for ${chatId}: ${e instanceof Error ? e.message : String(e)}\n`);
    return [];
  }
}

