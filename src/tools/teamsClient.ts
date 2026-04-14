import { execFileSync } from "child_process";
import { connectWithRetry } from "../auth/tokenExtractor.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";
import { loadCachedAuth, saveCachedAuth, clearCachedAuthFile } from "../auth/authFileCache.js";
import { stripHtml, urlHostMatches } from "../shared.js";

const CDP_PORT = 9222;

/** Match both legacy and new Outlook domains */
function isOutlookUrl(url: string): boolean {
  return urlHostMatches(url, "outlook.office.com") || urlHostMatches(url, "outlook.cloud.microsoft.com");
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

// Try to extract Graph token from an existing CDP tab via raw WebSocket
async function extractGraphTokenFromTab(wsUrl: string): Promise<string | null> {
  return new Promise((resolve) => {
    const ws = new WebSocket(wsUrl);
    const timer = setTimeout(() => {
      process.stderr.write("[alfred:warn] Teams MSAL extraction timed out (5s) — tab may be unresponsive\n");
      try { ws.close(); } catch {}
      resolve(null);
    }, 5_000);
    ws.addEventListener("open", () => {
      ws.send(JSON.stringify({ id: 1, method: "Runtime.evaluate", params: { expression: GRAPH_MSAL_EXTRACT_JS, returnByValue: true } }));
    });
    ws.addEventListener("message", (event: MessageEvent) => {
      clearTimeout(timer);
      try { ws.close(); } catch {}
      try {
        const msg = JSON.parse(event.data as string) as { result?: { result?: { value?: string } } };
        resolve(msg.result?.result?.value ?? null);
      } catch (e) {
        process.stderr.write(`[alfred:warn] Teams MSAL parse error: ${e instanceof Error ? e.message : String(e)}\n`);
        resolve(null);
      }
    });
    ws.addEventListener("error", () => {
      clearTimeout(timer);
      process.stderr.write("[alfred:warn] Teams MSAL extraction WebSocket error\n");
      try { ws.close(); } catch {}
      resolve(null);
    });
  });
}

interface TokenCache { token: string; expiresAt: number; }
let teamsTokenCache: TokenCache | null = null;

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

  try {
    execFileSync("curl", ["-s", "--max-time", "3", `http://localhost:${CDP_PORT}/json/version`], { timeout: 5_000 });
  } catch {
    throw new Error("Alfred is not running. Open Alfred.app first.");
  }

  progress("🔐 Acquiring Graph token via Teams/Outlook in Alfred...");

  // Step 1: try to read Graph token directly from MSAL cache in open tabs (fast, no new page)
  const listRes = await fetch(`http://localhost:${CDP_PORT}/json/list`).catch(() => null);
  const targets = listRes?.ok
    ? await listRes.json() as Array<{ webSocketDebuggerUrl?: string; type?: string; url?: string }>
    : [];

  // Prefer Teams tab — it has broader Graph scopes (Chat.Read etc.)
  const teamsTarget = targets.find(t => t.type === "page" && t.webSocketDebuggerUrl &&
    t.url && urlHostMatches(t.url, "teams.microsoft.com"));
  const outlookTarget = targets.find(t => t.type === "page" && t.webSocketDebuggerUrl &&
    t.url && isOutlookUrl(t.url));
  const candidates = [teamsTarget, outlookTarget].filter(Boolean) as typeof targets;

  if (!teamsTarget) {
    process.stderr.write("[alfred] No Teams tab found in Alfred — Graph token capture may fail. Open teams.microsoft.com in Alfred.\n");
  }

  // Try MSAL extraction from each candidate tab
  for (let attempt = 0; attempt < 3; attempt++) {
    for (const t of candidates) {
      const token = await extractGraphTokenFromTab(t.webSocketDebuggerUrl!);
      if (token) {
        const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
        teamsTokenCache = { token, expiresAt };
        saveCachedAuth("teamsGraphToken", token, expiresAt);
        progress("✅ Graph token acquired from MSAL cache");
        return token;
      }
    }
    if (attempt < 2) await new Promise(r => setTimeout(r, 2_000));
  }

  // Step 2: MSAL miss (likely encrypted on outlook.cloud.microsoft.com) —
  // use raw CDP network interception with Page.reload to capture Graph token
  const interceptTarget = teamsTarget ?? outlookTarget;
  if (interceptTarget?.webSocketDebuggerUrl) {
    progress("🔑 MSAL cache miss — capturing Graph token via network interception...");
    const capturedToken = await new Promise<string | null>((resolve) => {
      const ws = new WebSocket(interceptTarget.webSocketDebuggerUrl!);
      let found: string | null = null;

      const timer = setTimeout(() => {
        process.stderr.write("[alfred:warn] Graph token network interception timed out (12s)\n");
        try { ws.close(); } catch {}
        resolve(null);
      }, 12_000);

      let msgId = 0;
      const send = (method: string, params?: Record<string, unknown>) =>
        ws.send(JSON.stringify({ id: ++msgId, method, params }));

      ws.addEventListener("open", () => send("Network.enable"));

      ws.addEventListener("message", (event: MessageEvent) => {
        try {
          const msg = JSON.parse(event.data as string) as { id?: number; method?: string; params?: Record<string, unknown> };
          if (msg.id === 1) {
            // Network enabled — trigger Graph API calls + Page.reload as backup
            send("Runtime.evaluate", {
              expression: [
                `fetch('https://graph.microsoft.com/v1.0/me', { credentials: 'include' }).catch(()=>{})`,
                `fetch('https://graph.microsoft.com/v1.0/me/chats?$top=1', { credentials: 'include' }).catch(()=>{})`,
              ].join(';'),
              awaitPromise: false,
            });
            // Reload page after 3s to trigger app's own Graph API calls with MSAL tokens
            setTimeout(() => { if (!found) send("Page.reload"); }, 3_000);
          }
          if (msg.method === "Network.requestWillBeSent") {
            const reqUrl = ((msg.params?.request as Record<string, unknown>)?.url ?? "") as string;
            // Only capture tokens from graph.microsoft.com requests (not Outlook REST)
            if (!reqUrl.includes("graph.microsoft.com")) return;
            const headers = ((msg.params?.request as Record<string, unknown>)?.headers ?? {}) as Record<string, string>;
            const auth = headers["Authorization"] ?? headers["authorization"] ?? "";
            if (!found && auth.startsWith("Bearer ")) {
              found = auth.slice(7);
              clearTimeout(timer);
              try { ws.close(); } catch {}
              resolve(found);
            }
          }
        } catch (e) { process.stderr.write(`[alfred:warn] Graph CDP interception error: ${e instanceof Error ? e.message : String(e)}\n`); }
      });

      ws.addEventListener("error", () => {
        clearTimeout(timer);
        process.stderr.write("[alfred:warn] Graph network interception WebSocket error\n");
        try { ws.close(); } catch {}
        resolve(null);
      });
    });

    if (capturedToken) {
      const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
      teamsTokenCache = { token: capturedToken, expiresAt };
      saveCachedAuth("teamsGraphToken", capturedToken, expiresAt);
      progress("✅ Graph token captured via network interception");
      return capturedToken;
    }
  }

  // Step 3: fallback — use Playwright to capture token from existing Teams tab
  progress("📡 Falling back to Playwright for Graph token capture...");
  const browser = await connectWithRetry();

  try {
    const ctx = browser.contexts()[0];
    if (!ctx) throw new Error("No browser context found");

    // Prefer reusing an existing Teams tab — avoids opening new windows
    const existingPages = ctx.pages();
    let page = existingPages.find(p => urlHostMatches(p.url(), "teams.microsoft.com"));

    let isNewPage = false;
    if (!page) {
      // No Teams tab — try Outlook tab (both domains), else reuse any tab
      page = existingPages.find(p => isOutlookUrl(p.url()))
          ?? existingPages.find(p => p.url().startsWith("http"))
          ?? await ctx.newPage();
      isNewPage = !existingPages.includes(page);
    }

    let capturedToken: string | null = null;
    await page.route("**/graph.microsoft.com/**", async (route) => {
      const auth = route.request().headers()["authorization"] ?? "";
      if (!capturedToken && auth.startsWith("Bearer ")) capturedToken = auth.slice(7);
      await route.continue();
    });

    progress("📡 Loading Teams to capture Graph token (broad scopes)...");
    if (!urlHostMatches(page.url(), "teams.microsoft.com")) {
      await page.goto("https://teams.microsoft.com/v2/", { waitUntil: "domcontentloaded", timeout: 20_000 }).catch((e) => { process.stderr.write(`[alfred:warn] Teams page.goto failed: ${e instanceof Error ? e.message : String(e)}\n`); });
    } else {
      await page.reload({ waitUntil: "domcontentloaded", timeout: 20_000 }).catch((e) => { process.stderr.write(`[alfred:warn] Teams page.reload failed: ${e instanceof Error ? e.message : String(e)}\n`); });
    }

    const deadline = Date.now() + 10_000;
    while (!capturedToken && Date.now() < deadline) await page.waitForTimeout(500);

    // Clean up route interceptor and only close pages we created
    await page.unroute("**/graph.microsoft.com/**").catch((e) => { process.stderr.write(`[alfred:warn] unroute failed: ${e instanceof Error ? e.message : String(e)}\n`); });
    if (isNewPage) await page.close();

    if (!capturedToken) throw new Error(
      "Could not capture Graph token.\n" +
      "Open https://teams.microsoft.com in the Alfred Chrome window and log in, then retry.\n" +
      "Teams tokens have the broad Graph scopes needed for transcripts and chats."
    );

    const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
    teamsTokenCache = { token: capturedToken, expiresAt };
    saveCachedAuth("teamsGraphToken", capturedToken, expiresAt);
    progress("✅ Graph token acquired");
    return capturedToken;
  } finally {
    // Do NOT call browser.close() — it kills the user's actual Alfred Chrome process.
    // The CDP connection wrapper is GC'd; Alfred Chrome keeps running.
  }
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
        `This usually means the token was captured from Outlook (which lacks Chat scopes).\n` +
        `Fix: Open the Teams tab in Alfred and retry — Teams tokens have broader scopes.`
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

  // Direct lookup by meeting ID
  if (opts.meetingId) {
    progress(`📋 Fetching transcript for meeting ${opts.meetingId}...`);
    return [await fetchTranscriptForMeeting(opts.meetingId, token, progress)];
  }

  // Search calendar for matching meetings
  const start = opts.startDate ?? new Date(Date.now() - 30 * 86400_000).toISOString().slice(0, 10);
  const end   = opts.endDate   ?? new Date().toISOString().slice(0, 10);

  progress(`🔍 Searching calendar ${start} → ${end} for meetings with transcripts...`);

  const params = new URLSearchParams({
    startDateTime: `${start}T00:00:00Z`,
    endDateTime:   `${end}T23:59:59Z`,
    $select: "id,subject,start,end,attendees,onlineMeeting",
    $top: "50",
  });
  if (opts.search) params.set("$search", `"${opts.search.replace(/"/g, "")}"`);


  const calData = await graphFetch(`/me/calendarView?${params}`, token);
  const events = (calData.value as Record<string, unknown>[] ?? [])
    .filter(e => e.onlineMeeting);

  progress(`📅 Found ${events.length} online meeting(s) in range`);

  const results: MeetingTranscript[] = [];
  for (const event of events) {
    const joinUrl = (event.onlineMeeting as { joinUrl?: string })?.joinUrl;
    if (!joinUrl) continue;

    try {
      // Resolve calendar event → onlineMeeting ID
      const meetingData = await graphFetch(
        `/me/onlineMeetings?$filter=JoinWebUrl eq '${joinUrl.replace(/'/g, "''")}'`,
        token
      );
      const meeting = (meetingData.value as Record<string, unknown>[])?.[0];
      if (!meeting) continue;

      const meetingId = meeting.id as string;
      const result = await fetchTranscriptForMeeting(meetingId, token, progress);
      result.subject    = event.subject as string;
      result.start      = (event.start as { dateTime: string })?.dateTime ?? result.start;
      result.end        = (event.end   as { dateTime: string })?.dateTime;
      result.attendees  = ((event.attendees as { emailAddress: { address: string } }[]) ?? [])
        .map(a => a.emailAddress?.address).filter(Boolean);
      results.push(result);
    } catch (e) {
      process.stderr.write(`[alfred:warn] transcript fetch skipped for meeting: ${e instanceof Error ? e.message : String(e)}\n`);
    }
  }

  progress(`✅ Retrieved ${results.length} transcript(s)`);
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
  const token = await acquireTeamsGraphToken(progress);

  // Direct lookup by chat ID
  if (opts.chatId) {
    progress(`💬 Fetching chat ${opts.chatId}...`);
    const chatData = await graphFetch(
      `/me/chats/${opts.chatId}?$expand=members`,
      token
    );
    const chat = mapChat(chatData);
    if (opts.includeMessages) {
      chat.messages = await fetchChatMessages(opts.chatId, token, opts.top ?? 50, progress);
    }
    return [chat];
  }

  progress("💬 Fetching Teams chats...");
  const params = new URLSearchParams({
    $expand: "members",
    $top: String(opts.top ?? 50),
    $orderby: "lastMessagePreview/createdDateTime desc",
  });

  const data = await graphFetch(`/me/chats?${params}`, token);
  let chats = (data.value as Record<string, unknown>[] ?? []).map(mapChat);

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
    for (const chat of chats) {
      chat.messages = await fetchChatMessages(chat.chatId, token, 25, progress);
    }
  }

  return chats;
}

function mapChat(raw: Record<string, unknown>): TeamsChat {
  const members = ((raw.members as Record<string, unknown>[]) ?? [])
    .map(m => (m.displayName as string) || (m.email as string) || "")
    .filter(Boolean);
  return {
    chatId:         raw.id as string,
    topic:          (raw.topic as string | null) ?? undefined,
    chatType:       raw.chatType as string,
    members,
    lastMessageAt:  (raw.lastMessagePreview as Record<string, unknown> | null)?.createdDateTime as string | undefined,
  };
}

async function fetchChatMessages(
  chatId: string,
  token: string,
  top: number,
  progress: ProgressFn
): Promise<ChatMessage[]> {
  progress(`📨 Fetching messages for chat ${chatId}...`);
  try {
    const data = await graphFetch(
      `/me/chats/${chatId}/messages?$top=${top}&$orderby=createdDateTime desc`,
      token
    );
    return ((data.value as Record<string, unknown>[]) ?? []).map(m => ({
      id:              m.id as string,
      from:            ((m.from as Record<string, unknown>)?.user as Record<string, unknown>)?.displayName as string ?? "Unknown",
      body:            stripHtml((m.body as Record<string, unknown>)?.content as string ?? ""),
      createdDateTime: m.createdDateTime as string,
    })).filter(m => m.body);
  } catch (e) {
    process.stderr.write(`[alfred:warn] fetchChatMessages failed for ${chatId}: ${e instanceof Error ? e.message : String(e)}\n`);
    return [];
  }
}

async function fetchTranscriptForMeeting(
  meetingId: string,
  token: string,
  progress: ProgressFn
): Promise<MeetingTranscript> {
  progress(`📋 Fetching transcripts for meeting ${meetingId}...`);

  const txData = await graphFetch(`/me/onlineMeetings/${meetingId}/transcripts`, token);
  const transcripts = txData.value as Record<string, unknown>[] ?? [];

  if (transcripts.length === 0) {
    return { meetingId, subject: "", start: "", attendees: [], transcript: "(No transcript available)" };
  }

  // Get the most recent transcript
  const latest = transcripts[0];
  const transcriptId = latest.id as string;

  progress(`📄 Downloading transcript ${transcriptId}...`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content?$format=text/vtt`,
    { headers: { Authorization: `Bearer ${token}`, Accept: "text/vtt" }, signal: AbortSignal.timeout(30_000) }
  );

  let transcriptText = "(Could not download transcript)";
  if (res.ok) {
    const vtt = await res.text();
    // Strip VTT timestamps, keep only the spoken text
    transcriptText = vtt
      .split("\n")
      .filter(l => l && !l.startsWith("WEBVTT") && !l.match(/^\d+$/) && !l.match(/\d\d:\d\d:\d\d/))
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
  }

  return { meetingId, subject: "", start: "", attendees: [], transcriptId, transcript: transcriptText };
}
