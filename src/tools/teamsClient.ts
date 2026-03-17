import { execFileSync } from "child_process";
import { connectWithRetry } from "../auth/tokenExtractor.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";

const CDP_PORT = 9222;
const TOKEN_CACHE_MS = 45 * 60 * 1000;

// ---------------------------------------------------------------------------
// Teams webhook config (set once via configure_teams_webhook tool)
// ---------------------------------------------------------------------------

let webhookUrl: string | null = null;

export function setTeamsWebhook(url: string): void {
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
  const res = await fetch(webhookUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      type: "message",
      attachments: [{
        contentType: "application/vnd.microsoft.card.adaptive",
        content: card,
      }],
    }),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Teams webhook error: ${res.status} ${res.statusText}${text ? ` — ${text}` : ""}`);
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
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Teams webhook error: ${res.status} ${res.statusText}${text ? ` — ${text}` : ""}`);
  }

  progress("✅ Teams notification sent");
}

// ---------------------------------------------------------------------------
// Graph Bearer token via Teams web page (CDP route interception)
// ---------------------------------------------------------------------------

interface TokenCache { token: string; expiresAt: number; }
let teamsTokenCache: TokenCache | null = null;

async function acquireTeamsGraphToken(progress: ProgressFn): Promise<string> {
  if (teamsTokenCache && Date.now() < teamsTokenCache.expiresAt) {
    const mins = Math.round((teamsTokenCache.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Teams Graph token (~${mins} min remaining)`);
    return teamsTokenCache.token;
  }

  try {
    execFileSync("curl", ["-s", "--max-time", "1", `http://localhost:${CDP_PORT}/json/version`], { timeout: 2_000 });
  } catch {
    throw new Error("ChromeLink is not running. Open ChromeLink.app first.");
  }

  progress("🔐 Acquiring Graph token via Teams/Outlook in ChromeLink...");
  const browser = await connectWithRetry();

  try {
    const ctx = browser.contexts()[0];
    if (!ctx) throw new Error("No browser context found");

    const page = await ctx.newPage();
    let capturedToken: string | null = null;

    await page.route("**/*", async (route) => {
      const auth = route.request().headers()["authorization"] ?? "";
      if (!capturedToken && auth.startsWith("Bearer ")) {
        capturedToken = auth.slice(7);
      }
      await route.continue();
    });

    // Try Teams first (richer scopes), fall back to Outlook
    progress("📡 Loading Teams to capture Graph token...");
    await page.goto("https://teams.microsoft.com/_#/conversations/", {
      waitUntil: "domcontentloaded",
      timeout: 20_000,
    }).catch(() => {});

    const deadline = Date.now() + 8_000;
    while (!capturedToken && Date.now() < deadline) {
      await page.waitForTimeout(500);
    }

    // Fall back to Outlook if Teams didn't yield a token
    if (!capturedToken) {
      progress("📡 Trying Outlook as fallback...");
      await page.goto("https://outlook.office.com/mail/", {
        waitUntil: "domcontentloaded",
        timeout: 20_000,
      }).catch(() => {});

      const deadline2 = Date.now() + 8_000;
      while (!capturedToken && Date.now() < deadline2) {
        await page.waitForTimeout(500);
      }
    }

    await page.close();

    if (!capturedToken) {
      throw new Error(
        "Could not capture Graph token. Make sure Teams or Outlook is loaded in ChromeLink."
      );
    }

    teamsTokenCache = { token: capturedToken, expiresAt: Date.now() + TOKEN_CACHE_MS };
    progress("✅ Graph token acquired");
    return capturedToken;
  } finally {
    await browser.close();
  }
}

async function graphFetch(path: string, token: string): Promise<Record<string, unknown>> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    if (res.status === 401) teamsTokenCache = null;
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
  if (opts.search) params.set("$search", `"${opts.search}"`);

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
        `/me/onlineMeetings?$filter=JoinWebUrl eq '${encodeURIComponent(joinUrl)}'`,
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
    } catch {
      // Meeting might not have a transcript — skip silently
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
      body:            ((m.body as Record<string, unknown>)?.content as string ?? "").replace(/<[^>]+>/g, "").trim(),
      createdDateTime: m.createdDateTime as string,
    })).filter(m => m.body);
  } catch {
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
    { headers: { Authorization: `Bearer ${token}`, Accept: "text/vtt" } }
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
