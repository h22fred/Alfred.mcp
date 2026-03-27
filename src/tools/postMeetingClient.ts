import { getCalendarEvents } from "./outlookClient.js";
import { getTeamsTranscript, postAdaptiveCard } from "./teamsClient.js";
import { fetchOpportunities } from "./dynamicsClient.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";

export interface PostMeetingCandidate {
  // Meeting details
  meetingSubject: string;
  meetingStart: string;
  meetingEnd?: string;
  organizer?: string;
  attendees: { name: string; email: string }[];
  attendeeNames: string[];   // convenience flat list for display
  durationMinutes?: number;

  // Transcript (if available)
  transcriptAvailable: boolean;
  transcript?: string;

  // Opportunity match (best guess from attendee domains / subject / email domains)
  suggestedOpportunityId?: string;
  suggestedOpportunityName?: string;
  suggestedAccountName?: string;
  matchReason?: string;  // e.g. "subject", "domain", "organizer"

  // Raw data for Claude to reason about
  calendarEvent: Record<string, unknown>;
}

// Internal domains to exclude from attendee domain matching
const INTERNAL_DOMAINS = new Set([
  "servicenow.com", "now.com",
  "gmail.com", "outlook.com", "hotmail.com", "yahoo.com", "live.com",
  "microsoft.com", "google.com",
]);

function attendeeDomainWord(email: string): string | null {
  const domain = email.split("@")[1]?.toLowerCase();
  if (!domain || INTERNAL_DOMAINS.has(domain)) return null;
  return domain.split(".")[0]; // e.g. "pmi" from "pmi.org"
}

function matchScore(accountName: string, domainWord: string): boolean {
  const account = accountName.toLowerCase().replace(/[^a-z0-9]/g, "");
  const dw = domainWord.replace(/[^a-z0-9]/g, "");
  if (dw.length < 3) return false;
  return account.includes(dw) || dw.includes(account.slice(0, Math.max(4, account.length - 2)));
}

/**
 * Detects recently ended customer meetings, fetches transcripts where available,
 * and returns structured candidates for Claude to analyse and propose engagements from.
 *
 * Claude Desktop's Claude does the reasoning — this tool just collects the raw data.
 */
export async function detectPostMeetingEngagements(opts: {
  hoursBack?: number;   // how far back to look (default 24h)
  search?: string;      // optional keyword filter on meeting subject
}, progress: ProgressFn = () => {}): Promise<PostMeetingCandidate[]> {
  const hoursBack = opts.hoursBack ?? 24;
  const now = new Date();
  const since = new Date(now.getTime() - hoursBack * 3_600_000);

  const startDate = since.toISOString().slice(0, 10);
  const endDate   = now.toISOString().slice(0, 10);

  progress(`🔍 Scanning meetings from last ${hoursBack}h (${startDate} → ${endDate})...`);

  // 1. Fetch calendar events + opportunities in parallel (independent data sources)
  const [events, opportunities] = await Promise.all([
    getCalendarEvents(startDate, endDate, opts.search, progress),
    fetchOpportunities({ myOpportunitiesOnly: true, minNnacv: 100_000, top: 100 }, progress)
      .catch((): { opportunityid: string; name: string; accountName: string }[] => []),
  ]);

  // Keep only online meetings that have already ended
  const ended = events.filter(e => {
    if (!e.isOnlineMeeting) return false;
    const endTime = e.end ? new Date(e.end) : null;
    return endTime && endTime < now;
  });

  progress(`📅 Found ${ended.length} ended online meeting(s)`);

  if (ended.length === 0) {
    return [];
  }

  // 2. Fetch all transcripts for the week in one call, then match per meeting
  progress("📋 Fetching this week's transcripts...");
  let allTranscripts: Awaited<ReturnType<typeof getTeamsTranscript>> = [];
  let transcriptError: string | undefined;
  try {
    allTranscripts = await getTeamsTranscript({ startDate, endDate }, progress);
  } catch (e) {
    transcriptError = e instanceof Error ? e.message : String(e);
    process.stderr.write(`[alfred:warn] Transcript fetch failed: ${transcriptError}\n`);
    progress(`⚠️ Could not fetch transcripts: ${transcriptError} — continuing without them`);
  }

  const candidates: PostMeetingCandidate[] = [];

  for (const event of ended) {
    progress(`📋 Processing: "${event.subject}"...`);

    let transcript: string | undefined;
    let transcriptAvailable = false;

    const match = allTranscripts.find(t =>
      t.subject?.toLowerCase().includes(event.subject?.toLowerCase().slice(0, 20) ?? "")
      || Math.abs(new Date(t.start).getTime() - new Date(event.start).getTime()) < 30 * 60_000
    );

    if (match?.transcript && match.transcript !== "(No transcript available)") {
      transcript = match.transcript;
      transcriptAvailable = true;
    }

    // Best-effort opportunity matching: subject words, organizer name, attendee email domains
    let suggestedOpportunityId: string | undefined;
    let suggestedOpportunityName: string | undefined;
    let suggestedAccountName: string | undefined;
    let matchReason: string | undefined;

    // Collect external domain words from all attendees (excl. internal)
    const externalDomainWords = (event.attendees ?? [])
      .map(a => attendeeDomainWord(a.email))
      .filter((d): d is string => d !== null);
    if (event.organizerEmail) {
      const od = attendeeDomainWord(event.organizerEmail);
      if (od) externalDomainWords.push(od);
    }

    for (const opp of opportunities) {
      const accountWords = opp.accountName.split(/\s+/).filter(w => w.length > 3);

      const inSubject = accountWords.some(w =>
        event.subject?.toLowerCase().includes(w.toLowerCase())
      );
      const inOrganizer = event.organizer &&
        accountWords.some(w => event.organizer!.toLowerCase().includes(w.toLowerCase()));
      const inDomain = externalDomainWords.some(dw => matchScore(opp.accountName, dw));

      if (inSubject) {
        suggestedOpportunityId   = opp.opportunityid;
        suggestedOpportunityName = opp.name;
        suggestedAccountName     = opp.accountName;
        matchReason              = "subject";
        break;
      } else if (inDomain && !suggestedOpportunityId) {
        suggestedOpportunityId   = opp.opportunityid;
        suggestedOpportunityName = opp.name;
        suggestedAccountName     = opp.accountName;
        matchReason              = "domain";
      } else if (inOrganizer && !suggestedOpportunityId) {
        suggestedOpportunityId   = opp.opportunityid;
        suggestedOpportunityName = opp.name;
        suggestedAccountName     = opp.accountName;
        matchReason              = "organizer";
      }
    }

    // Duration
    let durationMinutes: number | undefined;
    if (event.start && event.end) {
      durationMinutes = Math.round(
        (new Date(event.end).getTime() - new Date(event.start).getTime()) / 60_000
      );
    }

    candidates.push({
      meetingSubject:          event.subject,
      meetingStart:            event.start,
      meetingEnd:              event.end,
      organizer:               event.organizer,
      attendees:               event.attendees ?? [],
      attendeeNames:           (event.attendees ?? []).map(a => a.name).filter(Boolean),
      durationMinutes,
      transcriptAvailable,
      transcript,
      suggestedOpportunityId,
      suggestedOpportunityName,
      suggestedAccountName,
      matchReason,
      calendarEvent:           event as unknown as Record<string, unknown>,
    });
  }

  progress(`✅ ${candidates.length} meeting candidate(s) ready for review`);
  return candidates;
}

/**
 * Post an Adaptive Card per candidate to Teams, summarising the meeting and
 * prompting the user to approve engagement creation in Claude.
 */
export async function notifyPostMeetingCandidates(
  candidates: PostMeetingCandidate[],
  progress: ProgressFn = () => {}
): Promise<number> {
  if (candidates.length === 0) return 0;

  let posted = 0;
  for (const c of candidates) {
    const attendeeList = c.attendeeNames.length
      ? c.attendeeNames.slice(0, 8).join(", ") + (c.attendeeNames.length > 8 ? ` +${c.attendeeNames.length - 8} more` : "")
      : "—";

    const matchLine = c.suggestedOpportunityName
      ? `**Suggested opp:** ${c.suggestedOpportunityName} (${c.suggestedAccountName ?? "—"}) — matched by ${c.matchReason}`
      : "_No matching opportunity found — please specify manually._";

    const durationLine = c.durationMinutes ? `${c.durationMinutes} min` : "—";
    const transcriptLine = c.transcriptAvailable ? "Transcript available" : "No transcript";

    const body: Record<string, unknown>[] = [
      { type: "TextBlock", text: `📋 Post-meeting: **${c.meetingSubject}**`, weight: "Bolder", size: "Medium", wrap: true },
      { type: "FactSet", facts: [
        { title: "When", value: `${c.meetingStart?.slice(0, 16).replace("T", " ")} · ${durationLine}` },
        { title: "Attendees", value: attendeeList },
        { title: "Transcript", value: transcriptLine },
      ]},
      { type: "TextBlock", text: matchLine, wrap: true, spacing: "Small" },
      { type: "TextBlock", text: '_Open Claude and say: "Log the engagement for this meeting"_', size: "Small", isSubtle: true, wrap: true, spacing: "Medium" },
    ];

    try {
      await postAdaptiveCard({
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.4",
        body,
        actions: [
          { type: "Action.OpenUrl", title: "Open Claude", url: "https://claude.ai" },
        ],
      }, progress);
      posted++;
    } catch (err) {
      progress(`⚠️ Failed to post card for "${c.meetingSubject}": ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  progress(`✅ Posted ${posted}/${candidates.length} meeting notification(s) to Teams`);
  return posted;
}
