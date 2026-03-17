import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  fetchOpportunities,
  fetchOpportunityById,
  fetchEngagementsByOpportunity,
  fetchEngagementById,
  createEngagement,
  updateEngagement,
  searchProducts,
  getProductById,
  buildDescription,
  createTimelineNote,
  listTimelineNotes,
  deleteTimelineNote,
  type EngagementType,
  type OpportunityFilter,
  type EngagementDescription,
} from "./tools/dynamicsClient.js";
import { closeBrowser, setManualCookies, ensureChromeLink, clearAuthCache } from "./auth/tokenExtractor.js";
import { getCalendarEvents, getEmails, setOutlookCookies, clearGraphTokenCache } from "./tools/outlookClient.js";
import { setTeamsWebhook, postTeamsNotification, getTeamsTranscript, getTeamsChats } from "./tools/teamsClient.js";
import { runHygieneSweep, formatHygieneReport } from "./tools/hygieneClient.js";
import { detectPostMeetingEngagements } from "./tools/postMeetingClient.js";

const DYNAMICS_BASE_URL = "https://servicenow.crm.dynamics.com";

type Engagement = import("./tools/dynamicsClient.js").Engagement;

function engagementLink(e: Engagement): string | null {
  const id = e.sn_engagementid ?? "";
  return id ? `${DYNAMICS_BASE_URL}/main.aspx?etn=sn_engagement&id=${id}&pagetype=entityrecord` : null;
}

function engagementSummary(e: Engagement, action: "Created" | "Updated"): string {
  const link = engagementLink(e);
  const lines = [
    `✅ ${action}: **${e.sn_name}** (${e.sn_engagementnumber ?? e.sn_engagementid ?? "—"})`,
    `Type: ${e.engagementTypeName ?? "—"}`,
    `Status: ${e.statecode === 0 ? "Open" : "Complete"}`,
    ...(e.sn_completeddate ? [`Completed: ${e.sn_completeddate.slice(0, 10)}`] : []),
    ...(e.sn_description ? [`\n${e.sn_description}`] : []),
    ...(link ? [`\n🔗 ${link}`] : []),
  ];
  return lines.join("\n");
}

function engagementListItem(e: Engagement): string {
  const link = engagementLink(e);
  const status = e.statecode === 0 ? "Open" : "Complete";
  const completed = e.sn_completeddate ? ` · ${e.sn_completeddate.slice(0, 10)}` : "";
  const lines = [
    `**${e.sn_name}** (${e.sn_engagementnumber ?? "—"}) · ${e.engagementTypeName ?? "—"} · ${status}${completed}`,
    ...(e.sn_description ? [e.sn_description] : []),
    ...(link ? [`🔗 ${link}`] : []),
  ];
  return lines.join("\n");
}

// Sends a log message visible in Claude Desktop's tool execution UI
function makeProgress(srv: McpServer) {
  return (msg: string) => {
    console.error(`[progress] ${msg}`);
    srv.server.sendLoggingMessage({ level: "info", data: msg });
  };
}

const ENGAGEMENT_TYPES = [
  "Business Case",
  "Customer Business Review",
  "Demo",
  "Discovery",
  "EBC",
  "Post Sale Engagement",
  "POV",
  "RFx",
  "Technical Win",
  "Workshop",
] as const;

const server = new McpServer({
  name: "sc-engagement-mcp",
  version: "1.0.0",
});

// ---------------------------------------------------------------------------
// Tool: open_chrome_debug
// ---------------------------------------------------------------------------
server.tool(
  "open_chrome_debug",
  `Launch ChromeLink (Chrome with remote debugging on port 9222) if it's not already running. Opens Dynamics, Outlook and Teams tabs automatically.

IMPORTANT: Call this tool AUTOMATICALLY — without asking the user — whenever any tool fails with an error mentioning:
- "ChromeLink not running"
- "Chrome debug port not available"
- "No page targets"
- "CDP" or "debug port"
- "stale" or "session"
- "Could not capture Graph token"
- "401" or "unauthorized"
- "not logged in" or "not logged into"

This tool also clears all cached auth tokens, so call it proactively if you suspect a stale session.

IMPORTANT AFTER CALLING THIS TOOL:
- Tell the user: "ChromeLink is open — please log into Dynamics, Outlook and Teams, then let me know when you're ready."
- STOP and wait for the user to confirm they are logged in before retrying any other tool.
- Do NOT automatically retry the original tool — the user must log in first.`,
  {},
  async () => {
    const progress = makeProgress(server);
    // Clear all token caches — ensures fresh auth after any Chrome restart
    clearAuthCache();
    clearGraphTokenCache();
    await ensureChromeLink(progress);
    return {
      content: [{ type: "text", text: "✅ ChromeLink is open. Please log into Dynamics, Outlook and Teams in the Chrome window using ServiceNow SSO, then tell me when you're ready and I'll continue." }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: provide_token  (manual fallback)
// ---------------------------------------------------------------------------
server.tool(
  "provide_cookie",
  "Manually provide Dynamics 365 session cookies. Get them from Chrome DevTools: open Dynamics → F12 → Network tab → click any request → copy the Cookie header value.",
  { cookie: z.string().describe("The full Cookie header value from a Dynamics request") },
  async ({ cookie }) => {
    setManualCookies(cookie);
    return {
      content: [{ type: "text", text: "✅ Session cookies accepted and cached. You can now use list_opportunities and other tools." }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: list_opportunities
// ---------------------------------------------------------------------------
server.tool(
  "list_opportunities",
  `List open opportunities from Dynamics 365.

Defaults to the current user's pipeline only (SC or territory). Only set my_opportunities_only=false if the user explicitly asks for all opportunities, a colleague's pipeline, a region, or a manager view.

IMPORTANT: Before calling this tool, always ask the user these two questions if they haven't specified:
1. "100K+ NNACV only, or all sizes?" (default: 100K+ only)
2. "All your accounts, or a specific account?" (default: all — if they name one, pass it as search)

Ask both together in one message. Only call this tool once you have their answers.`,
  {
    top: z.number().optional().describe("Max number of results (default 50)"),
    search: z.string().optional().describe("Filter by opportunity or account name (partial match)"),
    min_nnacv: z.number().optional().describe("Minimum NNACV in USD — default 100000 ($100K+). Set to 0 for no filter."),
    my_opportunities_only: z.boolean().optional().describe("Filter to current user's owned opportunities only — default true."),
    include_closed: z.boolean().optional().describe("Include won/lost/closed opportunities — default false (open only). Set true when user asks about a specific opp by OPTY number or explicitly wants closed deals."),
  },
  async ({ top, search, min_nnacv, my_opportunities_only, include_closed }) => {
    const progress = makeProgress(server);
    const filter: OpportunityFilter = {
      top,
      search,
      minNnacv: min_nnacv ?? 100000,
      myOpportunitiesOnly: my_opportunities_only ?? true,
      includeClosed: include_closed ?? false,
    };
    const opportunities = await fetchOpportunities(filter, progress);
    return {
      content: [{ type: "text", text: JSON.stringify(opportunities, null, 2) }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_opportunity
// ---------------------------------------------------------------------------
server.tool(
  "get_opportunity",
  `Get a single opportunity by its Dynamics ID.

After fetching the opportunity, ALWAYS enrich it by calling the account_insights MCP tool with:
"Show current subscriptions, license utilization, and renewal data for [accountName]"

Then present a combined summary:
- What they're buying (the opportunity)
- What they already own (products + seats purchased)
- How much they're using (utilization % and used/total seats)
- Deal type inference: upsell (expanding existing product), cross-sell (new product line), or new logo
Example output: "SITA has CSM Pro — 600/1400 seats used (43%). This TPSM opportunity is an upsell."`,
  { opportunity_id: z.string().describe("Dynamics opportunity GUID") },
  async ({ opportunity_id }) => {
    const progress = makeProgress(server);
    const opp = await fetchOpportunityById(opportunity_id, progress);
    return {
      content: [{ type: "text", text: JSON.stringify(opp, null, 2) }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: list_engagements
// ---------------------------------------------------------------------------
server.tool(
  "list_engagements",
  `List all engagements linked to a specific opportunity.

If account_insights data has not yet been fetched for this account, also call account_insights with:
"Show current subscriptions and license utilization for [accountName]"
This gives context on what the customer owns — useful when reviewing engagement history.`,
  { opportunity_id: z.string().describe("Dynamics opportunity GUID") },
  async ({ opportunity_id }) => {
    const progress = makeProgress(server);
    const engagements = await fetchEngagementsByOpportunity(opportunity_id, progress);
    if (engagements.length === 0) {
      return { content: [{ type: "text", text: "No engagements found for this opportunity." }] };
    }
    const text = engagements.map(engagementListItem).join("\n\n---\n\n");
    return { content: [{ type: "text", text }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: search_products
// ---------------------------------------------------------------------------
server.tool(
  "search_products",
  "Search for Dynamics products by name — use this to find the product ID before creating an engagement",
  { name: z.string().describe("Product name or partial name to search for") },
  async ({ name }) => {
    const progress = makeProgress(server);
    const products = await searchProducts(name, progress);
    return {
      content: [{ type: "text", text: JSON.stringify(products, null, 2) }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_product
// ---------------------------------------------------------------------------
server.tool(
  "get_product",
  "Look up a Dynamics product by its GUID — useful to identify a product from an existing engagement",
  { product_id: z.string().describe("Dynamics product GUID") },
  async ({ product_id }) => {
    const progress = makeProgress(server);
    const product = await getProductById(product_id, progress);
    return {
      content: [{ type: "text", text: JSON.stringify(product, null, 2) }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: create_engagement
// ---------------------------------------------------------------------------
server.tool(
  "create_engagement",
  `Create a new engagement record in Dynamics 365. Account is auto-derived from the opportunity.

IMPORTANT: Always show the user a full summary of what will be created (name, type, use case, key points, next actions) and get explicit confirmation BEFORE calling this tool.

Always populate the structured description fields for every engagement type:
- use_case, key_points (label auto-adapts per type), next_actions, risks, stakeholders
A timeline note is created automatically on creation.`,
  {
    opportunity_id: z.string().describe("Dynamics opportunity GUID"),
    primary_product_id: z.string().describe("Dynamics product GUID (use search_products to find it)"),
    name: z.string().describe("Short engagement name / subject"),
    type: z.enum(ENGAGEMENT_TYPES).describe("Engagement type"),
    completed_date: z.string().optional().describe("ISO date when engagement was completed, e.g. 2026-03-16"),
    // Structured description (applies to all engagement types)
    use_case: z.string().optional().describe("Use case name (e.g. ICW, ITSM)"),
    key_points: z.array(z.string()).optional().describe("Key points — label auto-adapts per type (e.g. 'Milestones achieved' for Tech Win, 'Objectives identified' for Discovery, 'Demo delivered' for Demo)"),
    next_actions: z.array(z.string()).optional().describe("List of next actions to complete"),
    risks: z.string().optional().describe("Risks or help required (use '-' if none)"),
    stakeholders: z.string().optional().describe("Stakeholders (e.g. 'Brent Harrison and Lucio')"),
    // Plain text fallback
    notes: z.string().optional().describe("Plain text description (used only if structured fields are not provided)"),
  },
  async ({ opportunity_id, primary_product_id, name, type, completed_date, use_case, key_points, next_actions, risks, stakeholders, notes }) => {
    const progress = makeProgress(server);
    progress(`🎯 Creating engagement: "${name}" (${type})`);
    const opp = await fetchOpportunityById(opportunity_id, progress);
    progress(`🏢 Account resolved: ${opp.accountName}`);

    const desc: EngagementDescription = { engagementType: type as EngagementType, useCase: use_case, keyPoints: key_points, nextActions: next_actions, risks, stakeholders };
    const hasStructured = use_case || key_points?.length || next_actions?.length || stakeholders;
    const finalNotes = hasStructured ? buildDescription(desc) : notes;

    const engagement = await createEngagement({
      opportunityId: opportunity_id,
      accountId: opp.accountid,
      primaryProductId: primary_product_id,
      name,
      type: type as EngagementType,
      notes: finalNotes,
      completedDate: completed_date,
    }, progress);

    return {
      content: [{ type: "text", text: engagementSummary(engagement, "Created") }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: update_engagement
// ---------------------------------------------------------------------------
server.tool(
  "update_engagement",
  `Update an existing engagement record in Dynamics 365.

IMPORTANT: Always show the user exactly what will change (field by field) and get explicit confirmation BEFORE calling this tool.

Always use the structured description fields to keep the description current (applies to all engagement types).
A timeline_title + timeline_text should always be provided to log what changed.`,
  {
    engagement_id: z.string().describe("Dynamics sn_engagement GUID"),
    name: z.string().optional().describe("Updated engagement name"),
    type: z.enum(ENGAGEMENT_TYPES).optional().describe("Updated engagement type"),
    completed_date: z.string().optional().describe("Updated completed date (ISO format e.g. 2026-03-16)"),
    mark_complete: z.boolean().optional().describe("Set to true to mark the engagement as Complete (sets statecode=1, statuscode=2)"),
    // Structured description fields (all types)
    use_case: z.string().optional().describe("Use case name"),
    key_points: z.array(z.string()).optional().describe("Full updated key points list — label auto-adapts per engagement type"),
    next_actions: z.array(z.string()).optional().describe("Full updated list of next actions"),
    risks: z.string().optional().describe("Risks or help required"),
    stakeholders: z.string().optional().describe("Stakeholders"),
    notes: z.string().optional().describe("Plain text description (only if structured fields not used)"),
    // Timeline note
    timeline_title: z.string().optional().describe("Title for the timeline note (e.g. 'Discovery update - requirements captured')"),
    timeline_text: z.string().optional().describe("Body text for the timeline note"),
  },
  async ({ engagement_id, name, type, completed_date, mark_complete, use_case, key_points, next_actions, risks, stakeholders, notes, timeline_title, timeline_text }) => {
    const progress = makeProgress(server);
    const desc: EngagementDescription = { engagementType: type as EngagementType | undefined, useCase: use_case, keyPoints: key_points, nextActions: next_actions, risks, stakeholders };
    const hasStructured = use_case || key_points?.length || next_actions?.length || stakeholders;
    const updated = await updateEngagement(engagement_id, {
      name,
      type: type as EngagementType | undefined,
      completedDate: completed_date,
      markComplete: mark_complete,
      description: hasStructured ? desc : undefined,
      notes: hasStructured ? undefined : notes,
      timelineTitle: timeline_title,
      timelineText: timeline_text,
    }, progress);
    return {
      content: [{ type: "text", text: engagementSummary(updated, "Updated") }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: list_timeline_notes
// ---------------------------------------------------------------------------
server.tool(
  "list_timeline_notes",
  "List all timeline notes (annotations) on an engagement — use this to find note IDs before deleting",
  { engagement_id: z.string().describe("Dynamics sn_engagement GUID") },
  async ({ engagement_id }) => {
    const progress = makeProgress(server);
    const notes = await listTimelineNotes(engagement_id, progress);
    return { content: [{ type: "text", text: JSON.stringify(notes, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: delete_timeline_note
// ---------------------------------------------------------------------------
server.tool(
  "delete_timeline_note",
  "Delete a specific timeline note by its annotation ID (use list_timeline_notes to find IDs). IMPORTANT: Always confirm with the user before deleting.",
  { annotation_id: z.string().describe("Dynamics annotation GUID") },
  async ({ annotation_id }) => {
    const progress = makeProgress(server);
    await deleteTimelineNote(annotation_id, progress);
    return { content: [{ type: "text", text: `✅ Timeline note ${annotation_id} deleted.` }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_engagement
// ---------------------------------------------------------------------------
server.tool(
  "get_engagement",
  "Get a single engagement record by its Dynamics ID",
  { engagement_id: z.string().describe("Dynamics sn_engagement GUID") },
  async ({ engagement_id }) => {
    const progress = makeProgress(server);
    const engagement = await fetchEngagementById(engagement_id, progress);
    return { content: [{ type: "text", text: engagementListItem(engagement) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: provide_outlook_cookie
// ---------------------------------------------------------------------------
server.tool(
  "provide_outlook_cookie",
  "Manually provide Outlook session cookies. Get them from Chrome DevTools: open outlook.office.com → F12 → Network tab → click any request → copy the Cookie header value.",
  { cookie: z.string().describe("The full Cookie header value from an Outlook Web request") },
  async ({ cookie }) => {
    setOutlookCookies(cookie);
    return {
      content: [{ type: "text", text: "✅ Outlook session cookies accepted. You can now use get_calendar_events and search_emails." }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_calendar_events
// ---------------------------------------------------------------------------
server.tool(
  "get_calendar_events",
  `Fetch calendar events from Outlook via the debug Chrome window.

Requires the user to be logged into https://outlook.office.com in the ChromeLink Chrome window.
No Azure registration needed — the request runs inside the already-authenticated browser tab.

IMPORTANT: Before calling this tool, ask the user:
1. "Which date range? (e.g. 'this week', 'next 2 weeks', specific dates)"
2. "Any keyword to filter by? (e.g. 'PMI', 'ICW', 'standup' — or leave blank for all)"`,
  {
    start_date: z.string().describe("Start date in ISO format, e.g. 2026-03-16"),
    end_date:   z.string().describe("End date in ISO format, e.g. 2026-03-20"),
    search:     z.string().optional().describe("Optional keyword to filter event subjects"),
  },
  async ({ start_date, end_date, search }) => {
    const progress = makeProgress(server);
    const events = await getCalendarEvents(start_date, end_date, search, progress);
    return {
      content: [{ type: "text", text: JSON.stringify(events, null, 2) }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: search_emails
// ---------------------------------------------------------------------------
server.tool(
  "search_emails",
  `Search or list emails from Outlook via the debug Chrome window.

Requires the user to be logged into https://outlook.office.com in the ChromeLink Chrome window.
No Azure registration needed — the request runs inside the already-authenticated browser tab.

Can search across all mail by keyword, or list a folder (inbox, sentitems, drafts).`,
  {
    search:      z.string().optional().describe("Full-text search query across all mail (e.g. 'PMI renewal', 'budget')"),
    folder:      z.string().optional().describe("Mail folder to list: 'inbox' (default), 'sentitems', 'drafts'"),
    top:         z.number().optional().describe("Max number of messages to return (default 25)"),
    unread_only: z.boolean().optional().describe("If true, return only unread messages (only applies when not searching)"),
  },
  async ({ search, folder, top, unread_only }) => {
    const progress = makeProgress(server);
    const messages = await getEmails(
      { search, folder: folder ?? "inbox", top: top ?? 25, unreadOnly: unread_only },
      progress
    );
    return {
      content: [{ type: "text", text: JSON.stringify(messages, null, 2) }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: configure_teams_webhook
// ---------------------------------------------------------------------------
server.tool(
  "configure_teams_webhook",
  "Set the Teams incoming webhook URL for notifications. Create one in Teams: channel → ... → Connectors → Incoming Webhook.",
  { webhook_url: z.string().describe("Teams incoming webhook URL") },
  async ({ webhook_url }) => {
    setTeamsWebhook(webhook_url);
    return { content: [{ type: "text", text: "✅ Teams webhook configured. Notifications will post to that channel." }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: post_teams_notification
// ---------------------------------------------------------------------------
server.tool(
  "post_teams_notification",
  "Post a notification to the configured Teams channel via incoming webhook. Requires configure_teams_webhook to be set up first.",
  {
    title: z.string().describe("Notification title"),
    body:  z.string().describe("Notification body (supports markdown)"),
  },
  async ({ title, body }) => {
    const progress = makeProgress(server);
    await postTeamsNotification(title, body, progress);
    return { content: [{ type: "text", text: `✅ Posted to Teams: "${title}"` }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_teams_transcript
// ---------------------------------------------------------------------------
server.tool(
  "get_teams_transcript",
  `Fetch Teams meeting transcripts via Microsoft Graph. Use this to auto-generate engagement descriptions from recorded meetings.

Requires ChromeLink to be running with Teams or Outlook open. The Graph token is captured automatically.`,
  {
    search:     z.string().optional().describe("Keyword to match meeting subject (e.g. 'PMI ICW')"),
    start_date: z.string().optional().describe("Search from this date (ISO, e.g. 2026-01-01) — defaults to 30 days ago"),
    end_date:   z.string().optional().describe("Search to this date (ISO) — defaults to today"),
    meeting_id: z.string().optional().describe("If already known, fetch transcript for this specific meeting ID directly"),
  },
  async ({ search, start_date, end_date, meeting_id }) => {
    const progress = makeProgress(server);
    const transcripts = await getTeamsTranscript({ search, startDate: start_date, endDate: end_date, meetingId: meeting_id }, progress);
    return { content: [{ type: "text", text: JSON.stringify(transcripts, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_teams_chats
// ---------------------------------------------------------------------------
server.tool(
  "get_teams_chats",
  `Fetch Teams chat conversations via Microsoft Graph. Can list recent chats, search by person/topic, or fetch messages from a specific chat.

Requires ChromeLink to be running with Teams or Outlook open. The Graph token is captured automatically.

Use cases:
- "Show my recent Teams chats with PMI"
- "Get messages from my chat with John"
- "Show DMs about SITA renewal"`,
  {
    search:           z.string().optional().describe("Filter chats by topic or member name (e.g. 'PMI', 'John Smith')"),
    chat_id:          z.string().optional().describe("Fetch a specific chat by its Graph chat ID"),
    include_messages: z.boolean().optional().describe("Include recent messages for each chat (default false — list chats only)"),
    top:              z.number().optional().describe("Max chats to return (default 50)"),
  },
  async ({ search, chat_id, include_messages, top }) => {
    const progress = makeProgress(server);
    const chats = await getTeamsChats(
      { search, chatId: chat_id, includeMessages: include_messages ?? false, top },
      progress
    );
    return { content: [{ type: "text", text: JSON.stringify(chats, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: run_hygiene_sweep
// ---------------------------------------------------------------------------
server.tool(
  "run_hygiene_sweep",
  `Scan your open opportunities and flag missing SC-owned engagement milestones.

Required SC milestones: Discovery, Demo, Technical Win
Optional SC milestones: RFx, Business Case, Workshop, POV, EBC

Always runs for the current user's pipeline only. Optionally posts results to Teams.`,
  {
    post_to_teams: z.boolean().optional().describe("Post the report to Teams (requires configure_teams_webhook)"),
    min_nnacv:     z.number().optional().describe("Minimum NNACV filter in USD (default $100K)"),
  },
  async ({ post_to_teams, min_nnacv }) => {
    const progress = makeProgress(server);
    const results = await runHygieneSweep({
      postToTeams: post_to_teams ?? false,
      minNnacv: min_nnacv ?? 100_000,
    }, progress);
    return { content: [{ type: "text", text: formatHygieneReport(results) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: detect_post_meeting_engagements
// ---------------------------------------------------------------------------
server.tool(
  "detect_post_meeting_engagements",
  `Scan recently ended Teams meetings, fetch transcripts where available, and return structured candidates for engagement creation.

After calling this tool, analyse each candidate and:
1. Determine the engagement type (Discovery / Demo / Tech Win etc.) from the transcript/subject
2. Extract use_case, key_points, next_actions, risks, stakeholders
3. Confirm the matched opportunity with the user
4. Ask for approval before calling create_engagement

If a transcript is available it will be included — use it to pre-fill the engagement description.
If no transcript, use the meeting subject, attendees and calendar notes for best-effort pre-fill.`,
  {
    hours_back: z.number().optional().describe("How many hours back to scan for ended meetings (default 24)"),
    search:     z.string().optional().describe("Optional keyword to filter meeting subjects (e.g. 'PMI', 'SITA')"),
  },
  async ({ hours_back, search }) => {
    const progress = makeProgress(server);
    const candidates = await detectPostMeetingEngagements({ hoursBack: hours_back, search }, progress);

    if (candidates.length === 0) {
      return { content: [{ type: "text", text: "No ended online meetings found in the specified window." }] };
    }

    return { content: [{ type: "text", text: JSON.stringify(candidates, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Start server
// ---------------------------------------------------------------------------
const transport = new StdioServerTransport();
await server.connect(transport);
console.error("[sc-engagement-mcp] Server running on stdio");

// ChromeLink is launched on-demand when tools need it (via ensureChromeLink inside getAuthCookies)
// Do NOT auto-launch at startup — avoids spawning extra Chrome windows

process.on("SIGINT", async () => {
  await closeBrowser();
  process.exit(0);
});
