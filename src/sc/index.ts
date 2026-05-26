import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { readFileSync, existsSync, writeFileSync } from "fs";
import { homedir } from "os";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { execFileSync } from "child_process";
import { DYNAMICS_HOST, alfredConfig as _baseConfig, ALL_ENGAGEMENT_TYPES } from "../config.js";
import { requireGuid, makeProgress, WriteRateLimiter, regenerateAlfredApp, externalData } from "../shared.js";
import {
  fetchOpportunities,
  fetchOpportunityById,
  fetchEngagementsByOpportunity,
  fetchEngagementById,
  createEngagement,
  updateEngagement,
  buildDescription,
  listTimelineNotes,
  deleteTimelineNote,
  fetchAccountById,
  searchAccounts,
  addAttendeesToEngagement,
  addSelfToEngagement,
  addCollabTeamToEngagement,
  deleteEngagement,
  fetchCollaborationTeam,
  searchContacts,
  resolveOpportunityId,
  getForecastSummary,
  searchSystemUsers,
  type EngagementType,
  type OpportunityFilter,
  type EngagementDescription,
} from "../tools/dynamicsClient.js";
import { getCalendarEvents } from "../tools/outlookClient.js";
import { registerCommonTools } from "../common-tools.js";

const DYNAMICS_BASE_URL = DYNAMICS_HOST;

// ---------------------------------------------------------------------------
// Security helpers
// ---------------------------------------------------------------------------

const engagementWriteLimiter = new WriteRateLimiter(10, 10 * 60 * 1000); // 10 per 10 min
const deleteWriteLimiter      = new WriteRateLimiter(3,  10 * 60 * 1000); // 3 per 10 min

// ---------------------------------------------------------------------------
// User config — loaded from shared config.ts
// ---------------------------------------------------------------------------
const alfredConfig = _baseConfig;
const isSSC     = alfredConfig.role === "ssc";
const isManager = alfredConfig.role === "manager";

// Which engagement types this user has enabled (defaults to all if not configured)
const activeEngagementTypes: string[] = alfredConfig.engagementTypes?.length
  ? alfredConfig.engagementTypes
  : [...ALL_ENGAGEMENT_TYPES];

process.stderr.write(
  `[alfred] Config loaded — role: ${alfredConfig.role ?? "sc"}, ` +
  `engagement types: ${activeEngagementTypes.join(", ")}\n`
);

type Engagement = import("../tools/dynamicsClient.js").Engagement;

function engagementLink(e: Engagement): string | null {
  const id = e.sn_engagementid ?? "";
  return id ? `${DYNAMICS_BASE_URL}/main.aspx?etn=sn_engagement&id=${id}&pagetype=entityrecord` : null;
}

function engagementSummary(e: Engagement, action: "Created" | "Updated"): string {
  const link = engagementLink(e);
  const lines = [
    `✅ ${action}: **${e.sn_name}** (${e.sn_engagementnumber ?? e.sn_engagementid ?? "—"})`,
    `Type: ${e.engagementTypeName ?? "—"}`,
    `Status: ${e.statuscode === 876130000 ? "Cancelled" : e.statecode === 0 ? "Open" : "Complete"}`,
    ...(e.sn_completeddate ? [`Completed: ${e.sn_completeddate.slice(0, 10)}`] : []),
    ...(link ? [`\n🔗 Open in Dynamics: ${link}`] : []),
    ...(e.sn_description ? [`\n${e.sn_description}`] : []),
  ];
  return lines.join("\n");
}

const ENGAGEMENT_TYPES = ALL_ENGAGEMENT_TYPES;

const server = new McpServer({
  name: "sc-engagement-mcp",
  version: "1.0.0",
});

// Register the 37 shared tools (identical/cosmetic between SC and Sales)
registerCommonTools(server, "sc");

// ---------------------------------------------------------------------------
// Tool: list_opportunities
// ---------------------------------------------------------------------------
server.tool(
  "list_opportunities",
  isSSC
    ? `List open opportunities from Dynamics 365.

This user is an SSC (Sales Support Consultant) — they do not have an assigned pipeline in Dynamics. Default to searching across ALL opportunities (my_opportunities_only=false). If the user says "show MY opportunities" or "my pipeline", set my_opportunities_only=true — this filters to opportunities where they are on the collaboration team.

IMPORTANT: Before calling this tool, always ask:
1. "Which account or opportunity are you looking for?" (or "your collaboration team opps?" if they said "my")
2. "100K+ NNACV only, or all sizes?" (default: 100K+ only)

NOTE: $0 NNACV opportunities are excluded by default (noise). If the user explicitly asks for $0 deals, set include_zero_value=true.

DISPLAY: Always show the nnacv field as the primary deal value (labelled "NNACV"). Never show totalamount as the deal size — it is ACV (full contract value including renewals) and inflates pipeline figures. If the user asks about ACV specifically, show totalamount labelled as "ACV" alongside NNACV.`
    : isManager
    ? `List open opportunities from Dynamics 365.

This user is an SC Manager — they want to see their team's pipeline, not just their own. Always search across ALL opportunities (my_opportunities_only=false). Their territory filter is applied automatically when my_opportunities_only=true, but they may want to search by SC name or account instead.

IMPORTANT: Before calling this tool, always ask:
1. "Your whole team's pipeline, a specific SC, or a specific account?"
   — If a specific SC or account is named, pass it as the search field
2. "100K+ NNACV only, or all sizes?" (default: 100K+ only)

NOTE: $0 NNACV opportunities are excluded by default (noise). If the user explicitly asks for $0 deals, set include_zero_value=true.

DISPLAY: Always show the nnacv field as the primary deal value (labelled "NNACV"). Never show totalamount as the deal size — it is ACV (full contract value including renewals) and inflates pipeline figures. If the user asks about ACV specifically, show totalamount labelled as "ACV" alongside NNACV.`
    : `List open opportunities from Dynamics 365.

Defaults to the current user's pipeline only (SC or territory). Only set my_opportunities_only=false if the user explicitly asks for all opportunities, a colleague's pipeline, a region, or a manager view.

IMPORTANT: Before calling this tool, always ask the user these two questions if they haven't specified:
1. "100K+ NNACV only, or all sizes?" (default: 100K+ only)
2. "All your accounts, or a specific account?" (default: all — if they name one, pass it as search)

Ask both together in one message. Only call this tool once you have their answers.

NOTE: $0 NNACV opportunities are excluded by default (noise). If the user explicitly asks for $0 deals, set include_zero_value=true.

DISPLAY: Always show the nnacv field as the primary deal value (labelled "NNACV"). Never show totalamount as the deal size — it is ACV (full contract value including renewals) and inflates pipeline figures. If the user asks about ACV specifically, show totalamount labelled as "ACV" alongside NNACV.

CROSS-REFERENCE: After presenting pipeline results, compare with the Data_Analytics_Connection account_insights tool. Note: Dynamics data is live CRM state; Data Analytics is data lake (may lag by up to 24h). Flag any discrepancies between the two sources.`,
  {
    top: z.number().optional().describe("Max number of results (default 50)"),
    search: z.string().optional().describe("Filter by opportunity or account name (partial match)"),
    min_nnacv: z.number().optional().describe("Minimum NNACV in USD — default 100000 ($100K+). Set to 0 for no filter. Negative NNACV deals are always included."),
    my_opportunities_only: z.boolean().optional().describe(
      isSSC
        ? "SSC mode — default false (search all). Set true when user says 'my opportunities' — filters to their collaboration team."
        : isManager
        ? "Manager mode — default false (search all/team). Set true to use territory filter for the full team view."
        : "Filter to current user's owned opportunities only — default true."
    ),
    include_closed: z.boolean().optional().describe("Include won/lost/closed opportunities — default false (open only). Set true when user asks about a specific opp by OPTY number or explicitly wants closed deals."),
    include_zero_value: z.boolean().optional().describe("Include $0 NNACV opportunities — default false (excluded as noise). Set true only if user explicitly asks for $0 deals."),
    colleague_name: z.string().optional().describe("Filter by a colleague's name (partial match) — returns all opportunities where they appear as either AE (owner) or SC. Use for coverage/backup scenarios, e.g. 'show me Stéphane's pipeline'."),
  },
  async ({ top, search, min_nnacv, my_opportunities_only, include_closed, include_zero_value, colleague_name }) => {
    const progress = makeProgress(server);
    const myOpps = my_opportunities_only ?? (isSSC || isManager ? false : true);
    const filter: OpportunityFilter = {
      top,
      search,
      minNnacv: min_nnacv ?? 100000,
      myOpportunitiesOnly: colleague_name ? false : myOpps,
      // SSC: filter by collaboration team membership (not SC field which they don't have)
      ...(isSSC && myOpps && !colleague_name ? { myOppsFilterField: "collab" as const } : {}),
      includeClosed: include_closed ?? false,
      includeZeroValue: include_zero_value ?? false,
      colleagueSearch: colleague_name,
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
  `Get a single opportunity by its Dynamics GUID or OPTY number (e.g. OPTY5328326).

After fetching the opportunity, ALWAYS enrich it by calling the account_insights MCP tool with:
"Show current subscriptions, license utilization, and renewal data for [accountName]"

Then present a combined summary:
- What they're buying (the opportunity)
- What they already own (products + seats purchased)
- How much they're using (utilization % and used/total seats)
- Deal type inference: upsell (expanding existing product), cross-sell (new product line), or new logo
Example output: "Acme Corp has CSM Pro — 600/1400 seats used (43%). This TPSM opportunity is an upsell."

DISPLAY: Show both values clearly labelled — "NNACV: $X | ACV: $Y". NNACV (nnacv field) is the primary metric. ACV (totalamount) is the full contract value and should always be secondary. Never present totalamount as "deal value" without the ACV label.`,
  { opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number (e.g. OPTY5328326)") },
  async ({ opportunity_id }) => {
    const progress = makeProgress(server);
    const id = await resolveOpportunityId(opportunity_id, progress);
    const opp = await fetchOpportunityById(id, progress);
    const link = `${DYNAMICS_BASE_URL}/main.aspx?etn=opportunity&pagetype=entityrecord&id=${opp.opportunityid}`;

    // Fetch collaboration team, engagements, and timeline notes in parallel
    const [collabTeam, engagements, notes] = await Promise.all([
      fetchCollaborationTeam(id, progress).catch(() => []),
      fetchEngagementsByOpportunity(id, progress).catch(() => []),
      listTimelineNotes(id, progress).catch(() => []),
    ]);

    let extraSections = "";

    // Collaboration team
    if (collabTeam.length > 0) {
      extraSections += "\n\n--- Collaboration Team ---\n" +
        collabTeam.map(m =>
          `${m.isPrimary ? "⭐ " : ""}${m.userName} — ${m.collaborationRole}${m.jobRole ? ` (${m.jobRole})` : ""}`
        ).join("\n");
    }

    // Engagements summary
    if (engagements.length > 0) {
      extraSections += "\n\n--- Engagements ---\n" +
        engagements.map(e =>
          `${e.statusName === "Active" ? "🟢" : e.statusName === "Completed" ? "✅" : "⬜"} ${e.engagementTypeName ?? "?"} — ${e.statusName ?? "?"}${e.sn_completeddate ? ` (${e.sn_completeddate.slice(0, 10)})` : ""}`
        ).join("\n");
    } else {
      extraSections += "\n\n--- No engagements on this opportunity ---";
    }

    // Timeline notes
    if (notes.length > 0) {
      extraSections += "\n\n--- Opportunity Timeline Notes ---\n" +
        notes.map(n =>
          `[${n.createdon?.slice(0, 10) ?? "—"}] ${n.subject ?? "(no subject)"}\n${n.notetext ?? ""}`
        ).join("\n\n");
    } else {
      extraSections += "\n\n--- No timeline notes on this opportunity ---";
    }

    return {
      content: [{ type: "text", text: JSON.stringify({ ...opp, dynamicsLink: link }, null, 2) + extraSections }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_account
// ---------------------------------------------------------------------------
server.tool(
  "get_account",
  `Get full account details from Dynamics 365 by account ID.

Returns: industry, website, phone, employees, revenue, address, owner (AE), SC name.
Use this to understand the customer context before creating engagements or reviewing deals.

After fetching, ALSO call account_insights with the account name to get subscription/utilization data.`,
  { account_id: z.string().describe("Dynamics account GUID") },
  async ({ account_id }) => {
    const id = requireGuid(account_id, "account_id");
    const progress = makeProgress(server);
    const account = await fetchAccountById(id, progress);
    return { content: [{ type: "text", text: JSON.stringify(account, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: search_accounts
// ---------------------------------------------------------------------------
server.tool(
  "search_accounts",
  "Search Dynamics 365 accounts by name — useful to find an account ID or get account details when you only have the name.",
  { name: z.string().describe("Account name or partial name to search for") },
  async ({ name }) => {
    const progress = makeProgress(server);
    const accounts = await searchAccounts(name, progress);
    return { content: [{ type: "text", text: JSON.stringify(accounts, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: search_contacts
// ---------------------------------------------------------------------------
server.tool(
  "search_contacts",
  `Search Dynamics 365 contacts by name or email. Returns job title, email, phone, and account.

Use this to enrich the stakeholders field with external customer titles (e.g. "Carlo Tamburrini, VP IT Operations").
Optionally filter by account_id to scope results to a specific customer.`,
  {
    query: z.string().describe("Contact name or email to search for"),
    account_id: z.string().optional().describe("Optional account GUID to scope results to a specific customer"),
  },
  async ({ query, account_id }) => {
    if (account_id) requireGuid(account_id, "account_id");
    const progress = makeProgress(server);
    const contacts = await searchContacts(query, { accountId: account_id }, progress);
    if (contacts.length === 0) {
      return { content: [{ type: "text", text: "No contacts found." }] };
    }
    const text = contacts.map(c =>
      `**${c.fullname}**${c.jobtitle ? ` — ${c.jobtitle}` : ""}` +
      `${c.emailaddress1 ? `\n  Email: ${c.emailaddress1}` : ""}` +
      `${c.telephone1 ? `\n  Phone: ${c.telephone1}` : ""}` +
      `${c.accountName ? `\n  Account: ${c.accountName}` : ""}`
    ).join("\n\n");
    return { content: [{ type: "text", text }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: create_engagement
// ---------------------------------------------------------------------------
server.tool(
  "create_engagement",
  `Create a new engagement record in Dynamics 365. Account is auto-derived from the opportunity.

**This user's engagement types:** ${activeEngagementTypes.join(", ")}
Suggest only from this list unless the user explicitly asks for a different type.

IMPORTANT: Always show the user a full summary of what will be created (name, type, use case, key points, next actions) and get explicit confirmation BEFORE calling this tool.

Always populate the structured description fields for every engagement type:
- use_case, key_points (label auto-adapts per type), next_actions, risks, stakeholders
A timeline note is created automatically on creation.

IMPORTANT — primary_product_id MUST match the linked opportunity's Business Unit / product. Never guess — use search_products and cross-check the opportunity's Business Unit List before selecting.

FORMAT: All text fields must use bullet points (• item), never prose paragraphs. Keep each bullet to one line.

STAKEHOLDERS: Internal ServiceNow people — names only, no titles. External customer contacts — include business title if known.

Do NOT append internal SC attribution (e.g. "SC: Fredrik Holmstrom") to any text field — Dynamics captures the author automatically.

BEFORE generating any content: call list_engagements on the opportunity to read existing engagement content, and get_opportunity to read the opportunity timeline. Only write what is genuinely NEW — never duplicate what is already logged.

When creating from a calendar event or meeting, always pass the attendees list — they are automatically linked as Active Participants (internal company colleagues) and Active Engagement Contacts (external customers).

⚠️ OWNER OVERRIDE (owner_name): if set, resolve the name first and show the user the EXACT full name matched. Ask "Create this under [full name]'s name?" and wait for explicit yes before setting owner_confirmed=true. Wrong ownership cannot be undone from Alfred.

AFTER EVERY SUCCESSFUL CREATE: Always present the result to the user as:
✅ [Engagement Name] (ENG#) — [Open in Dynamics](link)
The Dynamics link is in the tool response. This applies to EVERY engagement created, including bulk runs. Never omit the link.`,
  {
    opportunity_id: z.string().describe("Dynamics opportunity GUID"),
    primary_product_id: z.string().describe("Dynamics product GUID (use search_products to find it)"),
    name: z.string().describe("Short engagement name / subject"),
    type: z.enum(ENGAGEMENT_TYPES).describe("Engagement type"),
    completed_date: z.string().optional().describe("ISO date when engagement was completed, e.g. 2026-03-16"),
    // Structured description (applies to all engagement types)
    use_case: z.string().optional().describe("Use case name (e.g. ICW, ITSM)"),
    key_points: z.array(z.string()).optional().describe("Key points — label auto-adapts per type (e.g. 'Milestones achieved' for Tech Win, 'Objectives identified' for Discovery, 'Demo delivered' for Demo)"),
    secondary_points: z.array(z.string()).optional().describe("Type-specific secondary bullets — auto-labelled per type: Discovery='Key questions uncovered', Demo/Workshop/EBC='Customer reactions/feedback', Business Case='Quantified benefits', POV='Customer results', CBR='Action items agreed'"),
    submission_date: z.string().optional().describe("RFx only — submission/due date for the RFP/RFI"),
    next_actions: z.array(z.string()).optional().describe("List of next actions to complete"),
    risks: z.string().optional().describe("Risks or help required (use '-' if none)"),
    stakeholders: z.string().optional().describe("Stakeholders — use search_contacts to look up external customer titles. Internal SN people: names only, no titles."),
    // Plain text fallback
    notes: z.string().optional().describe("Plain text description (used only if structured fields are not provided)"),
    // Attendees — pass from the calendar event, linked automatically after creation
    attendees: z.array(z.object({
      name:  z.string(),
      email: z.string(),
    })).optional().describe("Meeting attendees from the calendar event. Always pass these when creating from a meeting — internal (company email addresses) become Active Participants, external become Active Engagement Contacts."),
    include_self: z.boolean().optional().describe("When true, adds the current logged-in user as an Active Participant on the new engagement."),
    include_collaboration_team: z.boolean().optional().describe("When true, automatically fetches the opportunity's collaboration team and adds all members as Active Participants on the new engagement."),
    owner_name: z.string().optional().describe("Name of the SC who should own this engagement (partial match). Use ONLY when explicitly asked to create on behalf of a colleague. Leave unset to own it yourself."),
    owner_confirmed: z.boolean().optional().describe("REQUIRED when owner_name is set. Must be explicitly true AFTER showing the user the resolved full name and getting their explicit approval. Never set this without user confirmation of the exact name shown."),
    confirmed: z.boolean().optional().describe("MUST be true to actually create. Omit or set false to get a dry-run preview first. Always preview before creating."),
  },
  async ({ opportunity_id, primary_product_id, name, type, completed_date, use_case, key_points, secondary_points, submission_date, next_actions, risks, stakeholders, notes, attendees, include_self, include_collaboration_team, owner_name, owner_confirmed, confirmed }) => {
    requireGuid(opportunity_id, "opportunity_id");
    requireGuid(primary_product_id, "primary_product_id");

    const progress = makeProgress(server);

    // Resolve owner early — needed in both preview and creation paths
    const ownerUsers = owner_name ? await searchSystemUsers(owner_name, progress) : [];
    const resolvedOwner = ownerUsers[0];

    // Dry-run: return a preview without writing anything
    if (!confirmed) {
      const opp = await fetchOpportunityById(opportunity_id, progress);
      const ownerLine = owner_name
        ? resolvedOwner
          ? `**⚠️ Owner override:** This will be created under **${resolvedOwner.fullname}**'s name — NOT yours. Verify this is correct before confirming.`
          : `**⚠️ Owner override requested ("${owner_name}") but no matching user found** — will create under your own name.`
        : `**Owner:** You (current user)`;
      return {
        content: [{ type: "text", text:
          `📋 **Dry-run preview — nothing has been created yet.**\n\n` +
          `**Engagement:** ${name}\n` +
          `**Type:** ${type}\n` +
          `**Opportunity:** ${opp.name}\n` +
          `**Account:** ${opp.accountName ?? "—"}\n` +
          `**Completed date:** ${completed_date ?? "not set"}\n` +
          `**Attendees to link:** ${attendees?.length ?? 0}\n` +
          `**Include self:** ${include_self ? "Yes" : "No"}\n` +
          `**Include collaboration team:** ${include_collaboration_team ? "Yes" : "No"}\n` +
          `${ownerLine}\n\n` +
          (owner_name && resolvedOwner
            ? `⚠️ **Creating under someone else's name is irreversible from Alfred** — the owner must reassign it back manually in Dynamics if incorrect.\n\nCall again with \`confirmed: true\` AND \`owner_confirmed: true\` to create this engagement.`
            : `Call again with \`confirmed: true\` to create this engagement.`)
        }],
      };
    }

    // Bug 16: duplicate check BEFORE owner_confirmed gate — no point asking for confirmation
    // if Dynamics would reject the create anyway
    const existingEngs = await fetchEngagementsByOpportunity(opportunity_id, progress);
    const dupEng = existingEngs.find(e => e.engagementTypeName === type);
    if (dupEng) {
      const isCancelled = dupEng.statusName?.toLowerCase().includes("cancel");
      return {
        content: [{ type: "text", text: isCancelled
          ? `❌ A cancelled ${type} already exists on this opportunity: **${dupEng.sn_name}** (${dupEng.sn_engagementnumber ?? dupEng.sn_engagementid}). Reopen the existing one instead of creating a duplicate.`
          : `❌ A ${type} engagement already exists on this opportunity: **${dupEng.sn_name}** (${dupEng.sn_engagementnumber ?? dupEng.sn_engagementid}). Collaborate on the existing one instead of creating a duplicate.`
        }],
      };
    }

    // Safety gate: block creation under another user's name until owner_confirmed is explicitly true
    if (owner_name && resolvedOwner && !owner_confirmed) {
      return {
        content: [{ type: "text", text:
          `🛑 **Owner confirmation required.**\n\n` +
          `This engagement will be created under **${resolvedOwner.fullname}** (${resolvedOwner.internalemailaddress ?? resolvedOwner.systemuserid})'s name.\n\n` +
          `⚠️ This cannot be undone from Alfred — ${resolvedOwner.fullname} would need to manually reassign it in Dynamics if incorrect.\n\n` +
          `Please confirm with the user: "Create this engagement under ${resolvedOwner.fullname}'s name?" — then retry with \`owner_confirmed: true\`.`
        }],
      };
    }

    engagementWriteLimiter.check("create_engagement");
    progress(`🎯 Creating engagement: "${name}" (${type})`);
    const opp = await fetchOpportunityById(opportunity_id, progress);
    progress(`🏢 Account resolved: ${opp.accountName}`);

    const ownerId = resolvedOwner?.systemuserid;
    if (owner_name && !ownerId) {
      progress(`⚠️ No user found for "${owner_name}" — creating under your own name`);
    } else if (ownerId) {
      progress(`👤 Owner: ${resolvedOwner!.fullname}`);
    }

    const desc: EngagementDescription = { engagementType: type as EngagementType, useCase: use_case, keyPoints: key_points, secondaryPoints: secondary_points, submissionDate: submission_date, nextActions: next_actions, risks, stakeholders };
    const hasStructured = use_case || key_points?.length || secondary_points?.length || next_actions?.length || stakeholders;
    const finalNotes = hasStructured ? buildDescription(desc) : notes;

    const engagement = await createEngagement({
      opportunityId: opportunity_id,
      accountId: opp.accountid,
      primaryProductId: primary_product_id,
      name,
      type: type as EngagementType,
      notes: finalNotes,
      completedDate: completed_date,
      ownerId,
    }, progress);

    // Auto-link attendees if provided
    if (attendees?.length && engagement.sn_engagementid) {
      progress(`👥 Linking ${attendees.length} attendee(s)...`);
      await addAttendeesToEngagement(engagement.sn_engagementid, attendees, progress);
    }

    // Auto-link self if requested
    if (include_self && engagement.sn_engagementid) {
      await addSelfToEngagement(engagement.sn_engagementid, progress);
    }

    // Auto-link collaboration team members if requested
    if (include_collaboration_team && engagement.sn_engagementid) {
      progress("👥 Adding opportunity collaboration team as participants...");
      await addCollabTeamToEngagement(engagement.sn_engagementid, opportunity_id, progress);
    }

    // Check which SC required milestones are still missing after this creation
    const SC_REQUIRED = ["Discovery", "Demo", "Technical Win"];
    const allEngagements = await fetchEngagementsByOpportunity(opportunity_id, progress);
    const activeTypeNames = allEngagements
      .filter(e => !e.statusName?.toLowerCase().includes("cancel"))
      .map(e => e.engagementTypeName ?? "")
      .filter(Boolean);
    const stillMissing = SC_REQUIRED.filter(t => !activeTypeNames.includes(t));

    let text = engagementSummary(engagement, "Created");
    if (stillMissing.length > 0) {
      text += `\n\n⚠️ **Missing SC milestones on this opp:** ${stillMissing.join(", ")}. Want me to create ${stillMissing.length === 1 ? "one" : "them"} now?`;
    } else {
      text += `\n\n✅ All 3 SC milestones (Discovery, Demo, Technical Win) are now logged on this opp.`;
    }

    return {
      content: [{ type: "text", text }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: add_engagement_attendees
// ---------------------------------------------------------------------------
server.tool(
  "add_engagement_attendees",
  `Add meeting attendees to an engagement in Dynamics 365.

- Internal attendees (company email addresses) are added as Active Participants (sn_engagementassignee → systemuser)
- External attendees (customers) are added as Active Engagement Contacts (sn_engagementcontact → contact)
- Attendees not found in Dynamics are reported but do not cause failure

Use this after creating or updating an engagement when you have a list of meeting attendees from the calendar event.`,
  {
    engagement_id: z.string().describe("Dynamics sn_engagement GUID"),
    attendees: z.array(z.object({
      name:  z.string().describe("Attendee display name"),
      email: z.string().describe("Attendee email address"),
    })).describe("List of meeting attendees — split automatically into internal participants and external contacts"),
  },
  async ({ engagement_id, attendees }) => {
    engagementWriteLimiter.check("add_engagement_attendees");
    const id = requireGuid(engagement_id, "engagement_id");
    const progress = makeProgress(server);
    progress(`👥 Adding ${attendees.length} attendee(s) to engagement ${id}...`);

    const results = await addAttendeesToEngagement(id, attendees, progress);

    const participants = results.filter(r => r.type === "participant");
    const contacts     = results.filter(r => r.type === "contact");
    const notFound     = results.filter(r => r.type === "not_found");

    const lines = [
      `**Attendees added to engagement**`,
      participants.length ? `👤 Participants (internal): ${participants.map(r => r.name).join(", ")}` : "",
      contacts.length     ? `🤝 Contacts (external): ${contacts.map(r => r.name).join(", ")}` : "",
      notFound.length     ? `⚠️ Not found in Dynamics: ${notFound.map(r => r.email).join(", ")}` : "",
    ].filter(Boolean);

    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: update_engagement
// ---------------------------------------------------------------------------
server.tool(
  "update_engagement",
  `Update an existing engagement record in Dynamics 365.

IMPORTANT: Always show the user exactly what will change (field by field) and get explicit confirmation BEFORE calling this tool.

To REOPEN a completed engagement, set mark_complete=false — this PATCHes statecode=0, statuscode=1.

Always use the structured description fields to keep the description current (applies to all engagement types).
A timeline_title + timeline_text should always be provided to log what changed.

IMPORTANT — primary_product_id MUST match the linked opportunity's Business Unit / product. Never guess — use search_products and cross-check against the opportunity before setting.

FORMAT: timeline_text and all text fields must use bullet points (• item), never prose paragraphs. Keep each bullet to one line.

BEFORE generating any content: read the existing engagement (get_engagement), list_engagements on the opportunity, and get_opportunity timeline. Only write what is genuinely NEW — never duplicate what is already logged.`,
  {
    engagement_id: z.string().describe("Dynamics sn_engagement GUID"),
    name: z.string().optional().describe("Updated engagement name"),
    type: z.enum(ENGAGEMENT_TYPES).optional().describe("Updated engagement type"),
    primary_product_id: z.string().optional().describe("Updated primary product GUID (sn_productfamilies) — use search_products to find the correct GUID"),
    completed_date: z.string().optional().describe("Updated completed date (ISO format e.g. 2026-03-16)"),
    mark_complete: z.boolean().optional().describe("Set to true to mark Complete. Set to false to REOPEN a completed engagement (sets statecode=0, statuscode=1)."),
    // Structured description fields (all types)
    use_case: z.string().optional().describe("Use case name"),
    key_points: z.array(z.string()).optional().describe("Full updated key points list — label auto-adapts per engagement type"),
    secondary_points: z.array(z.string()).optional().describe("Type-specific secondary bullets — Discovery='Key questions uncovered', Demo/Workshop/EBC='Customer reactions/feedback', Business Case='Quantified benefits', POV='Customer results', CBR='Action items agreed'"),
    submission_date: z.string().optional().describe("RFx only — submission/due date"),
    next_actions: z.array(z.string()).optional().describe("Full updated list of next actions"),
    risks: z.string().optional().describe("Risks or help required"),
    stakeholders: z.string().optional().describe("Stakeholders — use search_contacts for external titles. Internal SN: names only."),
    notes: z.string().optional().describe("Plain text description (only if structured fields not used)"),
    // Timeline note
    timeline_title: z.string().optional().describe("Title for the timeline note. Use the engagement event date (completed_date or known date), NOT today. Format: '{Type} – {date}' for updates, e.g. 'Discovery – 2026-05-14'. Say 'updated', never 'created'."),
    timeline_text: z.string().optional().describe("Body text for the timeline note"),
  },
  async ({ engagement_id, name, type, primary_product_id, completed_date, mark_complete, use_case, key_points, secondary_points, submission_date, next_actions, risks, stakeholders, notes, timeline_title, timeline_text }) => {
    engagementWriteLimiter.check("update_engagement");
    const id = requireGuid(engagement_id, "engagement_id");
    const progress = makeProgress(server);
    const desc: EngagementDescription = { engagementType: type as EngagementType | undefined, useCase: use_case, keyPoints: key_points, secondaryPoints: secondary_points, submissionDate: submission_date, nextActions: next_actions, risks, stakeholders };
    const hasStructured = use_case || key_points?.length || secondary_points?.length || next_actions?.length || stakeholders;
    const updated = await updateEngagement(id, {
      name,
      type: type as EngagementType | undefined,
      primaryProductId: primary_product_id,
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
// ---------------------------------------------------------------------------
// Tool: delete_timeline_note
// ---------------------------------------------------------------------------
server.tool(
  "delete_timeline_note",
  "Delete a specific timeline note by its annotation ID (use list_timeline_notes to find IDs). IMPORTANT: Always show the user the note subject and text, ask for confirmation, then ask AGAIN 'Are you sure? This cannot be undone.' Only call this tool after two explicit confirmations.",
  {
    annotation_id: z.string().describe("Dynamics annotation GUID"),
    confirmed: z.boolean().optional().describe("MUST be true to actually delete. Omit to do a dry-run preview first. Always preview before deleting — this is irreversible."),
  },
  async ({ annotation_id, confirmed }) => {
    requireGuid(annotation_id, "annotation_id");
    if (!confirmed) {
      return { content: [{ type: "text", text:
        `📋 **Dry-run — nothing deleted yet.**\n\nAnnotation ID: \`${annotation_id}\`\n\nCall again with \`confirmed: true\` to permanently delete this note.`
      }] };
    }
    deleteWriteLimiter.check("delete_timeline_note");
    const progress = makeProgress(server);
    await deleteTimelineNote(annotation_id, progress);
    return { content: [{ type: "text", text: `✅ Timeline note ${annotation_id} deleted.` }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_forecast_summary
// ---------------------------------------------------------------------------
server.tool(
  "get_forecast_summary",
  `Get a forecast summary — Committed / Best Case / Pipeline breakdown with NNACV totals.

Shows deal counts and values by forecast category, plus at-risk alerts and quarterly timing.
Use quarter param (e.g. "Q2 2026") to filter by close date within a specific quarter.

Always show the nnacv field as the primary deal value. Display format: NNACV: $X | ACV: $Y`,
  {
    quarter: z.string().optional().describe('Filter to a specific quarter, e.g. "Q2 2026"'),
    account_name: z.string().optional().describe("Filter by account name (partial match)"),
    my_opps_only: z.boolean().optional().describe("Filter to your own opps (default true)"),
  },
  async ({ quarter, account_name, my_opps_only }) => {
    const progress = makeProgress(server);
    const forecast = await getForecastSummary({
      myOppsOnly: my_opps_only ?? true,
      myOppsFilterField: isSSC ? "collab" : "owner",
      accountSearch: account_name,
      quarter,
    }, progress);

    const lines: string[] = [
      `## Forecast Summary${quarter ? ` — ${quarter}` : ""}`,
      "",
      `| Category | Opps | NNACV |`,
      `|----------|------|-------|`,
      `| **Committed** | ${forecast.byCategory.find(c => c.categoryCode === 100000003)?.count ?? 0} | $${forecast.committed.toLocaleString()} |`,
      `| **Best Case** | ${forecast.byCategory.find(c => c.categoryCode === 100000002)?.count ?? 0} | $${forecast.bestCase.toLocaleString()} |`,
      `| **Pipeline** | ${forecast.byCategory.find(c => c.categoryCode === 100000001)?.count ?? 0} | $${forecast.pipeline.toLocaleString()} |`,
      `| **Omitted** | ${forecast.byCategory.find(c => c.categoryCode === 100000004)?.count ?? 0} | $${forecast.omitted.toLocaleString()} |`,
      "",
      `**Total pipeline (excl. Omitted):** $${forecast.totalPipeline.toLocaleString()} across ${forecast.oppCount} opps`,
      `**Closing this quarter:** ${forecast.closingThisQuarter} | **Next quarter:** ${forecast.closingNextQuarter}`,
      `**At risk (overdue/closing <30d/no date):** ${forecast.atRiskCount}`,
    ];

    // Top deals per category
    for (const cat of forecast.byCategory) {
      if (cat.opps.length === 0) continue;
      lines.push("", `### ${cat.category} — $${cat.nnacv.toLocaleString()}`);
      for (const o of cat.opps.slice(0, 10)) {
        const close = o.closeDate ? o.closeDate.slice(0, 10) : "no date";
        lines.push(`- **${o.name}** (${o.account}) | $${o.nnacv.toLocaleString()} | close: ${close}`);
      }
      if (cat.opps.length > 10) lines.push(`- _...and ${cat.opps.length - 10} more_`);
    }

    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: delete_engagement
// ---------------------------------------------------------------------------
server.tool(
  "delete_engagement",
  `Delete a CANCELLED engagement record from Dynamics 365. Only cancelled engagements can be deleted.

IMPORTANT: This is irreversible. Always:
1. Call with confirmed=false first — shows the engagement details and verifies it is cancelled
2. Show the user the engagement name, type, and opportunity
3. Ask explicitly: "Are you sure you want to permanently delete this? This cannot be undone."
4. Only call with confirmed=true after the user confirms a second time`,
  {
    engagement_id: z.string().describe("Dynamics sn_engagement GUID"),
    confirmed: z.boolean().optional().describe("MUST be true to actually delete. Omit or false for dry-run preview. Always preview first — deletion is irreversible."),
  },
  async ({ engagement_id, confirmed }) => {
    requireGuid(engagement_id, "engagement_id");
    const progress = makeProgress(server);

    // Always fetch first — needed for preview and for the cancelled guard
    const engagement = await fetchEngagementById(engagement_id, progress);
    const isCancelled = engagement.statuscode === 876130000 || engagement.statusName?.toLowerCase().includes("cancel");

    if (!confirmed) {
      const status = engagement.statusName ?? (isCancelled ? "Cancelled" : "Not cancelled");
      const canDelete = isCancelled ? "✅ Status is Cancelled — eligible for deletion." : "🚫 Status is NOT Cancelled — this engagement cannot be deleted.";
      return { content: [{ type: "text", text:
        `📋 **Dry-run — nothing deleted yet.**\n\n` +
        `**Engagement:** ${engagement.sn_name} (${engagement.sn_engagementnumber ?? engagement_id})\n` +
        `**Type:** ${engagement.engagementTypeName ?? "—"}\n` +
        `**Opportunity:** ${engagement.opportunityName ?? "—"}\n` +
        `**Status:** ${status}\n\n` +
        `${canDelete}\n\n` +
        `Call again with \`confirmed: true\` to permanently delete.`
      }] };
    }

    if (!isCancelled) {
      return { content: [{ type: "text", text:
        `🚫 Cannot delete — engagement status is "${engagement.statusName ?? "unknown"}", not Cancelled.\n\nOnly cancelled engagements can be deleted. Cancel it in Dynamics first if you want to remove it.`
      }] };
    }

    deleteWriteLimiter.check("delete_engagement");
    await deleteEngagement(engagement_id, progress);
    return { content: [{ type: "text", text:
      `✅ Deleted: **${engagement.sn_name}** (${engagement.sn_engagementnumber ?? engagement_id})\n` +
      `Type: ${engagement.engagementTypeName ?? "—"} · Opportunity: ${engagement.opportunityName ?? "—"}`
    }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_collaboration_team
// ---------------------------------------------------------------------------
server.tool(
  "get_collaboration_team",
  `View the full collaboration team on an opportunity — lists all SCs, Specialists,
Renewal AMs, and other collaborators with their role, job role, primary status, and access level.`,
  { opportunity_id: z.string().describe("Dynamics opportunity GUID") },
  async ({ opportunity_id }) => {
    const id = requireGuid(opportunity_id, "opportunity_id");
    const progress = makeProgress(server);
    const members = await fetchCollaborationTeam(id, progress);
    if (members.length === 0) {
      return { content: [{ type: "text", text: "No collaboration team members found on this opportunity." }] };
    }
    const lines = members.map(m =>
      `• **${m.userName}** — ${m.collaborationRole}${m.jobRole ? ` (${m.jobRole})` : ""}${m.isPrimary ? " ⭐ Primary" : ""} · ${m.accessLevel}`
    );
    return { content: [{ type: "text", text: `Collaboration Team (${members.length} members):\n\n${lines.join("\n")}` }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_calendar_events
// ---------------------------------------------------------------------------
server.tool(
  "get_calendar_events",
  `Fetch calendar events from Outlook via the debug Chrome window.

Requires the user to be logged into Outlook (outlook.cloud.microsoft.com) in the Alfred Chrome window.
No Azure registration needed — the request runs inside the already-authenticated browser tab.

IMPORTANT: Before calling this tool, ask the user:
1. "Which date range? (e.g. 'this week', 'next 2 weeks', specific dates)"
2. "Any keyword to filter by? (e.g. 'Fabrikam', 'ICW', 'standup' — or leave blank for all)"`,
  {
    start_date:  z.string().describe("Start date in ISO format, e.g. 2026-03-16"),
    end_date:    z.string().describe("End date in ISO format, e.g. 2026-03-20"),
    search:      z.string().optional().describe("Optional keyword to filter event subjects, organizer, or attendee names. ALWAYS provide this when looking for specific meetings — without it, ALL events in the range are returned."),
    top:         z.number().optional().describe("Max events to fetch from Graph API (default 100). Use 25–50 for targeted searches."),
    categories:  z.array(z.string()).optional().describe("Filter to events matching these Outlook categories (e.g. ['on-site meeting']). Applied client-side after fetch."),
  },
  async ({ start_date, end_date, search, top, categories }) => {
    const progress = makeProgress(server);
    const events = await getCalendarEvents(start_date, end_date, search, progress, top ?? 100, categories);
    // bodyPreview and id are already stripped in outlookClient — return directly
    return {
      content: [{ type: "text", text: externalData("Outlook calendar", events) }],
    };
  }
);

// ---------------------------------------------------------------------------
// Tool: assess_tech_win
// ---------------------------------------------------------------------------
server.tool(
  "assess_tech_win",
  `Fetch a Technical Win engagement from Dynamics and assess whether it meets the Definition of Done.

This tool pulls the engagement description, timeline notes, and attendees — then you MUST apply the full Tech Win coaching framework below.

**ROLE:** You are an SC Manager validating a Tech Win. Be friendly and supportive, but do not accept vagueness. If something is unknown or unconfirmed, treat it as an action item — not a blocker, but something that must be resolved.

**DEFINITION OF DONE — Tech Win requires ALL four criteria confirmed by a named customer technical champion:**

1. **Business Challenges Solved** — Our solution uniquely addresses the customer's business challenges.
2. **Use Cases Covered** — All agreed use cases have been demonstrated and validated.
3. **Architecture Requirements Met** — All architectural expectations (scalability, integrations, performance) are fulfilled.
4. **Security & Privacy Requirements Met** — All security, compliance, and privacy concerns have been addressed.

**For each criterion assess:**
- A) What was required
- B) How it was addressed
- C) Who confirmed it (name, role, when)
- 🔴/🟡/🟢 Traffic light: 🟢 = explicitly confirmed, 🟡 = partially addressed, 🔴 = missing or unknown

**Validation check:**
- Has everything been explicitly confirmed by the customer?
- By whom and how (email, meeting, workshop)?
- What is still pending?
- At least one named technical champion must be on the engagement record as primary contact.

**Confidence rating:** Score 1–5 and explain. Only 4–5 = genuine Tech Win.

**Output format:**
1. Traffic light summary of all four criteria
2. Gaps and action items (be specific — "we don't know" is an action item)
3. Confidence score with explanation
4. If Tech Win is achieved: suggest clean Description text ready to save back to Dynamics
5. If not achieved: suggested next steps to get there`,
  {
    engagement_id: z.string().optional().describe("Dynamics Technical Win engagement GUID"),
    opportunity_id: z.string().optional().describe("Opportunity GUID — Alfred will find the Technical Win engagement on this opp"),
  },
  async ({ engagement_id, opportunity_id }) => {
    const progress = makeProgress(server);

    let twEngagement: Engagement | undefined;

    if (engagement_id) {
      requireGuid(engagement_id, "engagement_id");
      twEngagement = await fetchEngagementById(engagement_id, progress);
    } else if (opportunity_id) {
      requireGuid(opportunity_id, "opportunity_id");
      const all = await fetchEngagementsByOpportunity(opportunity_id, progress);
      twEngagement = all.find(e => e.engagementTypeName === "Technical Win");
      if (!twEngagement) {
        return { content: [{ type: "text", text: "No Technical Win engagement found on this opportunity." }] };
      }
    } else {
      return { content: [{ type: "text", text: "Provide either engagement_id or opportunity_id." }] };
    }

    if (!twEngagement.sn_engagementid) {
      return { content: [{ type: "text", text: "Engagement found but has no ID — cannot fetch notes." }] };
    }

    progress("📋 Fetching timeline notes...");
    const notes = await listTimelineNotes(twEngagement.sn_engagementid, progress);

    const payload = {
      engagement: twEngagement,
      timelineNotes: notes,
    };

    return { content: [{ type: "text", text: externalData("Technical Win engagement + notes", payload) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: update_alfred
// ---------------------------------------------------------------------------
server.tool(
  "update_alfred",
  "Pull the latest Alfred update from GitHub and rebuild. Use this to update Alfred without re-running the full setup.",
  {},
  async () => {
    const progress = makeProgress(server);
    const __filename = fileURLToPath(import.meta.url);
    const installDir = join(dirname(__filename), "..", "..");  // dist/sc/index.js → root

    progress("📡 Checking for updates...");

    let gitOutput: string;
    try {
      gitOutput = execFileSync("git", ["-C", installDir, "pull", "--ff-only"], { encoding: "utf8", timeout: 30_000 });
    } catch (e: unknown) {
      const msg = (e instanceof Error ? e.message : String(e)).toLowerCase();
      const isDiverged = msg.includes("not possible to fast-forward") || msg.includes("diverged");
      const rawMsg = e instanceof Error ? e.message : String(e);
      return { content: [{ type: "text", text: isDiverged
        ? `❌ Local branch has diverged from remote — cannot fast-forward.\n\n` +
          `**Safe fix:** open Terminal and run:\n` +
          `\`\`\`bash\ncd ${installDir}\ngit pull --rebase\n\`\`\`\n\n` +
          `⚠️ Do NOT run \`git reset --hard\` — that discards local changes permanently.`
        : `❌ Git pull failed:\n\`\`\`\n${rawMsg}\n\`\`\``
      }] };
    }

    const alreadyUpToDate = gitOutput.includes("Already up to date");
    if (alreadyUpToDate) {
      return { content: [{ type: "text", text: "✅ Alfred is already up to date — no rebuild needed." }] };
    }

    progress("🔨 New version pulled — rebuilding...");

    try {
      execFileSync("npm", ["run", "build"], {
        encoding: "utf8",
        cwd: installDir,
        timeout: 60_000,
        env: { ...process.env, PATH: process.env.PATH ?? "/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin" },
      });
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      return { content: [{ type: "text", text: `❌ Build failed:\n\`\`\`\n${msg}\n\`\`\`` }] };
    }

    // Update installedVersion in config
    try {
      const newSha = execFileSync("git", ["-C", installDir, "rev-parse", "--short", "HEAD"], { encoding: "utf8", timeout: 5_000 }).trim();
      const configPath = join(homedir(), ".alfred-config.json");
      const config = existsSync(configPath) ? JSON.parse(readFileSync(configPath, "utf8")) : {};
      config.installedVersion = newSha;
      writeFileSync(configPath, JSON.stringify(config, null, 2));
    } catch (e) { process.stderr.write(`[alfred:warn] version config persist failed: ${e instanceof Error ? e.message : String(e)}\n`); }

    // Migrate crontab paths: scripts/ → setup/ (one-time after repo restructure)
    try {
      const cron = execFileSync("crontab", ["-l"], { encoding: "utf8", timeout: 5_000 });
      if (cron.includes("/scripts/hygiene-sweep.mjs") || cron.includes("/scripts/post-meeting-sweep.mjs")) {
        const fixed = cron
          .replace(/\/scripts\/hygiene-sweep\.mjs/g, "/setup/hygiene-sweep.mjs")
          .replace(/\/scripts\/post-meeting-sweep\.mjs/g, "/setup/post-meeting-sweep.mjs");
        execFileSync("crontab", ["-"], { input: fixed, timeout: 5_000 });
        progress("🔧 Migrated cron job paths (scripts/ → setup/)");
      }
    } catch (e) { process.stderr.write(`[alfred:warn] cron migration failed: ${e instanceof Error ? e.message : String(e)}\n`); }

    // Post-update migration: install Playwright Chromium + remove old Chrome launcher
    try {
      const appMsg = regenerateAlfredApp(installDir);
      if (appMsg) progress(appMsg);
    } catch (e) { process.stderr.write(`[alfred:warn] post-update migration failed: ${e instanceof Error ? e.message : String(e)}\n`); }

    return { content: [{ type: "text", text:
      `✅ **Alfred updated and rebuilt!**\n\n` +
      `**Changes pulled:**\n\`\`\`\n${gitOutput.trim()}\n\`\`\`\n\n` +
      `ℹ️ Alfred.app / Alfred.bat on your Desktop has been removed — Alfred now launches its browser automatically in the background. You don't need a Desktop launcher anymore.\n\n` +
      `⚠️ Restart Claude Desktop to load the new version.`
    }] };
  }
);

// ---------------------------------------------------------------------------
// Startup version check — cached for 24h to avoid slow git fetch on every start
// ---------------------------------------------------------------------------
let versionStatus = "";
try {
  const __fn = fileURLToPath(import.meta.url);
  const installDir = join(dirname(__fn), "..", "..");
  const localSha = execFileSync("git", ["-C", installDir, "rev-parse", "--short", "HEAD"], { encoding: "utf8", timeout: 5_000 }).trim();

  const cacheFile = join(homedir(), ".alfred-version-check");
  const CACHE_TTL = 24 * 60 * 60 * 1000; // 24 hours
  let shouldFetch = true;

  if (existsSync(cacheFile)) {
    try {
      const cache = JSON.parse(readFileSync(cacheFile, "utf8"));
      if (cache.localSha === localSha && Date.now() - cache.timestamp < CACHE_TTL) {
        shouldFetch = false;
        if (cache.versionStatus) {
          versionStatus = cache.versionStatus;
          console.error(`[alfred] ${versionStatus} (cached)`);
        } else {
          console.error(`[alfred] ✅ Up to date (${localSha}) (cached)`);
        }
      }
    } catch (e) { process.stderr.write(`[alfred:warn] version cache parse failed: ${e instanceof Error ? e.message : String(e)}\n`); }
  }

  if (shouldFetch) {
    execFileSync("git", ["-C", installDir, "fetch", "--quiet"], { timeout: 15_000 });
    const remoteSha = execFileSync("git", ["-C", installDir, "rev-parse", "--short", "origin/main"], { encoding: "utf8", timeout: 5_000 }).trim();
    if (localSha !== remoteSha) {
      const behind = execFileSync("git", ["-C", installDir, "rev-list", "--count", `${localSha}..origin/main`], { encoding: "utf8", timeout: 5_000 }).trim();
      versionStatus = `⚠️ Alfred update available — you are ${behind} commit(s) behind (${localSha} → ${remoteSha}). Tell the user to run update_alfred to get the latest version.`;
      console.error(`[alfred] ${versionStatus}`);
    } else {
      console.error(`[alfred] ✅ Up to date (${localSha})`);
    }
    try { writeFileSync(cacheFile, JSON.stringify({ localSha, versionStatus, timestamp: Date.now() })); } catch (e) { process.stderr.write(`[alfred:warn] version cache write failed: ${e instanceof Error ? e.message : String(e)}\n`); }
  }
} catch (e) { process.stderr.write(`[alfred:warn] version check failed: ${e instanceof Error ? e.message : String(e)}\n`); }

// Register prompt so Claude sees the update warning at session start
if (versionStatus) {
  server.prompt(
    "alfred_update_available",
    "IMPORTANT: Alfred has a pending update — inform the user immediately",
    () => ({
      messages: [{
        role: "user",
        content: { type: "text", text: versionStatus },
      }],
    })
  );
}

// ---------------------------------------------------------------------------
// Start server
// ---------------------------------------------------------------------------
const transport = new StdioServerTransport();
await server.connect(transport);
console.error("[sc-engagement-mcp] Server running on stdio");

// Alfred is launched on-demand when tools need it (via ensureAlfred inside getAuthCookies)
// Do NOT auto-launch at startup — avoids spawning extra Chrome windows

process.on("SIGINT", async () => {
  process.exit(0);
});
