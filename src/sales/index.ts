import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { DYNAMICS_HOST, ALL_ENGAGEMENT_TYPES, alfredConfig } from "../config.js";
import { requireGuid, makeProgress, WriteRateLimiter, FORECAST_NAMES, regenerateAlfredApp, externalData } from "../shared.js";
import {
  fetchOpportunities,
  fetchOpportunityById,
  fetchEngagementsByOpportunity,
  fetchEngagementById,
  createEngagement,
  updateEngagement,
  searchAccounts,
  fetchAccountById,
  createOpportunity,
  updateOpportunity,
  searchSystemUsers,
  fetchCurrentUserId,
  createTimelineNote,
  listTimelineNotes,
  deleteTimelineNote,
  deleteEngagement,
  buildDescription,
  searchContacts,
  fetchCollaborationTeam,
  addAttendeesToEngagement,
  resolveOpportunityId,
  getForecastSummary,
  type EngagementType,
  type EngagementDescription,
} from "../tools/dynamicsClient.js";
import { getCalendarEvents } from "../tools/outlookClient.js";
import { registerCommonTools } from "../common-tools.js";
import { execFileSync } from "child_process";
import { readFileSync, existsSync, writeFileSync } from "fs";
import { homedir } from "os";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const DYNAMICS_BASE_URL = DYNAMICS_HOST;

// ---------------------------------------------------------------------------
// User config — loaded from shared config.ts
// ---------------------------------------------------------------------------
const isSalesSpecialist = alfredConfig.role === "sales_specialist";
const isSalesManager    = alfredConfig.role === "sales_manager";

process.stderr.write(
  `[alfred] Config loaded — role: ${alfredConfig.role ?? "sales"}\n`
);

const opportunityWriteLimiter = new WriteRateLimiter(10, 10 * 60 * 1000); // 10 per 10 min
const engagementWriteLimiter  = new WriteRateLimiter(10, 10 * 60 * 1000); // 10 per 10 min
const deleteWriteLimiter      = new WriteRateLimiter(3,  10 * 60 * 1000); // 3 per 10 min

const ENGAGEMENT_TYPES = ALL_ENGAGEMENT_TYPES;

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

const OPP_TYPE_CODES: Record<string, number> = {
  "new business": 1,
  "new":          1,
  "renewal":      2,
  "existing":     3,
  "existing customer": 3,
  "upsell":       3,
};

const server = new McpServer({
  name: "alfred-sales",
  version: "1.0.0",
});

registerCommonTools(server, "sales");

// ---------------------------------------------------------------------------
// Tool: search_accounts
// ---------------------------------------------------------------------------
server.tool(
  "search_accounts",
  "Search Dynamics 365 accounts by name — use this to find the account ID before creating an opportunity.",
  { name: z.string().describe("Account name or partial name") },
  async ({ name }) => {
    const progress = makeProgress(server);
    const accounts = await searchAccounts(name, progress);
    return { content: [{ type: "text", text: JSON.stringify(accounts, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_account
// ---------------------------------------------------------------------------
server.tool(
  "get_account",
  "Get full account details from Dynamics 365 by account ID.",
  { account_id: z.string().describe("Dynamics account GUID") },
  async ({ account_id }) => {
    const progress = makeProgress(server);
    requireGuid(account_id, "account_id");
    const account = await fetchAccountById(account_id, progress);
    return { content: [{ type: "text", text: JSON.stringify(account, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: search_users
// ---------------------------------------------------------------------------
server.tool(
  "search_users",
  "Search Dynamics 365 system users by name — use this to find the GUID for a Sales Rep or SC when assigning them to an opportunity.",
  { name: z.string().describe("Full or partial name, e.g. 'Fredrik' or 'Alexis'") },
  async ({ name }) => {
    const progress = makeProgress(server);
    const users = await searchSystemUsers(name, progress);
    return { content: [{ type: "text", text: JSON.stringify(users, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_my_opportunities
// ---------------------------------------------------------------------------
server.tool(
  "get_my_opportunities",
  isSalesSpecialist
    ? `List open opportunities from Dynamics 365.

This user is a Sales Specialist — they do not own opportunities directly. Default to searching across ALL opportunities (my_opportunities_only=false). If the user says "show MY opportunities" or "my pipeline", set my_opportunities_only=true — this filters to opportunities where they are on the collaboration team.

IMPORTANT: Before calling this tool, always ask:
1. "Which account or opportunity are you looking for?" (or "your collaboration team opps?" if they said "my")
2. "100K+ NNACV only, or all sizes?" (default: 100K+ only)

NOTE: $0 NNACV opportunities are excluded by default (noise). If the user explicitly asks for $0 deals, set include_zero_value=true.

DISPLAY: Always show the nnacv field as the primary deal value (labelled "NNACV"). Never show totalamount as the deal size — it is ACV (full contract value including renewals) and inflates pipeline figures. If the user asks about ACV specifically, show totalamount labelled as "ACV" alongside NNACV.

CROSS-REFERENCE: After presenting pipeline results, compare with the Data_Analytics_Connection account_insights tool. Note: Dynamics data is live CRM state; Data Analytics is data lake (may lag by up to 24h). Flag any discrepancies between the two sources.`
    : isSalesManager
    ? `List open opportunities from Dynamics 365.

This user is a Sales Manager — they want to see their team's pipeline, not just their own. Default to searching all opportunities (my_opportunities_only=false). They may want to search by AE name or account.

IMPORTANT: Before calling this tool, always ask:
1. "Your whole team's pipeline, a specific AE, or a specific account?"
   — If a specific AE or account is named, pass it as the search field
2. "100K+ NNACV only, or all sizes?" (default: 100K+ only)

NOTE: $0 NNACV opportunities are excluded by default (noise). If the user explicitly asks for $0 deals, set include_zero_value=true.

DISPLAY: Always show the nnacv field as the primary deal value (labelled "NNACV"). Never show totalamount as the deal size — it is ACV (full contract value including renewals) and inflates pipeline figures. If the user asks about ACV specifically, show totalamount labelled as "ACV" alongside NNACV.

CROSS-REFERENCE: After presenting pipeline results, compare with the Data_Analytics_Connection account_insights tool. Note: Dynamics data is live CRM state; Data Analytics is data lake (may lag by up to 24h). Flag any discrepancies between the two sources.`
    : `List your open opportunities in Dynamics 365, optionally filtered by account name or minimum NNACV.

NOTE: $0 NNACV opportunities are excluded by default (noise). If the user explicitly asks for $0 deals, set include_zero_value=true. Negative NNACV deals are always included.

DISPLAY: Always show the nnacv field as the primary deal value (labelled "NNACV"). Never show totalamount as the deal size — it is ACV (full contract value including renewals) and inflates pipeline figures. If the user asks about ACV specifically, show totalamount labelled as "ACV" alongside NNACV.

CROSS-REFERENCE: After presenting pipeline results, compare with the Data_Analytics_Connection account_insights tool. Note: Dynamics data is live CRM state; Data Analytics is data lake (may lag by up to 24h). Flag any discrepancies between the two sources.`,
  {
    search:   z.string().optional().describe("Filter by account or opportunity name"),
    min_nnacv: z.number().optional().describe("Minimum NNACV in USD — default 100000 ($100K+). Set to 0 for no filter. Negative NNACV deals are always included."),
    top: z.number().optional().describe("Max results (default 50)"),
    include_closed: z.boolean().optional().describe("Include won/lost opportunities (default false)"),
    include_zero_value: z.boolean().optional().describe("Include $0 NNACV opportunities — default false (excluded as noise). Set true only if user explicitly asks for $0 deals."),
    my_opportunities_only: z.boolean().optional().describe(
      isSalesSpecialist
        ? "Specialist mode — default false (search all). Set true when user says 'my opportunities' — filters to their collaboration team."
        : isSalesManager
        ? "Manager mode — default false (team-wide). Set true for your personal pipeline only."
        : "Filter to your owned opportunities — default true."
    ),
    colleague_name: z.string().optional().describe("Filter by a colleague's name (partial match) — returns all opportunities where they appear as either AE (owner) or SC. Use for coverage/backup scenarios, e.g. 'show me Stéphane's pipeline'."),
    close_quarter: z.string().optional().describe("Filter by closing quarter using the CRM sn_closequarter field — format '26-Q3'. When set, returns all deals in that quarter (up to 200) regardless of close date sort order. Use for dashboard/forecast views: '26-Q2' for current quarter, '26-Q3' for next."),
  },
  async ({ search, min_nnacv, top, include_closed, include_zero_value, my_opportunities_only, colleague_name, close_quarter }) => {
    const progress = makeProgress(server);
    const myOpps = my_opportunities_only ?? (isSalesSpecialist || isSalesManager ? false : true);
    const opps = await fetchOpportunities({
      search,
      minNnacv: min_nnacv,
      myOpportunitiesOnly: colleague_name ? false : myOpps,
      // Sales Specialist: filter by collaboration team; AE/Manager: filter by owner
      myOppsFilterField: isSalesSpecialist && myOpps && !colleague_name ? "collab" : "owner",
      includeClosed: include_closed ?? false,
      includeZeroValue: include_zero_value ?? false,
      top: top ?? 50,
      colleagueSearch: colleague_name,
      closeQuarter: close_quarter,
    }, progress);
    return { content: [{ type: "text", text: JSON.stringify(opps, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_opportunity
// ---------------------------------------------------------------------------
server.tool(
  "get_opportunity",
  `Get a single opportunity by its Dynamics GUID or OPTY number (e.g. OPTY5328326).

Also fetches timeline notes attached to the opportunity — essential context for engagement creation and updates.

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

    return { content: [{ type: "text", text: JSON.stringify({ ...opp, dynamicsLink: link }, null, 2) + extraSections }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: create_opportunity
// ---------------------------------------------------------------------------
server.tool(
  "create_opportunity",
  `Create a new opportunity in Dynamics 365.

IMPORTANT: Always show a full summary and get explicit confirmation BEFORE calling with confirmed=true.

Required fields: account, name, close date.
Opportunity types: "New Business", "Renewal", "Existing Customer"
Forecast categories: "Pipeline" (default), "Best Case", "Committed"

Workflow:
1. Use search_accounts to find the account ID
2. Use search_users to find SC GUID if the user wants to assign one
3. Show a dry-run summary
4. Create only after explicit confirmation`,
  {
    account_id:        z.string().describe("Dynamics account GUID (use search_accounts to find)"),
    name:              z.string().describe("Opportunity name, e.g. 'Contoso — New ITSM 2026'"),
    close_date:        z.string().describe("Expected close date, ISO format e.g. '2026-12-31'"),
    opportunity_type:  z.string().optional().describe("'New Business', 'Renewal', or 'Existing Customer' (default: New Business)"),
    forecast_category: z.string().optional().describe("'Pipeline', 'Best Case', or 'Committed' (default: Pipeline)"),
    owner_id:          z.string().optional().describe("Sales Rep systemuser GUID (defaults to current user)"),
    sc_id:             z.string().optional().describe("Solution Consultant systemuser GUID"),
    notes:             z.string().optional().describe("Additional notes or description"),
    confirmed:         z.boolean().optional().describe("MUST be true to actually create. Omit for dry-run preview first."),
  },
  async ({ account_id, name, close_date, opportunity_type, forecast_category, owner_id, sc_id, notes, confirmed }) => {
    requireGuid(account_id, "account_id");
    if (sc_id) requireGuid(sc_id, "sc_id");
    if (owner_id) requireGuid(owner_id, "owner_id");

    const progress = makeProgress(server);

    // Validate close date format and warn if in the past
    const parsedClose = new Date(close_date);
    if (isNaN(parsedClose.getTime())) {
      return { content: [{ type: "text", text: `❌ Invalid close date: "${close_date}". Use ISO format, e.g. 2026-12-31.` }] };
    }
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const closeDateWarning = parsedClose < today ? "\n\n⚠️ **Close date is in the past** — update it after creation if needed." : "";

    const typeCode = OPP_TYPE_CODES[(opportunity_type ?? "new business").toLowerCase()] ?? 1;
    const forecastMap: Record<string, number> = {
      pipeline: 100000001, "best case": 100000002, committed: 100000003,
    };
    const forecastCode = forecastMap[(forecast_category ?? "pipeline").toLowerCase()] ?? 100000001;

    if (!confirmed) {
      const account = await fetchAccountById(account_id, progress);
      return { content: [{ type: "text", text:
        `📋 **Dry-run — nothing created yet.**\n\n` +
        `**Name:** ${name}\n` +
        `**Account:** ${account.name}\n` +
        `**Close date:** ${close_date}\n` +
        `**Type:** ${opportunity_type ?? "New Business"}\n` +
        `**Forecast:** ${forecast_category ?? "Pipeline"}\n` +
        `**SC:** ${sc_id ? sc_id : "not assigned"}\n` +
        `**Notes:** ${notes ?? "—"}` +
        closeDateWarning + `\n\n` +
        `Call again with \`confirmed: true\` to create this opportunity.`
      }] };
    }

    opportunityWriteLimiter.check("create_opportunity");

    // Default owner to current user if not provided
    let resolvedOwnerId = owner_id;
    if (!resolvedOwnerId) {
      resolvedOwnerId = await fetchCurrentUserId(progress);
    }

    const opp = await createOpportunity({
      name,
      accountId: account_id,
      closeDate: close_date,
      opportunityType: typeCode,
      forecastCategory: forecastCode,
      ownerId: resolvedOwnerId,
      scId: sc_id,
      notes,
    }, progress);

    const link = `${DYNAMICS_BASE_URL}/main.aspx?etn=opportunity&pagetype=entityrecord&id=${opp.opportunityid}`;
    return { content: [{ type: "text", text:
      `✅ **Opportunity created!**\n\n` +
      `**${opp.name}** (${opp.sn_number ?? opp.opportunityid})\n` +
      `Account: ${opp.accountName}\n` +
      `Close date: ${opp.estimatedclosedate?.slice(0, 10) ?? close_date}\n` +
      `Forecast: ${FORECAST_NAMES[forecastCode] ?? "—"}\n\n` +
      `🔗 ${link}`
    }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: update_opportunity
// ---------------------------------------------------------------------------
server.tool(
  "update_opportunity",
  `Update an existing opportunity in Dynamics 365 — close date, forecast category, owner, SC, name, probability, or win/loss notes.

Always show the current values and the proposed changes, then confirm before calling with confirmed=true.`,
  {
    opportunity_id:    z.string().describe("Dynamics opportunity GUID"),
    name:              z.string().optional().describe("New opportunity name"),
    close_date:        z.string().optional().describe("New close date, ISO format"),
    forecast_category: z.string().optional().describe("'Pipeline', 'Best Case', or 'Committed'"),
    owner_id:          z.string().optional().describe("New Sales Rep systemuser GUID"),
    sc_id:             z.string().optional().describe("New SC systemuser GUID"),
    notes:             z.string().optional().describe("Updated description/notes"),
    win_loss_notes:    z.string().optional().describe("Win/loss/no-decision reason notes — log why the deal was won, lost, or stalled"),
    probability:       z.number().optional().describe("Close probability percentage (0-100)"),
    confirmed:         z.boolean().optional().describe("MUST be true to actually update. Omit for dry-run."),
  },
  async ({ opportunity_id, name, close_date, forecast_category, owner_id, sc_id, notes, win_loss_notes, probability, confirmed }) => {
    requireGuid(opportunity_id, "opportunity_id");
    if (owner_id) requireGuid(owner_id, "owner_id");
    if (sc_id)    requireGuid(sc_id, "sc_id");

    const progress = makeProgress(server);
    const current = await fetchOpportunityById(opportunity_id, progress);

    const forecastMap: Record<string, number> = {
      pipeline: 100000001, "best case": 100000002, committed: 100000003,
    };
    const forecastCode = forecast_category
      ? (forecastMap[forecast_category.toLowerCase()] ?? undefined)
      : undefined;

    if (!confirmed) {
      const changes: string[] = [];
      if (name)             changes.push(`Name: ${current.name} → ${name}`);
      if (close_date)       changes.push(`Close date: ${current.estimatedclosedate?.slice(0,10) ?? "—"} → ${close_date}`);
      if (forecast_category) changes.push(`Forecast: ${current.forecastCategoryName ?? "—"} → ${forecast_category}`);
      if (owner_id)         changes.push(`Owner: → ${owner_id}`);
      if (sc_id)            changes.push(`SC: → ${sc_id}`);
      if (notes)            changes.push(`Notes updated`);
      if (win_loss_notes)   changes.push(`Win/Loss notes: ${win_loss_notes.slice(0, 80)}${win_loss_notes.length > 80 ? "..." : ""}`);
      if (probability !== undefined) changes.push(`Probability: ${current.probability ?? "—"}% → ${probability}%`);

      return { content: [{ type: "text", text:
        `📋 **Dry-run — nothing updated yet.**\n\n` +
        `**Opportunity:** ${current.name} (${current.sn_number ?? opportunity_id})\n\n` +
        `**Proposed changes:**\n${changes.length ? changes.map(c => `• ${c}`).join("\n") : "No changes specified"}\n\n` +
        `Call again with \`confirmed: true\` to apply.`
      }] };
    }

    opportunityWriteLimiter.check("update_opportunity");

    const updated = await updateOpportunity({
      opportunityId: opportunity_id,
      name, closeDate: close_date, forecastCategory: forecastCode,
      ownerId: owner_id, scId: sc_id, notes,
      winLossNotes: win_loss_notes, probability,
    }, progress);

    const link = `${DYNAMICS_BASE_URL}/main.aspx?etn=opportunity&pagetype=entityrecord&id=${updated.opportunityid}`;
    return { content: [{ type: "text", text:
      `✅ **Opportunity updated**\n\n` +
      `**${updated.name}** (${updated.sn_number ?? opportunity_id})\n` +
      `Close date: ${updated.estimatedclosedate?.slice(0,10) ?? "—"}\n` +
      `Forecast: ${updated.forecastCategoryName ?? "—"}\n\n` +
      `🔗 ${link}`
    }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: add_opportunity_note
// ---------------------------------------------------------------------------
server.tool(
  "add_opportunity_note",
  "Add a timeline note to an opportunity in Dynamics 365.",
  {
    opportunity_id: z.string().describe("Dynamics opportunity GUID"),
    title:   z.string().describe("Note title / subject"),
    body:    z.string().optional().describe("Note body text"),
  },
  async ({ opportunity_id, title, body }) => {
    opportunityWriteLimiter.check("add_opportunity_note");
    requireGuid(opportunity_id, "opportunity_id");
    const progress = makeProgress(server);
    await createTimelineNote(opportunity_id, title, body ?? "", progress);
    return { content: [{ type: "text", text: `✅ Note added to opportunity.` }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: list_opportunity_notes
// ---------------------------------------------------------------------------
server.tool(
  "list_opportunity_notes",
  "List timeline notes on an opportunity.",
  { opportunity_id: z.string().describe("Dynamics opportunity GUID") },
  async ({ opportunity_id }) => {
    requireGuid(opportunity_id, "opportunity_id");
    const progress = makeProgress(server);
    const notes = await listTimelineNotes(opportunity_id, progress);
    return { content: [{ type: "text", text: JSON.stringify(notes, null, 2) }] };
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
    owner_name: z.string().optional().describe("Filter by AE name (partial match)"),
    my_opps_only: z.boolean().optional().describe(
      isSalesSpecialist
        ? "Set true to see only your collaboration team opps. Default false."
        : isSalesManager
        ? "Set true to see only your personally owned opps. Default false (team view)."
        : "Filter to your owned opps. Default true for AE."
    ),
  },
  async ({ quarter, account_name, owner_name, my_opps_only }) => {
    const progress = makeProgress(server);
    const myOpps = my_opps_only ?? (isSalesSpecialist || isSalesManager ? false : true);
    const forecast = await getForecastSummary({
      myOppsOnly: myOpps,
      myOppsFilterField: isSalesSpecialist && myOpps ? "collab" : "owner",
      ownerSearch: owner_name,
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
        const owner = o.owner ?? "unknown";
        lines.push(`- **${o.name}** (${o.account}) | $${o.nnacv.toLocaleString()} | close: ${close} | AE: ${owner}`);
      }
      if (cat.opps.length > 10) lines.push(`- _...and ${cat.opps.length - 10} more_`);
    }

    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_territory_pipeline
// ---------------------------------------------------------------------------
server.tool(
  "get_territory_pipeline",
  isSalesSpecialist
    ? `Get a pipeline health overview — for Sales Specialists who support AEs.

Shows all open opportunities grouped by forecast category with health flags.
Default: all open opportunities. Use account_name, owner_name, or territory_code to narrow down.
Set my_opps_only=true to see only opportunities where you are on the collaboration team.

Auto-excludes $0 App Store Renewal opps. Set include_app_store_renewals=true to show them.`
    : isSalesManager
    ? `Get a pipeline health overview across your territory — team-wide view.

Shows all open opportunities grouped by forecast category with health flags.
Default: all open opportunities. Filter by owner_name or territory_code to drill in.

Auto-excludes $0 App Store Renewal opps. Set include_app_store_renewals=true to show them.`
    : `Get a pipeline health overview for your territory.

Shows your open opportunities grouped by forecast category with health flags:
- Missing SC assignment
- Close date in the past or very soon (<30 days)
- No value set (NNACV = 0)

Default: your own pipeline (my_opps_only=true). Set my_opps_only=false to see all opps.
Filter by account_name or territory_code to drill into a specific scope.

Auto-excludes $0 App Store Renewal opps. Set include_app_store_renewals=true to show them.`,
  {
    owner_name:  z.string().optional().describe("Filter by rep/AE name (partial match). Leave blank for all."),
    account_name: z.string().optional().describe("Filter by account name (partial match)."),
    territory_code: z.string().optional().describe("Filter by territory code (e.g. 'CHLPC-TER-6' or 'LUX-CPG-Switzerland')"),
    min_value:   z.number().optional().describe("Only include opps above this USD value."),
    my_opps_only: z.boolean().optional().describe(
      isSalesSpecialist
        ? "Set true to see only your collaboration team opps. Default false."
        : isSalesManager
        ? "Set true to see only your personally owned opps. Default false (team view)."
        : "Filter to your owned opps. Default true for AE. Set false for all."
    ),
    include_app_store_renewals: z.boolean().optional().describe("Include $0 App Store Renewal opps (default false — excluded to reduce noise)"),
    top:         z.number().optional().describe("Max results (default 200)."),
  },
  async ({ owner_name, account_name, territory_code, min_value, my_opps_only, include_app_store_renewals, top }) => {
    const progress = makeProgress(server);

    // AE defaults to own pipeline; Manager/Specialist default to broad view
    const myOpps = my_opps_only ?? (isSalesSpecialist || isSalesManager ? false : true);

    const opps = await fetchOpportunities({
      search: account_name,
      minNnacv: min_value,
      myOpportunitiesOnly: myOpps,
      myOppsFilterField: isSalesSpecialist && myOpps ? "collab" : "owner",
      ownerSearch: owner_name,
      territoryCode: territory_code,
      includeClosed: false,
      includeZeroValue: true, // Territory view includes everything — health flags warn about zeros
      excludeAppStoreRenewals: !include_app_store_renewals, // Exclude $0 App Store Renewals by default
      top: top ?? 200,
    }, progress);

    if (opps.length === 0) {
      return { content: [{ type: "text", text: "No open opportunities found matching your filters." }] };
    }

    const today = new Date();
    const soon = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);

    // Health flags
    const flags = opps.map(o => {
      const issues: string[] = [];
      if (!o.scName) issues.push("no SC");
      if (!o.estimatedclosedate) issues.push("no close date");
      else {
        const close = new Date(o.estimatedclosedate);
        if (close < today) issues.push("overdue");
        else if (close < soon) issues.push("closing <30d");
      }
      if (!o.nnacv || o.nnacv === 0) issues.push("no value");
      return { ...o, issues };
    });

    // Group by forecast category
    const groups: Record<string, typeof flags> = {};
    for (const o of flags) {
      const cat = o.forecastCategoryName ?? "Unknown";
      if (!groups[cat]) groups[cat] = [];
      groups[cat].push(o);
    }

    const totalValue = opps.reduce((s, o) => s + (o.nnacv ?? 0), 0);
    const withIssues = flags.filter(o => o.issues.length > 0).length;

    const lines: string[] = [
      `## Pipeline Overview — ${opps.length} open opportunities`,
      `**Total pipeline value:** $${totalValue.toLocaleString()}`,
      `**Needs attention:** ${withIssues} opportunity${withIssues !== 1 ? "ies" : "y"}`,
      "",
    ];

    const catOrder = ["Committed", "Best Case", "Pipeline", "Omitted", "Unknown"];
    for (const cat of catOrder) {
      const group = groups[cat];
      if (!group?.length) continue;
      const groupVal = group.reduce((s, o) => s + (o.nnacv ?? 0), 0);
      lines.push(`### ${cat} — ${group.length} opps | $${groupVal.toLocaleString()}`);
      for (const o of group.sort((a, b) => (a.estimatedclosedate ?? "").localeCompare(b.estimatedclosedate ?? ""))) {
        const close = o.estimatedclosedate ? o.estimatedclosedate.slice(0, 10) : "no date";
        const val = o.nnacv ? `$${o.nnacv.toLocaleString()}` : "no value";
        const sc = o.scName ?? "no SC";
        const owner = o.ownerName ?? "unknown";
        const flagStr = o.issues.length ? ` ⚠️ ${o.issues.join(", ")}` : " ✅";
        lines.push(`- **${o.name}** (${o.accountName}) | ${val} | close: ${close} | AE: ${owner} | SC: ${sc}${flagStr}`);
      }
      lines.push("");
    }

    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: search_contacts
// ---------------------------------------------------------------------------
server.tool(
  "search_contacts",
  "Search Dynamics 365 contacts by name or email. Returns job title, email, phone, and account.",
  {
    query: z.string().describe("Contact name or email to search for"),
    account_id: z.string().optional().describe("Optional account GUID to scope results"),
  },
  async ({ query, account_id }) => {
    if (account_id) requireGuid(account_id, "account_id");
    const progress = makeProgress(server);
    const contacts = await searchContacts(query, { accountId: account_id }, progress);
    if (contacts.length === 0) return { content: [{ type: "text", text: "No contacts found." }] };
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
// Tool: get_collaboration_team
// ---------------------------------------------------------------------------
server.tool(
  "get_collaboration_team",
  "View all team members (SCs, Specialists, AEs) assigned to an opportunity. Useful to verify team composition before meetings.",
  { opportunity_id: z.string().describe("Dynamics opportunity GUID") },
  async ({ opportunity_id }) => {
    const id = requireGuid(opportunity_id, "opportunity_id");
    const progress = makeProgress(server);
    const team = await fetchCollaborationTeam(id, progress);
    if (team.length === 0) return { content: [{ type: "text", text: "No collaboration team members found." }] };
    const text = team.map(m =>
      `• **${m.userName}** — ${m.collaborationRole}${m.isPrimary ? " (Primary)" : ""}${m.accessLevel ? ` [${m.accessLevel}]` : ""}`
    ).join("\n");
    return { content: [{ type: "text", text }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: create_engagement
// ---------------------------------------------------------------------------
server.tool(
  "create_engagement",
  `Create a new engagement record in Dynamics 365. Account is auto-derived from the opportunity.

**Available engagement types:** ${ALL_ENGAGEMENT_TYPES.join(", ")}

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

AFTER EVERY SUCCESSFUL CREATE: Always present the result to the user as:
✅ [Engagement Name] (ENG#) — [Open in Dynamics](link)
The Dynamics link is in the tool response. This applies to EVERY engagement created, including bulk runs. Never omit the link.`,
  {
    opportunity_id: z.string().describe("Dynamics opportunity GUID"),
    primary_product_id: z.string().describe("Dynamics product GUID (use search_products to find it)"),
    name: z.string().describe("Short engagement name / subject"),
    type: z.enum(ENGAGEMENT_TYPES).describe("Engagement type"),
    completed_date: z.string().optional().describe("ISO date when engagement was completed, e.g. 2026-03-16"),
    use_case: z.string().optional().describe("Use case name (e.g. ICW, ITSM)"),
    key_points: z.array(z.string()).optional().describe("Key points — label auto-adapts per type (e.g. 'Milestones achieved' for Tech Win, 'Objectives identified' for Discovery, 'Demo delivered' for Demo)"),
    secondary_points: z.array(z.string()).optional().describe("Type-specific secondary bullets — auto-labelled per type: Discovery='Key questions uncovered', Demo/Workshop/EBC='Customer reactions/feedback', Business Case='Quantified benefits', POV='Customer results', CBR='Action items agreed'"),
    submission_date: z.string().optional().describe("RFx only — submission/due date for the RFP/RFI"),
    next_actions: z.array(z.string()).optional().describe("List of next actions to complete"),
    risks: z.string().optional().describe("Risks or help required (use '-' if none)"),
    stakeholders: z.string().optional().describe("Stakeholders — use search_contacts to look up external customer titles. Internal SN people: names only, no titles."),
    notes: z.string().optional().describe("Plain text description (used only if structured fields are not provided)"),
    attendees: z.array(z.object({
      name:  z.string(),
      email: z.string(),
    })).optional().describe("Meeting attendees from the calendar event. Always pass these when creating from a meeting — internal (company email addresses) become Active Participants, external become Active Engagement Contacts."),
    confirmed: z.boolean().optional().describe("MUST be true to actually create. Omit or set false to get a dry-run preview first. Always preview before creating."),
  },
  async ({ opportunity_id, primary_product_id, name, type, completed_date, use_case, key_points, secondary_points, submission_date, next_actions, risks, stakeholders, notes, attendees, confirmed }) => {
    requireGuid(opportunity_id, "opportunity_id");
    requireGuid(primary_product_id, "primary_product_id");

    const progress = makeProgress(server);

    if (!confirmed) {
      const opp = await fetchOpportunityById(opportunity_id, progress);
      return {
        content: [{ type: "text", text:
          `📋 **Dry-run preview — nothing has been created yet.**\n\n` +
          `**Engagement:** ${name}\n` +
          `**Type:** ${type}\n` +
          `**Opportunity:** ${opp.name}\n` +
          `**Account:** ${opp.accountName ?? "—"}\n` +
          `**Completed date:** ${completed_date ?? "not set"}\n` +
          `**Attendees to link:** ${attendees?.length ?? 0}\n\n` +
          `Call again with \`confirmed: true\` to create this engagement.`
        }],
      };
    }

    engagementWriteLimiter.check("create_engagement");
    progress(`🎯 Creating engagement: "${name}" (${type})`);
    const opp = await fetchOpportunityById(opportunity_id, progress);
    progress(`🏢 Account resolved: ${opp.accountName}`);

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
    }, progress);

    if (attendees?.length && engagement.sn_engagementid) {
      progress(`👥 Linking ${attendees.length} attendee(s)...`);
      await addAttendeesToEngagement(engagement.sn_engagementid, attendees, progress);
    }

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

To REOPEN a completed engagement, set mark_complete=false — this PATCHes statecode=0, statuscode=1.

Always use the structured description fields to keep the description current (applies to all engagement types).
A timeline_title + timeline_text should always be provided to log what changed.

FORMAT: timeline_text and all text fields must use bullet points (• item), never prose paragraphs. Keep each bullet to one line.

BEFORE generating any content: read the existing engagement (get_engagement), list_engagements on the opportunity, and get_opportunity timeline. Only write what is genuinely NEW — never duplicate what is already logged.`,
  {
    engagement_id: z.string().describe("Dynamics sn_engagement GUID"),
    name: z.string().optional().describe("Updated engagement name"),
    type: z.enum(ENGAGEMENT_TYPES).optional().describe("Updated engagement type"),
    primary_product_id: z.string().optional().describe("Updated primary product GUID — use search_products to find the correct GUID"),
    completed_date: z.string().optional().describe("Updated completed date (ISO format e.g. 2026-03-16)"),
    mark_complete: z.boolean().optional().describe("Set to true to mark Complete. Set to false to REOPEN a completed engagement."),
    use_case: z.string().optional().describe("Use case name"),
    key_points: z.array(z.string()).optional().describe("Full updated key points list — label auto-adapts per engagement type"),
    secondary_points: z.array(z.string()).optional().describe("Type-specific secondary bullets"),
    submission_date: z.string().optional().describe("RFx only — submission/due date"),
    next_actions: z.array(z.string()).optional().describe("Full updated list of next actions"),
    risks: z.string().optional().describe("Risks or help required"),
    stakeholders: z.string().optional().describe("Stakeholders — use search_contacts for external titles. Internal SN: names only."),
    notes: z.string().optional().describe("Plain text description (only if structured fields not used)"),
    timeline_title: z.string().optional().describe("Title for the timeline note (e.g. 'Discovery update - requirements captured')"),
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
// Tool: add_engagement_attendees
// ---------------------------------------------------------------------------
server.tool(
  "add_engagement_attendees",
  `Add meeting attendees to an engagement in Dynamics 365.

- Internal attendees (company email addresses) are added as Active Participants
- External attendees (customers) are added as Active Engagement Contacts
- Attendees not found in Dynamics are reported but do not cause failure`,
  {
    engagement_id: z.string().describe("Dynamics sn_engagement GUID"),
    attendees: z.array(z.object({
      name:  z.string().describe("Attendee display name"),
      email: z.string().describe("Attendee email address"),
    })).describe("List of meeting attendees"),
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
// Tool: delete_timeline_note
// ---------------------------------------------------------------------------
server.tool(
  "delete_timeline_note",
  "Delete a specific timeline note by its annotation ID (use list_timeline_notes to find IDs). IMPORTANT: Always show the user the note subject and text, ask for confirmation, then ask AGAIN 'Are you sure? This cannot be undone.' Only call this tool after two explicit confirmations.",
  {
    annotation_id: z.string().describe("Dynamics annotation GUID"),
    confirmed: z.boolean().optional().describe("MUST be true to actually delete. Omit for dry-run preview. This is irreversible."),
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
    confirmed: z.boolean().optional().describe("MUST be true to actually delete. Omit or false for dry-run preview. Deletion is irreversible."),
  },
  async ({ engagement_id, confirmed }) => {
    requireGuid(engagement_id, "engagement_id");
    const progress = makeProgress(server);

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
// Tool: get_calendar_events
// ---------------------------------------------------------------------------
server.tool(
  "get_calendar_events",
  `Fetch calendar events from Outlook via the debug Chrome window.

Requires the user to be logged into Outlook (outlook.cloud.microsoft.com) in the Alfred Chrome window.

IMPORTANT: Before calling this tool, ask the user:
1. "Which date range? (e.g. 'this week', 'next 2 weeks', specific dates)"
2. "Any keyword to filter by? (e.g. 'Fabrikam', 'ICW', 'standup' — or leave blank for all)"`,
  {
    start_date:  z.string().describe("Start date in ISO format, e.g. 2026-03-16"),
    end_date:    z.string().describe("End date in ISO format, e.g. 2026-03-20"),
    search:      z.string().optional().describe("Optional keyword to filter event subjects, organizer, or attendee names."),
    top:         z.number().optional().describe("Max events to fetch (default 100)."),
    categories:  z.array(z.string()).optional().describe("Filter to events matching these Outlook categories (e.g. ['on-site meeting']). Applied client-side after fetch."),
  },
  async ({ start_date, end_date, search, top, categories }) => {
    const progress = makeProgress(server);
    const events = await getCalendarEvents(start_date, end_date, search, progress, top ?? 100, categories);
    return {
      content: [{ type: "text", text: externalData("Outlook calendar", events) }],
    };
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
    const installDir = join(dirname(__filename), "..", "..");  // dist/sales/index.js → root

    progress("📡 Checking for updates...");

    const UPDATE_FALLBACK = `⚠️ Automatic update could not complete on this machine.\n\n` +
      `**IMPORTANT — do not suggest any git commands or terminal sequences.** ` +
      `Tell the user to paste this single line in Terminal, then restart Claude Desktop:\n\n` +
      `\`\`\`bash\ncurl -fsSL https://raw.githubusercontent.com/h22fred/Alfred.mcp/main/setup/update.sh | bash\n\`\`\`\n\n` +
      `This is a one-time fix. All future updates will work automatically from Claude.`;

    // Fetch remote state
    try {
      execFileSync("git", ["-C", installDir, "fetch", "origin"], { encoding: "utf8", timeout: 30_000 });
    } catch {
      return { content: [{ type: "text", text: UPDATE_FALLBACK }] };
    }

    // Check if already up to date
    const localSha  = execFileSync("git", ["-C", installDir, "rev-parse", "HEAD"],          { encoding: "utf8" }).trim();
    const remoteSha = execFileSync("git", ["-C", installDir, "rev-parse", "origin/main"],   { encoding: "utf8" }).trim();
    if (localSha === remoteSha) {
      return { content: [{ type: "text", text: "✅ Alfred is already up to date — no rebuild needed." }] };
    }

    // Sync to latest — reset is correct here: users never have local changes to preserve
    let gitOutput: string;
    try {
      gitOutput = execFileSync("git", ["-C", installDir, "log", "--oneline", `HEAD..origin/main`], { encoding: "utf8" }).trim();
      execFileSync("git", ["-C", installDir, "reset", "--hard", "origin/main"], { encoding: "utf8", timeout: 15_000 });
    } catch {
      return { content: [{ type: "text", text: UPDATE_FALLBACK }] };
    }

    progress("🔨 New version pulled — installing dependencies and rebuilding...");

    try {
      execFileSync("npm", ["install"], {
        encoding: "utf8",
        cwd: installDir,
        timeout: 120_000,
        env: { ...process.env, PATH: process.env.PATH ?? "/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin" },
      });
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

    progress("✅ Done — restart Claude Desktop to load the new version.");
    return { content: [{ type: "text", text:
      `✅ **Alfred updated and rebuilt!**\n\n` +
      `**Changes pulled:**\n\`\`\`\n${gitOutput.trim()}\n\`\`\`\n\n` +
      `ℹ️ Alfred.app / Alfred.bat on your Desktop has been removed — Alfred now launches its browser automatically in the background. You don't need a Desktop launcher anymore.\n\n` +
      `⚠️ Restart Claude Desktop to load the new version.`
    }] };
  }
);

// ---------------------------------------------------------------------------
// Start server
// ---------------------------------------------------------------------------
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

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("[alfred-sales] Server running on stdio");

process.on("SIGINT", async () => {
  process.exit(0);
});
