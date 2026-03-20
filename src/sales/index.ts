import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { DYNAMICS_HOST, alfredConfig as _baseConfig } from "../config.js";
import {
  fetchOpportunities,
  fetchOpportunityById,
  searchAccounts,
  fetchAccountById,
  createOpportunity,
  updateOpportunity,
  searchSystemUsers,
  fetchCurrentUserId,
  createTimelineNote,
  listTimelineNotes,
} from "../tools/dynamicsClient.js";
import { closeBrowser, ensureAlfred } from "../auth/tokenExtractor.js";

const DYNAMICS_BASE_URL = DYNAMICS_HOST;

// ---------------------------------------------------------------------------
// Security helpers
// ---------------------------------------------------------------------------

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function requireGuid(value: string, label: string): string {
  if (!GUID_RE.test(value)) throw new Error(`Invalid ${label}: expected a Dynamics GUID.`);
  return value;
}

function makeProgress(srv: McpServer) {
  return (msg: string) => {
    console.error(`[progress] ${msg}`);
    srv.server.sendLoggingMessage({ level: "info", data: msg });
  };
}

const FORECAST_NAMES: Record<number, string> = {
  100000001: "Pipeline",
  100000002: "Best Case",
  100000003: "Committed",
  100000004: "Omitted",
};

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
  "List your open opportunities in Dynamics 365, optionally filtered by account name or minimum value.",
  {
    search:   z.string().optional().describe("Filter by account or opportunity name"),
    min_value: z.number().optional().describe("Minimum total value in USD"),
    include_closed: z.boolean().optional().describe("Include won/lost opportunities (default false)"),
  },
  async ({ search, min_value, include_closed }) => {
    const progress = makeProgress(server);
    const opps = await fetchOpportunities({
      search,
      minNnacv: min_value,
      myOpportunitiesOnly: true,
      includeClosed: include_closed ?? false,
      top: 50,
    }, progress);
    return { content: [{ type: "text", text: JSON.stringify(opps, null, 2) }] };
  }
);

// ---------------------------------------------------------------------------
// Tool: get_opportunity
// ---------------------------------------------------------------------------
server.tool(
  "get_opportunity",
  "Get a single opportunity by its Dynamics ID.",
  { opportunity_id: z.string().describe("Dynamics opportunity GUID") },
  async ({ opportunity_id }) => {
    const progress = makeProgress(server);
    requireGuid(opportunity_id, "opportunity_id");
    const opp = await fetchOpportunityById(opportunity_id, progress);
    const link = `${DYNAMICS_BASE_URL}/main.aspx?etn=opportunity&pagetype=entityrecord&id=${opp.opportunityid}`;
    return { content: [{ type: "text", text: JSON.stringify({ ...opp, dynamicsLink: link }, null, 2) }] };
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
    name:              z.string().describe("Opportunity name, e.g. 'Givaudan — New ITSM 2026'"),
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
        `**Notes:** ${notes ?? "—"}\n\n` +
        `Call again with \`confirmed: true\` to create this opportunity.`
      }] };
    }

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
  `Update an existing opportunity in Dynamics 365 — close date, forecast category, owner, SC, or name.

Always show the current values and the proposed changes, then confirm before calling with confirmed=true.`,
  {
    opportunity_id:    z.string().describe("Dynamics opportunity GUID"),
    name:              z.string().optional().describe("New opportunity name"),
    close_date:        z.string().optional().describe("New close date, ISO format"),
    forecast_category: z.string().optional().describe("'Pipeline', 'Best Case', or 'Committed'"),
    owner_id:          z.string().optional().describe("New Sales Rep systemuser GUID"),
    sc_id:             z.string().optional().describe("New SC systemuser GUID"),
    notes:             z.string().optional().describe("Updated description/notes"),
    confirmed:         z.boolean().optional().describe("MUST be true to actually update. Omit for dry-run."),
  },
  async ({ opportunity_id, name, close_date, forecast_category, owner_id, sc_id, notes, confirmed }) => {
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

      return { content: [{ type: "text", text:
        `📋 **Dry-run — nothing updated yet.**\n\n` +
        `**Opportunity:** ${current.name} (${current.sn_number ?? opportunity_id})\n\n` +
        `**Proposed changes:**\n${changes.length ? changes.map(c => `• ${c}`).join("\n") : "No changes specified"}\n\n` +
        `Call again with \`confirmed: true\` to apply.`
      }] };
    }

    const updated = await updateOpportunity({
      opportunityId: opportunity_id,
      name, closeDate: close_date, forecastCategory: forecastCode,
      ownerId: owner_id, scId: sc_id, notes,
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
// Tool: ensure_alfred
// ---------------------------------------------------------------------------
server.tool(
  "ensure_alfred",
  "Launch Alfred (Chrome with Dynamics session) if it is not already running. Call this if you get auth errors.",
  {},
  async () => {
    const progress = makeProgress(server);
    await ensureAlfred(progress);
    return { content: [{ type: "text", text: "✅ Alfred is running — Dynamics session ready." }] };
  }
);

// ---------------------------------------------------------------------------
// Start server
// ---------------------------------------------------------------------------
const transport = new StdioServerTransport();
await server.connect(transport);
console.error("[alfred-sales] Server running on stdio");

process.on("SIGINT", async () => {
  await closeBrowser();
  process.exit(0);
});
