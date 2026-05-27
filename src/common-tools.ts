/**
 * Common tool registrations — shared by both sc/index.ts and sales/index.ts.
 *
 * Export a single function that registers all 37 shared tools on the given server.
 * The `role` parameter is available for future per-role branching; for now every
 * tool in this module is identical between SC and Sales.
 *
 * SOURCE OF TRUTH: SC version (src/sc/index.ts) is the canonical copy for all
 * tools in this file, including the 9 cosmetically-different ones.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { readFileSync, existsSync, writeFileSync } from "fs";
import { homedir } from "os";
import { join } from "path";
import { execFileSync } from "child_process";

import { DYNAMICS_HOST } from "./config.js";
import { requireGuid, makeProgress, WriteRateLimiter, externalData, getMs365MigrationNotice } from "./shared.js";
import {
  fetchEngagementsByOpportunity,
  fetchEngagementsByAccount,
  fetchEngagementsGlobal,
  fetchEngagementById,
  fetchOpportunityById,
  fetchCollaborationTeam,
  createAccountEngagement,
  addAttendeesToEngagement,
  searchSystemUsers,
  searchProducts,
  getProductById,
  buildDescription,
  listTimelineNotes,
  fetchMyCollaborationOpportunities,
  fetchEngagementParticipants,
  fetchMyEngagementAssignments,
  resolveOpportunityId,
  listCollaborationNotes,
  createCollaborationNote,
  listActivities,
  createAppointment,
  createPhoneCall,
  createTask,
  completeActivity,
  createContact,
  listOpportunityContacts,
  addContactToOpportunity,
  listClosingPlan,
  createClosingPlanMilestone,
  updateClosingPlanMilestone,
  getOpportunitySummary,
  updateOpportunitySummary,
  listQuotes,
  type EngagementType,
  type EngagementDescription,
} from "./tools/dynamicsClient.js";
import { ensureAlfred, exitAlfred, restartAlfred, clearAuthCache } from "./auth/tokenExtractor.js";
import { clearGraphTokenCache } from "./tools/outlookClient.js";
import { postTeamsNotification, getTeamsTranscript, setTeamsWebhook } from "./tools/teamsClient.js";
import { runHygieneSweep, formatHygieneReport } from "./tools/hygieneClient.js";
import { detectPostMeetingEngagements, notifyPostMeetingCandidates } from "./tools/postMeetingClient.js";

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

const DYNAMICS_BASE_URL = DYNAMICS_HOST;

type Engagement = import("./tools/dynamicsClient.js").Engagement;

function engagementLink(e: Engagement): string | null {
  const id = e.sn_engagementid ?? "";
  return id ? `${DYNAMICS_BASE_URL}/main.aspx?etn=sn_engagement&id=${id}&pagetype=entityrecord` : null;
}

function engagementListItem(e: Engagement): string {
  const link = engagementLink(e);
  const status = e.statuscode === 876130000 ? "Cancelled" : e.statecode === 0 ? "Open" : "Complete";
  const completed = e.sn_completeddate ? ` · ${e.sn_completeddate.slice(0, 10)}` : "";
  const product = e.primaryProductName ? ` · 📦 ${e.primaryProductName}` : "";
  const lines = [
    `**${e.sn_name}** (${e.sn_engagementnumber ?? "—"}) · ${e.engagementTypeName ?? "—"} · ${status}${completed}${product}`,
    ...(link ? [`🔗 Open in Dynamics: ${link}`] : []),
    ...(e.sn_description ? [e.sn_description] : []),
  ];
  return lines.join("\n");
}

// ---------------------------------------------------------------------------
// Exported registration function
// ---------------------------------------------------------------------------

export function registerCommonTools(
  server: McpServer,
  _role: "sc" | "sales",
  writeLimiter?: WriteRateLimiter
): void {
  // Reuse the caller's limiter if provided — ensures common-tools operations count
  // against the same ceiling as sc/sales-specific write operations (M1 fix).
  const engagementWriteLimiter = writeLimiter ?? new WriteRateLimiter(10, 10 * 60 * 1000);

  // Emit migration notice to stderr on startup; injected into first tool response below.
  const migrationNotice = getMs365MigrationNotice();
  if (migrationNotice) process.stderr.write(`[alfred:notice] ${migrationNotice.replace(/\n/g, " | ")}\n`);

  // ---------------------------------------------------------------------------
  // Tool: open_chrome_debug
  // ---------------------------------------------------------------------------
  server.tool(
    "open_chrome_debug",
    `Launch Alfred (browser) if it's not already running. Opens Dynamics and Teams tabs automatically.

IMPORTANT: Call this tool AUTOMATICALLY — without asking the user — whenever any tool fails with an error mentioning:
- "Alfred not running"
- "Chrome debug port not available"
- "No page targets"
- "CDP" or "debug port"
- "stale" or "session"
- "Could not capture Graph token"
- "401" or "unauthorized"
- "not logged in" or "not logged into"

This tool also clears all cached auth tokens, so call it proactively if you suspect a stale session.

IMPORTANT AFTER CALLING THIS TOOL:
- Tell the user: "Alfred is open — please log into Dynamics and Teams, then let me know when you're ready."
- STOP and wait for the user to confirm they are logged in before retrying any other tool.
- Do NOT automatically retry the original tool — the user must log in first.`,
    {},
    async () => {
      const progress = makeProgress(server);
      // Clear all token caches — ensures fresh auth after any browser restart
      clearAuthCache();
      clearGraphTokenCache();
      await ensureAlfred(progress);
      const notice = getMs365MigrationNotice();
      const baseMsg = "✅ Alfred is open. Please log into Dynamics and Teams in the browser window, then tell me when you're ready and I'll continue.";
      return {
        content: [{ type: "text", text: notice ? `${notice}\n\n---\n\n${baseMsg}` : baseMsg }],
      };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: exit_alfred
  // ---------------------------------------------------------------------------
  server.tool(
    "exit_alfred",
    `Close the Alfred browser window. This only closes the Alfred Chrome instance (with the debug profile) — your regular Chrome is untouched.

Use this when:
- The user says "close Alfred", "exit Alfred", "stop Alfred", or "shut down Alfred"
- You need to free up resources
- The session is done

After exiting, all cached auth tokens are cleared. The user will need to relaunch Alfred from the Desktop to use CRM tools again.`,
    {},
    async () => {
      const progress = makeProgress(server);
      clearGraphTokenCache();
      const wasRunning = await exitAlfred(progress);
      return {
        content: [{ type: "text", text: wasRunning
          ? "✅ Alfred closed. To use CRM tools again, double-click Alfred on your Desktop."
          : "ℹ️ Alfred was not running." }],
      };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: restart_alfred
  // ---------------------------------------------------------------------------
  server.tool(
    "restart_alfred",
    `Restart the Alfred browser — closes and relaunches it with a fresh session.

Use this when:
- Auth tokens are stale or expired (401/403 errors)
- The user says "restart Alfred" or "refresh Alfred"
- Multiple tools are failing with connection errors
- After "exit Alfred" if the user wants to start fresh

This only restarts the Alfred Chrome instance — your regular Chrome is untouched.
The Alfred icon is preserved on relaunch.`,
    {},
    async () => {
      const progress = makeProgress(server);
      clearGraphTokenCache();
      await restartAlfred(progress);
      return {
        content: [{ type: "text", text: "✅ Alfred restarted. Please log into Dynamics, Outlook and Teams in the new window, then tell me when you're ready." }],
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
      const id = requireGuid(opportunity_id, "opportunity_id");
      const progress = makeProgress(server);
      const fetchedAt = new Date().toISOString();
      const engagements = await fetchEngagementsByOpportunity(id, progress);
      if (engagements.length === 0) {
        return { content: [{ type: "text", text: `No engagements found for this opportunity. (Retrieved fresh from Dynamics at ${fetchedAt})` }] };
      }
      const text = engagements.map(engagementListItem).join("\n\n---\n\n") + `\n\n_Retrieved from Dynamics at ${fetchedAt}_`;
      return { content: [{ type: "text", text }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: list_engagements_by_account
  // ---------------------------------------------------------------------------
  server.tool(
    "list_engagements_by_account",
    `List all engagements linked to a Dynamics account, regardless of whether they have an associated opportunity.

Use this when you need a complete picture of engagement activity on an account — some SCs log POVs and other milestones directly at the account level rather than on a specific opportunity, so those records won't appear in list_engagements.

Optionally filter by engagement type (e.g. type="POV") to keep the response lean.`,
    {
      account_id: z.string().describe("Dynamics account GUID"),
      type:       z.string().optional().describe("Optional engagement type filter, e.g. 'POV', 'Demo', 'Technical Win'"),
    },
    async ({ account_id, type }) => {
      const id = requireGuid(account_id, "account_id");
      const progress = makeProgress(server);
      const engagements = await fetchEngagementsByAccount(id, type, progress);
      if (engagements.length === 0) {
        const msg = type
          ? `No ${type} engagements found for this account.`
          : "No engagements found for this account.";
        return { content: [{ type: "text", text: msg }] };
      }
      const text = engagements.map(engagementListItem).join("\n\n---\n\n");
      return { content: [{ type: "text", text }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: create_account_engagement
  // ---------------------------------------------------------------------------
  server.tool(
    "create_account_engagement",
    `Create an engagement linked directly to an account — no opportunity required.

Use this for on-site visits, executive meetings, workshops, or QBRs that are not tied to a specific deal.
Best-fit types for account-level engagements:
- EBC — executive on-site visit or briefing
- Workshop — technical or working session on-site
- Customer Business Review — QBR or strategic review
- Discovery — on-site discovery not yet tied to a deal
- Post Sale Engagement — post-sale on-site activity

CONTENT GUIDELINES BY TYPE:
- Workshop / EBC / Customer Business Review: these are point-in-time events — keep description minimal. List topics covered and any key outcomes. 3–5 bullets max. Always auto-complete (will be marked Complete automatically).
- Discovery: capture the customer's current state, key pain points, and goals. 5–8 bullets.
- Post Sale Engagement: focus on what was delivered and any open follow-ups.

ATTENDEES: if attendees are provided (names or email addresses), add them as participants after creation.

⚠️ OWNER OVERRIDE: if owner_name is set, always resolve and show the full matched name to the user and get explicit approval BEFORE setting owner_confirmed=true. Creating under the wrong person's name cannot be undone from Alfred.

BEFORE creating: call list_engagements_by_account to check for existing engagements of the same type.

AFTER EVERY SUCCESSFUL CREATE, show the user a direct CRM link:
${DYNAMICS_HOST}/main.aspx?etn=sn_engagement&id=<sn_engagementid>&pagetype=entityrecord
Never omit the link.`,
    {
      account_id:         z.string().describe("Dynamics account GUID"),
      type:               z.enum(["EBC", "Workshop", "Customer Business Review", "Discovery", "Post Sale Engagement"] as const).describe("Engagement type"),
      name:               z.string().describe("Engagement name, e.g. 'On-Site EBC – Contoso – May 2026'"),
      primary_product_id: z.string().describe("Product family GUID — required by Dynamics. Use search_products to look up the GUID. Always resolve before creating."),
      key_points:         z.array(z.string()).optional().describe("Main topics covered / outcomes. Each item becomes a bullet point."),
      next_actions:       z.array(z.string()).optional().describe("Follow-up actions. Each item becomes a bullet point."),
      risks:              z.string().optional().describe("Risks or help required (free text)."),
      stakeholders:       z.string().optional().describe("Key stakeholders present (free text)."),
      notes:              z.string().optional().describe("Free-text notes (used when structured fields not provided)."),
      completed_date:     z.string().optional().describe("ISO date if already completed, e.g. '2026-05-08'"),
      attendees:          z.array(z.object({ name: z.string(), email: z.string() })).optional().describe("Attendees to add as participants after creation — provide name + email for each person."),
      owner_name:         z.string().optional().describe("Name of the SC who should own this engagement (partial match). Use ONLY when explicitly asked to create on behalf of a colleague."),
      owner_confirmed:    z.boolean().optional().describe("REQUIRED when owner_name is set. Must be explicitly true AFTER showing the user the resolved full name and getting their explicit approval."),
    },
    async ({ account_id, type, name, primary_product_id, key_points, next_actions, risks, stakeholders, notes, completed_date, attendees, owner_name, owner_confirmed }) => {
      const id = requireGuid(account_id, "account_id");
      const progress = makeProgress(server);
      engagementWriteLimiter.check("create_account_engagement");

      const ownerUsers = owner_name ? await searchSystemUsers(owner_name, progress) : [];
      const resolvedOwner = ownerUsers[0];
      const ownerId = resolvedOwner?.systemuserid;

      // Safety gate: block creation under another user's name until owner_confirmed is explicitly true
      if (owner_name && resolvedOwner && !owner_confirmed) {
        return { content: [{ type: "text", text:
          `🛑 **Owner confirmation required.**\n\n` +
          `This engagement will be created under **${resolvedOwner.fullname}** (${resolvedOwner.internalemailaddress ?? resolvedOwner.systemuserid})'s name.\n\n` +
          `⚠️ This cannot be undone from Alfred — ${resolvedOwner.fullname} would need to manually reassign it in Dynamics if incorrect.\n\n` +
          `Please confirm with the user: "Create this engagement under ${resolvedOwner.fullname}'s name?" — then retry with \`owner_confirmed: true\`.`
        }] };
      }

      if (owner_name && !ownerId) progress(`⚠️ No user found for "${owner_name}" — creating under your own name`);
      else if (ownerId) progress(`👤 Owner: ${resolvedOwner!.fullname}`);

      const hasStructured = key_points?.length || next_actions?.length || risks || stakeholders;
      const description = hasStructured
        ? buildDescription({ engagementType: type as EngagementType, keyPoints: key_points, nextActions: next_actions, risks, stakeholders })
        : undefined;

      const engagement = await createAccountEngagement({
        accountId: id,
        type,
        name,
        primaryProductId: primary_product_id,
        description,
        notes: description ? undefined : notes,
        completedDate: completed_date,
        ownerId,
      }, progress);

      const engId = engagement.sn_engagementid;
      if (engId && attendees?.length) {
        await addAttendeesToEngagement(engId, attendees, progress).catch(e => {
          process.stderr.write(`[alfred:warn] addAttendeesToEngagement failed: ${e instanceof Error ? e.message : String(e)}\n`);
        });
      }

      const link = `${DYNAMICS_HOST}/main.aspx?etn=sn_engagement&id=${engId}&pagetype=entityrecord`;
      return { content: [{ type: "text", text: `✅ Account engagement created: ${engagement.sn_name}\nID: ${engagement.sn_engagementnumber ?? engId}\nLink: ${link}` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: search_engagements
  // ---------------------------------------------------------------------------
  server.tool(
    "search_engagements",
    `Search all engagements across all accounts and opportunities — no ownership filter.

Use this when you need a region-wide or team-wide view. For example:
- "Show me all open POVs across EMEA"
- "List every Demo engagement in Germany"
- "Find all active POVs in the DACH region"
- Managers checking pipeline health across their whole territory

Filters:
- type: engagement type name, e.g. "POV", "Demo", "Discovery"
- status: "open" | "complete" | "all" (default: all)
- search: partial match on engagement name
- country: filter by account country, e.g. "Germany", "France", "United Kingdom"
- region: geographic rollup — supported values: EMEA, DACH, Nordics, Benelux, "Central Europe",
  "Southern Europe", "Eastern Europe", "Middle East", Africa, AMER, "North America", LATAM,
  APAC, ANZ, Japan, "Greater China", India
- top: max results (default 50)

country and region are mutually exclusive — if both supplied, country takes precedence.
Filtering by country/region only matches engagements with a linked account that has a country set.`,
    {
      type:    z.string().optional().describe("Engagement type, e.g. 'POV', 'Demo', 'Discovery'"),
      status:  z.enum(["open", "complete", "all"]).optional().describe("Status filter (default: all)"),
      search:  z.string().optional().describe("Partial match on engagement name"),
      country: z.string().optional().describe("Account country, e.g. 'Germany', 'France', 'United Kingdom'"),
      region:  z.string().optional().describe("Geographic region, e.g. 'EMEA', 'DACH', 'Central Europe', 'APAC'"),
      top:     z.number().optional().describe("Max results (default 50)"),
    },
    async ({ type, status, search, country, region, top }) => {
      const progress = makeProgress(server);
      const engagements = await fetchEngagementsGlobal({ type, status, search, country, region, top }, progress);
      if (engagements.length === 0) {
        return { content: [{ type: "text", text: "No engagements found matching your criteria." }] };
      }
      const lines = engagements.map(e => engagementListItem(e));
      return { content: [{ type: "text", text: `Found ${engagements.length} engagement(s):\n\n${lines.join("\n\n---\n\n")}` }] };
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
      const id = requireGuid(product_id, "product_id");
      const progress = makeProgress(server);
      const product = await getProductById(id, progress);
      return {
        content: [{ type: "text", text: externalData("Dynamics product", product) }],
      };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: search_products
  // ---------------------------------------------------------------------------
  server.tool(
    "search_products",
    "Search Dynamics 365 product families by name — useful for identifying products on opportunities.",
    { name: z.string().describe("Product name or partial name to search for") },
    async ({ name }) => {
      const progress = makeProgress(server);
      const products = await searchProducts(name, progress);
      if (products.length === 0) return { content: [{ type: "text", text: "No products found." }] };
      const text = products.map(p => `**${p.name}** (${p.productid})`).join("\n");
      return { content: [{ type: "text", text }] };
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
      const id = requireGuid(engagement_id, "engagement_id");
      const progress = makeProgress(server);
      const notes = await listTimelineNotes(id, progress);
      return { content: [{ type: "text", text: externalData("Dynamics timeline notes", notes) }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: list_collaboration_notes
  // ---------------------------------------------------------------------------
  server.tool(
    "list_collaboration_notes",
    `List collaboration notes on an opportunity. These are the notes from the "Collaboration Notes" activity type in Dynamics (General Notes, Next Steps, Sales Ops Update, Renewal Update, Prime Notes) — NOT the same as timeline annotations.

Always read collaboration notes BEFORE creating/updating them to avoid duplicates.`,
    { opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number (e.g. OPTY5328326)") },
    async ({ opportunity_id }) => {
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const notes = await listCollaborationNotes(id, progress);
      return { content: [{ type: "text", text: externalData("Dynamics collaboration notes", notes) }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: create_collaboration_note
  // ---------------------------------------------------------------------------
  server.tool(
    "create_collaboration_note",
    `Create a Collaboration Note on an opportunity in Dynamics 365.

Note types: General Notes, Sales Ops Update, Renewal Update, Next Steps, Prime Notes.

IMPORTANT:
- Always call list_collaboration_notes first to check what already exists
- Use bullet-point format for the note content, not prose paragraphs
- Keep notes concise and actionable`,
    {
      opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number"),
      note_type: z.enum(["General Notes", "Sales Ops Update", "Renewal Update", "Next Steps", "Prime Notes"]).describe("Type of collaboration note"),
      notes: z.string().describe("The note content — use bullet points, keep concise"),
    },
    async ({ opportunity_id, note_type, notes }) => {
      engagementWriteLimiter.check("create_collaboration_note");
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const note = await createCollaborationNote({ opportunityId: id, noteType: note_type, notes }, progress);
      return { content: [{ type: "text", text: `✅ Created **${note.noteType}** collaboration note on opportunity.\n\n${note.notes}` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: list_activities
  // ---------------------------------------------------------------------------
  server.tool(
    "list_activities",
    `List all activities (appointments, phone calls, tasks, etc.) on an opportunity.

Shows open activities by default. Set include_completed=true to see all. Filter by type with activity_type.`,
    {
      opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number"),
      include_completed: z.boolean().optional().describe("Include completed/canceled activities (default: open only)"),
      activity_type: z.string().optional().describe("Filter by type: 'appointment', 'phonecall', 'task'"),
      top: z.number().optional().describe("Max results (default 50)"),
    },
    async ({ opportunity_id, include_completed, activity_type, top }) => {
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const activities = await listActivities(id, progress, { includeCompleted: include_completed, activityType: activity_type, top });
      return { content: [{ type: "text", text: externalData("Dynamics activities", activities) }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: create_appointment
  // ---------------------------------------------------------------------------
  server.tool(
    "create_appointment",
    `Create an appointment linked to an opportunity in Dynamics 365.

For #NBM (Next Best Meeting) appointments, prefix the subject with "#NBM" (e.g. "#NBM Discovery with Fabrikam").

IMPORTANT: Always confirm the date, time, subject, and attendees with the user before creating.`,
    {
      opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number"),
      subject: z.string().describe("Appointment subject (prefix with #NBM for Next Best Meeting)"),
      start_time: z.string().describe("Start time in ISO format (e.g. 2026-05-15T10:00:00Z)"),
      end_time: z.string().optional().describe("End time in ISO format (default: 1 hour after start)"),
      description: z.string().optional().describe("Meeting description / agenda"),
      location: z.string().optional().describe("Meeting location"),
      required_attendees: z.array(z.string()).optional().describe("Required attendee email addresses"),
      optional_attendees: z.array(z.string()).optional().describe("Optional attendee email addresses"),
    },
    async ({ opportunity_id, subject, start_time, end_time, description, location, required_attendees, optional_attendees }) => {
      engagementWriteLimiter.check("create_appointment");
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const appt = await createAppointment({
        opportunityId: id, subject, startTime: start_time, endTime: end_time,
        description, location, requiredAttendees: required_attendees, optionalAttendees: optional_attendees,
      }, progress);
      return { content: [{ type: "text", text: `✅ Appointment created: **${appt.subject}**\nStart: ${appt.scheduledstart}\nEnd: ${appt.scheduledend}` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: log_phone_call
  // ---------------------------------------------------------------------------
  server.tool(
    "log_phone_call",
    `Log a phone call activity linked to an opportunity in Dynamics 365.

Use this when the user says they had a call, need to log a call, or want to record a phone conversation.`,
    {
      opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number"),
      subject: z.string().describe("Call subject (e.g. 'Follow-up call with CFO')"),
      description: z.string().optional().describe("Call notes / summary"),
      phone_number: z.string().optional().describe("Phone number called"),
      direction: z.enum(["outgoing", "incoming"]).optional().describe("Call direction (default: outgoing)"),
    },
    async ({ opportunity_id, subject, description, phone_number, direction }) => {
      engagementWriteLimiter.check("log_phone_call");
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const call = await createPhoneCall({
        opportunityId: id, subject, description, phoneNumber: phone_number,
        directionCode: direction !== "incoming",
      }, progress);
      return { content: [{ type: "text", text: `✅ Phone call logged: **${call.subject}**` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: create_follow_up_task
  // ---------------------------------------------------------------------------
  server.tool(
    "create_follow_up_task",
    `Create a follow-up task linked to an opportunity in Dynamics 365.

Use this when the user says they need to do something later, has a to-do, or needs to set a reminder for an opportunity.`,
    {
      opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number"),
      subject: z.string().describe("Task subject (e.g. 'Send proposal to legal')"),
      description: z.string().optional().describe("Task details"),
      due_date: z.string().optional().describe("Due date in ISO format (e.g. 2026-05-20)"),
    },
    async ({ opportunity_id, subject, description, due_date }) => {
      engagementWriteLimiter.check("create_follow_up_task");
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const task = await createTask({ opportunityId: id, subject, description, dueDate: due_date }, progress);
      return { content: [{ type: "text", text: `✅ Task created: **${task.subject}**${task.scheduledend ? `\nDue: ${task.scheduledend.slice(0, 10)}` : ""}` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: complete_activity
  // ---------------------------------------------------------------------------
  server.tool(
    "complete_activity",
    `Mark an activity (appointment, phone call, or task) as complete.

Use list_activities first to find the activity ID. Confirm with the user before completing.`,
    {
      activity_type: z.enum(["appointment", "phonecall", "task"]).describe("Type of activity to complete"),
      activity_id: z.string().describe("Dynamics activity GUID"),
    },
    async ({ activity_type, activity_id }) => {
      engagementWriteLimiter.check("complete_activity");
      const progress = makeProgress(server);
      await completeActivity(activity_type, activity_id, progress);
      return { content: [{ type: "text", text: `✅ ${activity_type} marked as complete.` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: list_opportunity_contacts
  // ---------------------------------------------------------------------------
  server.tool(
    "list_opportunity_contacts",
    `List contacts (stakeholders) linked to an opportunity.

Shows the stakeholder map: who's involved and their role (Champion, Economic Buyer, Decision Maker, etc.).`,
    { opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number") },
    async ({ opportunity_id }) => {
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const contacts = await listOpportunityContacts(id, progress);
      return { content: [{ type: "text", text: externalData("Dynamics contacts", contacts) }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: create_contact
  // ---------------------------------------------------------------------------
  server.tool(
    "create_contact",
    `Create a new contact in Dynamics 365. Use search_contacts first to check if the contact already exists.`,
    {
      first_name: z.string().describe("Contact first name"),
      last_name: z.string().describe("Contact last name"),
      email: z.string().optional().describe("Email address"),
      job_title: z.string().optional().describe("Job title (e.g. CTO, VP Engineering)"),
      phone: z.string().optional().describe("Phone number"),
      account_id: z.string().optional().describe("Parent account GUID — link the contact to this company"),
    },
    async ({ first_name, last_name, email, job_title, phone, account_id }) => {
      engagementWriteLimiter.check("create_contact");
      const progress = makeProgress(server);
      const contact = await createContact({ firstName: first_name, lastName: last_name, email, jobTitle: job_title, phone, accountId: account_id }, progress);
      return { content: [{ type: "text", text: `✅ Created contact: **${contact.fullname}**${email ? ` (${email})` : ""}${contact.jobtitle ? ` — ${contact.jobtitle}` : ""}` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: add_contact_to_opportunity
  // ---------------------------------------------------------------------------
  server.tool(
    "add_contact_to_opportunity",
    `Link a contact to an opportunity as a stakeholder. Optionally specify their role.

Roles: Champion, Economic Buyer, Technical Buyer, Coach, Decision Maker, Influencer, End User, Executive Sponsor.

Use search_contacts to find the contact ID first, then link them.`,
    {
      contact_id: z.string().describe("Dynamics contact GUID"),
      opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number"),
      role: z.string().optional().describe("Stakeholder role (e.g. Champion, Economic Buyer, Decision Maker)"),
    },
    async ({ contact_id, opportunity_id, role }) => {
      engagementWriteLimiter.check("add_contact_to_opportunity");
      const progress = makeProgress(server);
      requireGuid(contact_id, "contact_id");
      const oppId = await resolveOpportunityId(opportunity_id, progress);
      await addContactToOpportunity(contact_id, oppId, role, progress);
      return { content: [{ type: "text", text: `✅ Contact linked to opportunity${role ? ` as **${role}**` : ""}.` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: list_closing_plan
  // ---------------------------------------------------------------------------
  server.tool(
    "list_closing_plan",
    `View the closing plan milestones for an opportunity.

Shows structured milestones with due dates, status, and ownership. Great for tracking deal progress.
Set include_completed=true to see completed milestones too (default: only open).`,
    {
      opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number"),
      include_completed: z.boolean().optional().describe("Include completed milestones (default false)"),
    },
    async ({ opportunity_id, include_completed }) => {
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const milestones = await listClosingPlan(id, progress, { includeCompleted: include_completed });
      if (milestones.length === 0) {
        return { content: [{ type: "text", text: "No closing plan milestones found for this opportunity." }] };
      }
      return { content: [{ type: "text", text: externalData("Dynamics closing plan", milestones) }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: add_closing_plan_milestone
  // ---------------------------------------------------------------------------
  server.tool(
    "add_closing_plan_milestone",
    `Add a milestone to an opportunity's closing plan.

Creates a new milestone with a title, optional due date, and description.
Confirm with the user before creating.`,
    {
      opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number"),
      title: z.string().describe("Milestone title (e.g. 'Technical validation complete')"),
      due_date: z.string().optional().describe("Due date in ISO format (e.g. 2026-05-15)"),
      description: z.string().optional().describe("Milestone description or notes"),
      confirmed: z.boolean().describe("User must confirm before creating"),
    },
    async ({ opportunity_id, title, due_date, description, confirmed }) => {
      if (!confirmed) return { content: [{ type: "text", text: "⚠️ Please confirm to create this milestone." }] };
      engagementWriteLimiter.check("add_closing_plan_milestone");
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const milestone = await createClosingPlanMilestone({ opportunityId: id, title, dueDate: due_date, description }, progress);
      return { content: [{ type: "text", text: `✅ Milestone created: **${milestone.title}**${milestone.dueDate ? ` (due: ${milestone.dueDate.slice(0, 10)})` : ""}` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: update_closing_plan_milestone
  // ---------------------------------------------------------------------------
  server.tool(
    "update_closing_plan_milestone",
    `Update a closing plan milestone — mark complete, flag at risk, or change details.

Use list_closing_plan first to find the milestone ID.`,
    {
      milestone_id: z.string().describe("Closing plan milestone GUID"),
      complete: z.boolean().optional().describe("Set true to mark milestone as complete"),
      at_risk: z.boolean().optional().describe("Set true to flag as at risk, false to remove flag"),
      title: z.string().optional().describe("New title"),
      due_date: z.string().optional().describe("New due date in ISO format"),
      description: z.string().optional().describe("Updated description"),
    },
    async ({ milestone_id, complete, at_risk, title, due_date, description }) => {
      engagementWriteLimiter.check("update_closing_plan_milestone");
      const progress = makeProgress(server);
      await updateClosingPlanMilestone(milestone_id, {
        complete, atRisk: at_risk, title, dueDate: due_date, description,
      }, progress);
      const actions: string[] = [];
      if (complete) actions.push("marked complete");
      if (at_risk === true) actions.push("flagged at risk");
      if (at_risk === false) actions.push("risk flag removed");
      if (title) actions.push("title updated");
      if (due_date) actions.push("due date updated");
      return { content: [{ type: "text", text: `✅ Milestone updated: ${actions.join(", ") || "changes saved"}.` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: get_opportunity_summary
  // ---------------------------------------------------------------------------
  server.tool(
    "get_opportunity_summary",
    `Read the opportunity summary / deal review notes.

Returns summary content, notes, and annotations attached to the opportunity.
Useful for understanding deal context, history, and current status.`,
    { opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number") },
    async ({ opportunity_id }) => {
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const summaries = await getOpportunitySummary(id, progress);
      if (summaries.length === 0) {
        return { content: [{ type: "text", text: "No opportunity summary or notes found." }] };
      }
      return { content: [{ type: "text", text: externalData("Dynamics opportunity summary", summaries) }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: generate_opportunity_summary
  // ---------------------------------------------------------------------------
  server.tool(
    "generate_opportunity_summary",
    `Collect all CRM data for an opportunity and generate a comprehensive AI deal summary.

Fetches in parallel: opportunity details, all engagements, timeline notes (last 30),
contacts/stakeholders, activities, collaboration team, and any existing summary.

After receiving the data, generate a structured deal summary with these sections:
**Deal Overview** — name, account, stage, NNACV/ACV, close date, BU, forecast category
**Team** — SC, AE, collaboration team members with roles
**Engagement History** — all completed and active milestones grouped by type;
  note what's complete ✅ and what's in progress 🔄 or missing
**Stakeholders** — customer contacts with their role/title
**Recent Activity** — last 5 timeline notes and activities, newest first
**Current Status** — one-paragraph honest deal assessment: where we are, key risks, blockers
**Next Steps** — 3–5 concrete actions the SC/AE should take

Format for readability. After generating, ask the user: "Save this to Dynamics as the opportunity summary?" — if yes, call update_opportunity_summary with confirmed=true.`,
    { opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number") },
    async ({ opportunity_id }) => {
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);

      const [opp, engagements, notes, contacts, activities, team, existingSummary] = await Promise.all([
        fetchOpportunityById(id, progress).catch(() => null),
        fetchEngagementsByOpportunity(id, progress).catch(() => []),
        listTimelineNotes(id, progress).catch(() => []),
        listOpportunityContacts(id, progress).catch(() => []),
        listActivities(id, progress, { includeCompleted: true, top: 30 }).catch(() => []),
        fetchCollaborationTeam(id, progress).catch(() => []),
        getOpportunitySummary(id, progress).catch(() => []),
      ]);

      if (!opp) {
        return { content: [{ type: "text", text: "❌ Opportunity not found." }] };
      }

      const data = {
        opportunity: opp,
        engagements,
        timelineNotes: notes.slice(0, 30),
        contacts,
        activities: activities.slice(0, 30),
        collaborationTeam: team,
        existingSummary: existingSummary.length > 0 ? existingSummary[0] : null,
      };

      return { content: [{ type: "text", text: externalData("Dynamics opportunity context", data) }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: update_opportunity_summary
  // ---------------------------------------------------------------------------
  server.tool(
    "update_opportunity_summary",
    `Write or update the opportunity summary / deal review.

Updates an existing summary if one exists, or creates a new one.
Confirm with the user before writing.`,
    {
      opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number"),
      summary: z.string().describe("The summary content to write"),
      title: z.string().optional().describe("Summary title (default: 'Opportunity Summary')"),
      confirmed: z.boolean().describe("User must confirm before writing"),
    },
    async ({ opportunity_id, summary, title, confirmed }) => {
      if (!confirmed) return { content: [{ type: "text", text: "⚠️ Please confirm to write this summary." }] };
      engagementWriteLimiter.check("update_opportunity_summary");
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const result = await updateOpportunitySummary(id, summary, title, progress);
      return { content: [{ type: "text", text: `✅ Summary saved: **${result.title}**` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: list_quotes
  // ---------------------------------------------------------------------------
  server.tool(
    "list_quotes",
    `List quotes linked to an opportunity. Shows quote name, status, value, and dates.`,
    { opportunity_id: z.string().describe("Dynamics opportunity GUID or OPTY number") },
    async ({ opportunity_id }) => {
      const progress = makeProgress(server);
      const id = await resolveOpportunityId(opportunity_id, progress);
      const quotes = await listQuotes(id, progress);
      if (quotes.length === 0) {
        return { content: [{ type: "text", text: "No quotes found for this opportunity." }] };
      }
      const lines = quotes.map(q => {
        const val = q.totalAmount != null ? `$${q.totalAmount.toLocaleString()}` : "no value";
        const dates = [q.effectiveFrom?.slice(0, 10), q.effectiveTo?.slice(0, 10)].filter(Boolean).join(" → ");
        return `- **${q.name}**${q.quoteNumber ? ` (${q.quoteNumber})` : ""} | ${q.status} | ${val}${dates ? ` | ${dates}` : ""}`;
      });
      return { content: [{ type: "text", text: lines.join("\n") }] };
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
      const id = requireGuid(engagement_id, "engagement_id");
      const progress = makeProgress(server);
      const engagement = await fetchEngagementById(id, progress);
      return { content: [{ type: "text", text: engagementListItem(engagement) }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: search_my_engagements
  // ---------------------------------------------------------------------------
  server.tool(
    "search_my_engagements",
    `Find all engagements where the current user is listed as a participant (Active Participants / sn_engagementassignees).

This is the primary search tool for SSCs and Specialists who are added to engagements
but are not the engagement owner. Also useful for SCs to find engagements they collaborate on.

Supports filtering by engagement type (e.g. "Demo", "Discovery"), status (open/complete/all),
and free-text search on engagement name.`,
    {
      search: z.string().optional().describe("Filter by engagement name (partial match)"),
      engagement_type: z.string().optional().describe("Filter by type, e.g. 'Demo', 'Discovery', 'POV'"),
      status: z.enum(["open", "complete", "all"]).optional().describe("Filter by status (default: all)"),
      top: z.number().optional().describe("Max results (default 50)"),
    },
    async ({ search, engagement_type, status, top }) => {
      const progress = makeProgress(server);
      const engagements = await fetchMyEngagementAssignments(
        { search, engagementType: engagement_type, status, top },
        progress
      );
      if (engagements.length === 0) {
        return { content: [{ type: "text", text: "No engagements found where you are a participant." }] };
      }
      const lines = engagements.map(e => engagementListItem(e));
      return { content: [{ type: "text", text: `Found ${engagements.length} engagement(s):\n\n${lines.join("\n\n---\n\n")}` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: search_my_collaboration_opportunities
  // ---------------------------------------------------------------------------
  server.tool(
    "search_my_collaboration_opportunities",
    `Find all open opportunities where the current user is on the Collaboration Team.

This returns opportunities where you have been added as a collaborator (Solution Consultant,
Specialist, Renewal AM, etc.) — even if you are not the primary SC or owner.
Ideal for SSCs and Specialists to see their full pipeline.`,
    {},
    async () => {
      const progress = makeProgress(server);
      const opps = await fetchMyCollaborationOpportunities(progress);
      if (opps.length === 0) {
        return { content: [{ type: "text", text: "No open opportunities found where you are on the collaboration team." }] };
      }
      const lines = opps.map(o => {
        const link = `${DYNAMICS_BASE_URL}/main.aspx?etn=opportunity&id=${o.opportunityid}&pagetype=entityrecord`;
        return [
          `**${o.name}** (${o.sn_number ?? "—"})`,
          `Account: ${o.accountName} · Owner: ${o.ownerName ?? "—"} · SC: ${o.scName ?? "—"}`,
          `Close: ${o.estimatedclosedate?.slice(0, 10) ?? "—"} · NNACV: ${o.nnacv != null ? `$${o.nnacv.toLocaleString()}` : "—"} · ${o.forecastCategoryName ?? "—"}`,
          `🔗 ${link}`,
        ].join("\n");
      });
      return { content: [{ type: "text", text: `Found ${opps.length} opportunity/ies:\n\n${lines.join("\n\n---\n\n")}` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: get_engagement_participants
  // ---------------------------------------------------------------------------
  server.tool(
    "get_engagement_participants",
    `View the Active Participants on an engagement — lists all SCs/SSCs/Specialists
assigned to a specific engagement record.`,
    { engagement_id: z.string().describe("Dynamics sn_engagement GUID") },
    async ({ engagement_id }) => {
      const id = requireGuid(engagement_id, "engagement_id");
      const progress = makeProgress(server);
      const participants = await fetchEngagementParticipants(id, progress);
      if (participants.length === 0) {
        return { content: [{ type: "text", text: "No participants found on this engagement." }] };
      }
      const lines = participants.map(p =>
        `• **${p.userName}**${p.title ? ` — ${p.title}` : ""}${p.isPrimary ? " ⭐ Primary" : ""}`
      );
      return { content: [{ type: "text", text: `Active Participants (${participants.length}):\n\n${lines.join("\n")}` }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: post_teams_notification
  // ---------------------------------------------------------------------------
  server.tool(
    "post_teams_notification",
    "Post a SHORT notification to the configured Teams channel. Use only for status messages like 'Engagement logged' — do NOT use this to post CRM data, opportunity details, pipeline values, or customer information. Requires configure_teams_webhook to be set up first.",
    {
      title: z.string().max(100).describe("Notification title (max 100 chars)"),
      body:  z.string().max(500).describe("Notification body — brief status only, no CRM data (max 500 chars)"),
    },
    async ({ title, body }) => {
      if (title.length > 100 || body.length > 500) {
        return { content: [{ type: "text", text: "❌ Message too long. Title max 100 chars, body max 500 chars." }] };
      }
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

Requires Alfred to be running with Teams or Outlook open. The Graph token is captured automatically.`,
    {
      search:     z.string().optional().describe("Keyword to match meeting subject (e.g. 'Fabrikam ICW')"),
      start_date: z.string().optional().describe("Search from this date (ISO, e.g. 2026-01-01) — defaults to 30 days ago"),
      end_date:   z.string().optional().describe("Search to this date (ISO) — defaults to today"),
      meeting_id: z.string().optional().describe("If already known, fetch transcript for this specific meeting ID directly"),
    },
    async ({ search, start_date, end_date, meeting_id }) => {
      const progress = makeProgress(server);
      const transcripts = await getTeamsTranscript({ search, startDate: start_date, endDate: end_date, meetingId: meeting_id }, progress);

      if (transcripts.length === 0) {
        return { content: [{ type: "text", text: "No meetings with transcripts found." }] };
      }

      const formatted = transcripts.map(t => {
        const date = t.start ? new Date(t.start).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" }) : "—";
        const time = t.start ? new Date(t.start).toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" }) : "";
        const duration = t.start && t.end
          ? `${Math.round((new Date(t.end).getTime() - new Date(t.start).getTime()) / 60000)} min`
          : "";
        const attendeeList = t.attendees.length ? t.attendees.join(", ") : "—";

        const lines = [
          `## ${t.subject ?? "Untitled Meeting"}`,
          `📅 ${date}${time ? ` at ${time}` : ""}${duration ? ` · ${duration}` : ""}`,
          `👥 ${attendeeList}`,
        ];

        if (t.transcript) {
          lines.push("", "**Transcript:**", t.transcript);
        } else {
          lines.push("", "_No transcript available for this meeting._");
        }

        return lines.join("\n");
      }).join("\n\n---\n\n");

      return { content: [{ type: "text", text: externalData("Teams transcripts", formatted) }] };
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

Always runs for the current user's pipeline only. Optionally posts results to Teams.

CROSS-REFERENCE: After presenting results, use Data_Analytics_Connection account_insights to check customer health scores, license utilization, and renewal status for the accounts in the sweep. This adds context on which red/yellow accounts are also at-risk from a health perspective.`,
    {
      post_to_teams: z.boolean().optional().describe("Post the report to Teams (requires configure_teams_webhook)"),
      min_nnacv:     z.number().optional().describe("Minimum NNACV filter in USD (default $100K). $0 opportunities are always excluded."),
      exclude_app_store: z.boolean().optional().describe("Exclude App Store Renewal noise opportunities (default true)"),
    },
    async ({ post_to_teams, min_nnacv, exclude_app_store }) => {
      const progress = makeProgress(server);
      const results = await runHygieneSweep({
        postToTeams: post_to_teams ?? false,
        minNnacv: min_nnacv ?? 100_000,
        excludeAppStore: exclude_app_store,
      }, progress);
      return { content: [{ type: "text", text: formatHygieneReport(results) }] };
    }
  );

  // ---------------------------------------------------------------------------
  // Tool: uninstall_alfred
  // ---------------------------------------------------------------------------
  server.tool(
    "uninstall_alfred",
    `Uninstall Alfred from this machine. IMPORTANT: Always confirm with the user before running this.

Removes: cron jobs, Claude Desktop config entry, Alfred.app, version cache.
Optional: config file (~/.alfred-config.json), Chrome profile (~/.alfred-profile), log files.

After uninstall, the user must restart Claude Desktop.`,
    {
      remove_config: z.boolean().optional().describe("Also remove ~/.alfred-config.json (Dynamics URL, webhook, role). Default false."),
      remove_chrome_profile: z.boolean().optional().describe("Also remove ~/.alfred-profile (Chrome session, cookies). Default false."),
      remove_logs: z.boolean().optional().describe("Also remove ~/.alfred-hygiene.log and ~/.alfred-meetings.log. Default false."),
    },
    async ({ remove_config, remove_chrome_profile, remove_logs }) => {
      const progress = makeProgress(server);
      const results: string[] = [];
      const { unlinkSync, rmSync } = await import("fs");

      // 1. Remove cron jobs
      try {
        const currentCron = execFileSync("crontab", ["-l"], { encoding: "utf8", timeout: 5_000 }).trim();
        const newCron = currentCron.split("\n").filter(l => !l.includes("hygiene-sweep") && !l.includes("post-meeting-sweep")).join("\n");
        if (newCron !== currentCron) {
          execFileSync("crontab", ["-"], { input: newCron, timeout: 5_000 });
          results.push("✅ Cron jobs removed");
        } else {
          results.push("ℹ️ No Alfred cron jobs found");
        }
      } catch { results.push("ℹ️ No crontab (skipped)"); }

      // 2. Remove from Claude Desktop config
      try {
        const claudeConfig = join(homedir(), "Library", "Application Support", "Claude", "claude_desktop_config.json");
        if (existsSync(claudeConfig)) {
          const config = JSON.parse(readFileSync(claudeConfig, "utf8"));
          let removed = false;
          for (const key of ["alfred", "sc-engagement", "alfred-sales"]) {
            if (config.mcpServers?.[key]) { delete config.mcpServers[key]; removed = true; }
          }
          if (removed) {
            writeFileSync(claudeConfig, JSON.stringify(config, null, 2));
            results.push("✅ Removed from Claude Desktop config");
          } else {
            results.push("ℹ️ No Alfred entry in Claude Desktop config");
          }
        }
      } catch (e) { results.push(`⚠️ Could not update Claude Desktop config: ${e instanceof Error ? e.message : String(e)}`); }

      // 3. Remove Alfred.app
      try {
        const appPath = join(homedir(), "Desktop", "Alfred.app");
        if (existsSync(appPath)) { rmSync(appPath, { recursive: true }); results.push("✅ Alfred.app removed"); }
        else { results.push("ℹ️ Alfred.app not found"); }
      } catch (e) { results.push(`⚠️ Could not remove Alfred.app: ${e instanceof Error ? e.message : String(e)}`); }

      // 4. Remove version cache
      try {
        const cache = join(homedir(), ".alfred-version-check");
        if (existsSync(cache)) { unlinkSync(cache); results.push("✅ Version cache removed"); }
      } catch { /* non-fatal */ }

      // 5. Optional: config file
      if (remove_config) {
        try {
          const cfg = join(homedir(), ".alfred-config.json");
          if (existsSync(cfg)) { unlinkSync(cfg); results.push("✅ Config file removed"); }
        } catch (e) { results.push(`⚠️ Could not remove config: ${e instanceof Error ? e.message : String(e)}`); }
      }

      // 6. Optional: Playwright browser profile (~/.alfred-pw)
      if (remove_chrome_profile) {
        try {
          const profile = join(homedir(), ".alfred-pw");
          if (existsSync(profile)) { rmSync(profile, { recursive: true }); results.push("✅ Browser profile removed (~/.alfred-pw)"); }
        } catch (e) { results.push(`⚠️ Could not remove browser profile: ${e instanceof Error ? e.message : String(e)}`); }
      }

      // 7. Optional: log files
      if (remove_logs) {
        for (const log of [".alfred-hygiene.log", ".alfred-meetings.log"]) {
          try {
            const logPath = join(homedir(), log);
            if (existsSync(logPath)) { unlinkSync(logPath); results.push(`✅ ${log} removed`); }
          } catch { /* non-fatal */ }
        }
      }

      progress("🗑️ Uninstall complete");
      return { content: [{ type: "text", text:
        `🗑️ **Alfred uninstalled**\n\n${results.join("\n")}\n\n` +
        `${!remove_config ? "ℹ️ Config file kept (~/.alfred-config.json) — re-run with remove_config=true to delete\n" : ""}` +
        `${!remove_chrome_profile ? "ℹ️ Chrome profile kept (~/.alfred-profile) — re-run with remove_chrome_profile=true to delete\n" : ""}` +
        `\n⚠️ **Restart Claude Desktop** to complete the uninstall.`
      }] };
    }
  );

  // Note: update_alfred is intentionally NOT included in common-tools.ts.
  // It uses `import.meta.url` to find the install directory, and the relative
  // path differs between sc/index.ts and sales/index.ts (both resolve to root
  // correctly via "dist/sc/index.js → root" and "dist/sales/index.js → root").
  // To avoid duplicating the path comment and keep the logic clear, update_alfred
  // remains registered separately in each entry point.

  // ---------------------------------------------------------------------------
  // Tool: configure_teams_webhook
  // ---------------------------------------------------------------------------
  server.tool(
    "configure_teams_webhook",
    "Set the Teams incoming webhook URL for notifications. Create one in Teams: channel → ... → Connectors → Incoming Webhook.",
    { webhook_url: z.string().describe("Teams incoming webhook URL") },
    async ({ webhook_url }) => {
      let parsed: URL;
      try { parsed = new URL(webhook_url); } catch {
        return { content: [{ type: "text", text: "❌ Invalid URL." }] };
      }
      if (!parsed.hostname.endsWith(".webhook.office.com")) {
        return { content: [{ type: "text", text: "❌ URL must be a *.webhook.office.com Teams incoming webhook." }] };
      }
      setTeamsWebhook(webhook_url);
      // Persist to config so it's remembered across sessions
      try {
        const fs = await import("fs");
        const os = await import("os");
        const cfgPath = `${os.default.homedir()}/.alfred-config.json`;
        const cfg = JSON.parse(fs.default.readFileSync(cfgPath, "utf-8").toString());
        cfg.teamsWebhook = webhook_url;
        fs.default.writeFileSync(cfgPath, JSON.stringify(cfg, null, 2));
      } catch (e) { process.stderr.write(`[alfred:warn] webhook config persist failed: ${e instanceof Error ? e.message : String(e)}\n`); }
      return { content: [{ type: "text", text: "✅ Teams webhook configured and saved. Notifications will post to that channel." }] };
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
If no transcript, use the meeting subject, attendees and calendar notes for best-effort pre-fill.

CROSS-REFERENCE: For each matched account, use Data_Analytics_Connection account_insights to pull customer health, product subscriptions, and license utilization. This gives context for writing up the engagement — what the customer owns, how much they use it, and any risk signals.`,
    {
      hours_back: z.number().optional().describe("How many hours back to scan for ended meetings (default 24)"),
      search:     z.string().optional().describe("Optional keyword to filter meeting subjects (e.g. 'Fabrikam', 'Acme Corp')"),
      post_to_teams: z.boolean().optional().describe("Post a summary card per candidate to Teams (requires configure_teams_webhook). Default false."),
    },
    async ({ hours_back, search, post_to_teams }) => {
      const progress = makeProgress(server);
      const candidates = await detectPostMeetingEngagements({ hoursBack: hours_back, search }, progress);

      if (candidates.length === 0) {
        return { content: [{ type: "text", text: "No ended online meetings found in the specified window." }] };
      }

      // Post to Teams if requested
      if (post_to_teams) {
        await notifyPostMeetingCandidates(candidates, DYNAMICS_HOST, progress);
      }

      // Strip raw calendarEvent (large Graph API blob Claude doesn't need) and truncate transcripts
      const slim = candidates.map(({ calendarEvent: _raw, transcript, ...c }) => ({
        ...c,
        ...(transcript ? { transcript: transcript.length > 4000 ? transcript.slice(0, 4000) + "\n…[truncated]" : transcript } : {}),
      }));

      return { content: [{ type: "text", text: externalData("Teams calendar + transcripts", slim) }] };
    }
  );
}
