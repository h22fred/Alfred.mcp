import {
  fetchOpportunities,
  fetchEngagementsByOpportunity,
  type Opportunity,
  type Engagement,
} from "./dynamicsClient.js";
import { postTeamsNotification, postAdaptiveCard } from "./teamsClient.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";

// Fallback engagement types if none configured
const DEFAULT_REQUIRED: string[] = ["Discovery", "Demo", "Technical Win"];

// Short column header labels for the Teams card
const TYPE_ABBREV: Record<string, string> = {
  "Discovery": "DIS", "Demo": "Demo", "Technical Win": "TW",
  "RFx": "RFx", "Business Case": "BC", "Workshop": "WS", "POV": "POV", "EBC": "EBC",
  "Opportunity Summary": "OppSum", "Mutual Plan": "MP", "Budget": "Budget",
  "Implementation Plan": "ImpPlan", "Stakeholder Alignment": "StkhAl",
  "Customer Business Review": "CBR", "Post Sale Engagement": "PSE",
};
function abbrev(t: string): string { return TYPE_ABBREV[t] ?? t.slice(0, 4).toUpperCase(); }

function closeDateLabel(date?: string): string {
  if (!date) return "—";
  const d = new Date(date);
  const now = new Date();
  const daysOut = Math.round((d.getTime() - now.getTime()) / 86_400_000);
  const label = d.toLocaleDateString("en-GB", { month: "short", year: "2-digit" });
  if (daysOut < 0)  return `🔴 ${label}`;
  if (daysOut < 30) return `🟡 ${label}`;
  return label;
}

export interface HygieneResult {
  opportunity: Opportunity;
  engagements: Engagement[];
  missingRequired: string[];
  missingOptional: string[];
  status: "red" | "yellow" | "green";
}

/** Patterns that indicate noise opportunities (back-office auto-renewals, App Store, etc.) */
const NOISE_PATTERNS = [/app\s*store\s*renewal/i];

export async function runHygieneSweep(opts: {
  postToTeams?: boolean;
  minNnacv?: number;
  excludeAppStore?: boolean;
  engagementTypes?: string[];
  dynamicsUrl?: string;
}, progress: ProgressFn = () => {}): Promise<HygieneResult[]> {
  progress("🔍 Starting CRM hygiene sweep...");

  const requiredTypes = opts.engagementTypes?.length ? opts.engagementTypes : DEFAULT_REQUIRED;
  const excludeAppStore = opts.excludeAppStore ?? true; // default: skip noise

  const opps = await fetchOpportunities({
    myOpportunitiesOnly: true,
    // includeZeroValue defaults to false — $0 noise excluded automatically
    minNnacv: opts.minNnacv ?? 100_000,
    top: 200,
  }, progress);

  // Client-side filter for noise patterns (App Store renewals, etc.)
  const filtered = excludeAppStore
    ? opps.filter(o => !NOISE_PATTERNS.some(p => p.test(o.name)))
    : opps;

  if (filtered.length < opps.length) {
    progress(`🧹 Filtered ${opps.length - filtered.length} noise opportunities (App Store renewals etc.)`);
  }

  progress(`📋 Checking ${filtered.length} opportunities...`);
  const results: HygieneResult[] = [];

  // Batch engagement fetches in groups of 8 for ~8x speedup over sequential
  const BATCH_SIZE = 8;
  for (let i = 0; i < filtered.length; i += BATCH_SIZE) {
    const batch = filtered.slice(i, i + BATCH_SIZE);
    const batchResults = await Promise.all(
      batch.map(async (opp) => {
        const engagements = await fetchEngagementsByOpportunity(opp.opportunityid, progress);
        const activeEngagements = engagements.filter(e => !e.statusName?.toLowerCase().includes("cancel"));
        const typeNames = activeEngagements.map(e => e.engagementTypeName ?? "").filter(Boolean);

        const missingRequired = requiredTypes.filter(t => !typeNames.includes(t));
        const missingOptional: string[] = [];

        const status: HygieneResult["status"] =
          missingRequired.length > 0 ? "red" : "green";

        return { opportunity: opp, engagements, missingRequired, missingOptional, status } as HygieneResult;
      })
    );
    results.push(...batchResults);
    if (i + BATCH_SIZE < filtered.length) {
      progress(`📋 Checked ${Math.min(i + BATCH_SIZE, filtered.length)}/${filtered.length} opportunities...`);
    }
  }

  // Sort: red first, then green
  results.sort((a, b) => {
    const order = { red: 0, yellow: 1, green: 2 };
    return order[a.status] - order[b.status];
  });

  if (opts.postToTeams) {
    await postHygieneToTeams(results, requiredTypes, opts.dynamicsUrl, progress);
  }

  progress(`✅ Hygiene sweep complete — ${results.filter(r => r.status === "red").length} red, ${results.filter(r => r.status === "yellow").length} yellow, ${results.filter(r => r.status === "green").length} green`);
  return results;
}

function fmt(n?: number): string {
  if (!n) return "—";
  return n >= 1_000_000
    ? `$${(n / 1_000_000).toFixed(1)}M`
    : `$${Math.round(n / 1_000)}K`;
}

/** Truncate to `max` characters including the trailing ellipsis. */
function truncate(s: string, max: number): string {
  return s.length > max ? s.slice(0, max - 1) + "…" : s;  // result is exactly `max` chars
}

// Group results by account, sorted by worst status then total pipeline desc
function groupByAccount(results: HygieneResult[]): Map<string, HygieneResult[]> {
  const map = new Map<string, HygieneResult[]>();
  for (const r of results) {
    const key = r.opportunity.accountName || "Unknown Account";
    if (!map.has(key)) map.set(key, []);
    map.get(key)!.push(r);
  }
  // Sort accounts: worst status first, then by total pipeline desc
  const statusOrder = { red: 0, yellow: 1, green: 2 };
  const sorted = new Map([...map.entries()].sort((a, b) => {
    const aWorst = Math.min(...a[1].map(r => statusOrder[r.status]));
    const bWorst = Math.min(...b[1].map(r => statusOrder[r.status]));
    if (aWorst !== bWorst) return aWorst - bWorst;
    const aTotal = a[1].reduce((s, r) => s + (r.opportunity.totalamount ?? 0), 0);
    const bTotal = b[1].reduce((s, r) => s + (r.opportunity.totalamount ?? 0), 0);
    return bTotal - aTotal;
  }));
  return sorted;
}

function accountBlock(account: string, opps: HygieneResult[]): Record<string, unknown>[] {
  const blocks: Record<string, unknown>[] = [];
  const accountTotal = opps.reduce((s, r) => s + (r.opportunity.totalamount ?? 0), 0);
  const worstStatus = opps.some(r => r.status === "red") ? "red"
    : opps.some(r => r.status === "yellow") ? "yellow" : "green";
  const icon = worstStatus === "red" ? "🔴" : worstStatus === "yellow" ? "🟡" : "✅";

  // Account header row
  blocks.push({
    type: "ColumnSet", spacing: "Medium",
    columns: [
      {
        type: "Column", width: "stretch",
        items: [{ type: "TextBlock", text: `${icon}  **${account.toUpperCase()}**`, size: "Small", weight: "Bolder", wrap: false }],
      },
      {
        type: "Column", width: "auto",
        items: [{ type: "TextBlock", text: fmt(accountTotal), size: "Small", weight: "Bolder", color: "Accent", horizontalAlignment: "Right" }],
      },
    ],
  });

  // One row per opportunity — red sorted by close date asc (urgency), others by NNACV desc
  opps
    .sort((a, b) => {
      const statusOrder = { red: 0, yellow: 1, green: 2 };
      if (a.status !== b.status) return statusOrder[a.status] - statusOrder[b.status];
      if (a.status === "red") {
        const da = a.opportunity.estimatedclosedate ?? "9999";
        const db = b.opportunity.estimatedclosedate ?? "9999";
        return da < db ? -1 : da > db ? 1 : 0;
      }
      return (b.opportunity.totalamount ?? 0) - (a.opportunity.totalamount ?? 0);
    })
    .forEach(r => {
      const oppIcon = r.status === "red" ? "🔴" : r.status === "yellow" ? "🟡" : "✅";
      const missing = r.missingRequired.length
        ? r.missingRequired.join(" · ")
        : r.missingOptional.length
          ? `optional: ${r.missingOptional.slice(0, 3).join(", ")}${r.missingOptional.length > 3 ? "…" : ""}`
          : "complete ✓";
      const closeLabel = r.opportunity.estimatedclosedate
        ? r.opportunity.estimatedclosedate.slice(0, 10)
        : "—";
      blocks.push({
        type: "ColumnSet", spacing: "Small",
        columns: [
          {
            type: "Column", width: "stretch",
            items: [{ type: "TextBlock", text: `${oppIcon}  ${truncate(r.opportunity.name, 38)}`, size: "Small", wrap: false, isSubtle: true }],
          },
          {
            type: "Column", width: "auto",
            items: [{ type: "TextBlock", text: fmt(r.opportunity.totalamount), size: "Small", isSubtle: true, horizontalAlignment: "Right" }],
          },
          {
            type: "Column", width: "auto",
            items: [{ type: "TextBlock", text: closeLabel, size: "Small", isSubtle: true,
              color: r.status === "red" ? "Attention" : "Default", horizontalAlignment: "Right" }],
          },
          {
            type: "Column", width: "auto",
            items: [{
              type: "TextBlock", text: missing, size: "Small", wrap: false,
              color: r.status === "red" ? "Attention" : r.status === "yellow" ? "Warning" : "Good",
              horizontalAlignment: "Right",
            }],
          },
        ],
      });
    });

  return blocks;
}

async function postHygieneToTeams(results: HygieneResult[], requiredTypes: string[], dynamicsUrl: string | undefined, progress: ProgressFn): Promise<void> {
  const today = new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });

  const red    = results.filter(r => r.status === "red").length;
  const yellow = results.filter(r => r.status === "yellow").length;
  const green  = results.filter(r => r.status === "green").length;
  const totalPipeline = results.reduce((s, r) => s + (r.opportunity.totalamount ?? 0), 0);

  // Only show red/yellow — green opps need no action. Cap at 20 rows.
  const actionable = results.filter(r => r.status !== "green");
  const displayResults = actionable.slice(0, 20);
  const hiddenOverCap = actionable.length - displayResults.length;

  const body: Record<string, unknown>[] = [
    { type: "TextBlock", text: `📋 My CRM Hygiene — ${today}`, weight: "Bolder", size: "Large", wrap: true },
    { type: "TextBlock", text: `🔴 ${red} critical  ·  🟡 ${yellow} on track  ·  ✅ ${green} complete  ·  ${fmt(totalPipeline)} pipeline`, size: "Small", wrap: true, spacing: "Small" },
    { type: "TextBlock", text: `_Columns: ${requiredTypes.map(abbrev).join(" · ")}_`, size: "Small", isSubtle: true, wrap: true, spacing: "None" },
  ];

  for (const r of displayResults) {
    const icon = r.status === "red" ? "🔴" : "🟡";
    const typeNames = r.engagements.map(e => e.engagementTypeName ?? "");
    const checks = requiredTypes.map(t => `${abbrev(t)} ${typeNames.includes(t) ? "✅" : "❌"}`).join("  ");
    const close = closeDateLabel(r.opportunity.estimatedclosedate);
    const oppLink = dynamicsUrl
      ? `[↗ Dynamics](${dynamicsUrl}/main.aspx?etn=opportunity&pagetype=entityrecord&id=${r.opportunity.opportunityid})`
      : "";
    body.push(
      { type: "TextBlock", text: `${icon} **${truncate(r.opportunity.name, 40)}** — ${fmt(r.opportunity.totalamount)} · ${close}${oppLink ? `  ${oppLink}` : ""}`, size: "Small", wrap: true, spacing: "Small" },
      { type: "TextBlock", text: checks, size: "Small", isSubtle: true, wrap: true, spacing: "None" },
    );
  }

  const footerParts: string[] = [];
  if (green > 0) footerParts.push(`✅ ${green} complete hidden.`);
  if (hiddenOverCap > 0) footerParts.push(`+${hiddenOverCap} more not shown.`);
  body.push({ type: "TextBlock", text: footerParts.join("  ") || "\u200B", size: "Small", isSubtle: true, separator: true, spacing: "Medium", wrap: true });

  if (actionable.length > 0) {
    body.push({ type: "TextBlock", text: `Ask Claude: _"Create missing engagements for my red opportunities"_`, size: "Small", isSubtle: true, wrap: true, spacing: "Small" });
  }

  await postAdaptiveCard({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body,
    actions: [
      { type: "Action.OpenUrl", title: "Open Claude", url: "https://claude.ai" },
    ],
  }, progress);
}

export function formatHygieneReport(results: HygieneResult[]): string {
  const today = new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
  const red    = results.filter(r => r.status === "red").length;
  const yellow = results.filter(r => r.status === "yellow").length;
  const green  = results.filter(r => r.status === "green").length;
  const totalPipeline = results.reduce((s, r) => s + (r.opportunity.totalamount ?? 0), 0);

  const lines = [
    `**CRM Hygiene — ${today}**`,
    `🔴 ${red} critical · 🟡 ${yellow} on track · ✅ ${green} complete · Pipeline: ${fmt(totalPipeline)}`,
    "",
  ];

  const grouped = groupByAccount(results);

  for (const [account, opps] of grouped) {
    const accountTotal = opps.reduce((s, r) => s + (r.opportunity.totalamount ?? 0), 0);
    const worstIcon = opps.some(r => r.status === "red") ? "🔴"
      : opps.some(r => r.status === "yellow") ? "🟡" : "✅";

    lines.push(`${worstIcon} **${account}** — ${fmt(accountTotal)}`);

    opps
      .sort((a, b) => {
        const statusOrder = { red: 0, yellow: 1, green: 2 };
        if (a.status !== b.status) return statusOrder[a.status] - statusOrder[b.status];
        if (a.status === "red") {
          const da = a.opportunity.estimatedclosedate ?? "9999";
          const db = b.opportunity.estimatedclosedate ?? "9999";
          return da < db ? -1 : da > db ? 1 : 0;
        }
        return (b.opportunity.totalamount ?? 0) - (a.opportunity.totalamount ?? 0);
      })
      .forEach(r => {
        const icon = r.status === "red" ? "🔴" : r.status === "yellow" ? "🟡" : "✅";
        const nnacv = fmt(r.opportunity.totalamount);
        const closeDate = r.opportunity.estimatedclosedate ? ` · close ${r.opportunity.estimatedclosedate.slice(0, 10)}` : "";
        const missing = r.missingRequired.length
          ? `missing: ${r.missingRequired.join(", ")}`
          : r.status === "yellow" ? "required complete ✓" : "all complete ✓";
        lines.push(`   ${icon} ${r.opportunity.name} (${nnacv}${closeDate}) — ${missing}`);
      });

    lines.push("");
  }

  return lines.join("\n");
}
