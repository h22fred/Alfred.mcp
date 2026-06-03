import { fetchOpportunities } from "./dynamicsClient.js";
import { getAuthCookies } from "../auth/tokenExtractor.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";
import { DYNAMICS_HOST } from "../config.js";

export interface StalledDeal {
  oppId: string;
  oppName: string;
  accountName: string;
  nnacv: number;
  closeDate: string | null;
  daysSinceLastEngagement: number | null; // null = no engagement ever
  lastEngagementType: string | null;
  lastEngagementDate: string | null;
  stallScore: number; // days × nnacv — higher = more urgent
}

interface EngagementRow {
  _sn_opportunityid_value: string | null;
  sn_completeddate: string | null;
  createdon: string | null;
  sn_engagementtypeid: { sn_name?: string } | null;
}

export async function detectStalledDeals(
  options: {
    daysThreshold?: number;
    minNnacv?: number;
    myOpportunitiesOnly?: boolean;
  },
  progress: ProgressFn = () => {}
): Promise<StalledDeal[]> {
  const threshold = options.daysThreshold ?? 30;
  const minNnacv  = options.minNnacv ?? 0;

  progress("📊 Fetching open opportunities...");
  const opps = await fetchOpportunities(
    { myOpportunitiesOnly: options.myOpportunitiesOnly ?? true },
    progress
  );

  const eligible = opps.filter(o => (o.nnacv ?? 0) >= minNnacv);
  if (eligible.length === 0) return [];

  progress("🔍 Fetching engagement history...");

  const cookieHeader = await getAuthCookies(progress);
  const apiBase = `${DYNAMICS_HOST}/api/data/v9.2`;

  // Fetch all non-cancelled engagements, newest first.
  // $top=2000 covers any realistic pipeline; we only need the most recent per opp.
  const url =
    `${apiBase}/sn_engagements` +
    `?$select=_sn_opportunityid_value,sn_completeddate,createdon` +
    `&$expand=sn_engagementtypeid($select=sn_name)` +
    `&$filter=statecode ne 2` +
    `&$orderby=sn_completeddate desc` +
    `&$top=2000`;

  const res = await fetch(url, {
    headers: { Cookie: cookieHeader, Accept: "application/json", "OData-MaxVersion": "4.0", "OData-Version": "4.0" },
  });

  const lastEngagement = new Map<string, { date: string; type: string }>();

  if (res.ok) {
    const data = await res.json() as { value: EngagementRow[] };
    for (const eng of data.value ?? []) {
      const oppId = eng._sn_opportunityid_value;
      if (!oppId) continue;
      if (lastEngagement.has(oppId)) continue; // already have the most recent (ordered desc)
      const dateStr = eng.sn_completeddate ?? eng.createdon;
      if (!dateStr) continue;
      lastEngagement.set(oppId, {
        date: dateStr.slice(0, 10),
        type: eng.sn_engagementtypeid?.sn_name ?? "Unknown",
      });
    }
  } else {
    progress(`⚠️ Engagement fetch returned ${res.status} — stall dates unavailable, showing opps with no recorded activity`);
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const results: StalledDeal[] = [];

  for (const opp of eligible) {
    const last = lastEngagement.get(opp.opportunityid);
    let daysSince: number | null = null;

    if (last) {
      const lastDate = new Date(last.date);
      daysSince = Math.floor((today.getTime() - lastDate.getTime()) / (1000 * 60 * 60 * 24));
    }

    const isStalled = last === undefined || (daysSince !== null && daysSince >= threshold);
    if (!isStalled) continue;

    const nnacv = opp.nnacv ?? 0;
    const stallScore = daysSince === null
      ? 999_999_999 // never touched — always sort first
      : daysSince * Math.max(nnacv, 1);

    results.push({
      oppId:                    opp.opportunityid,
      oppName:                  opp.name,
      accountName:              opp.accountName ?? "—",
      nnacv,
      closeDate:                opp.estimatedclosedate ?? null,
      daysSinceLastEngagement:  daysSince,
      lastEngagementType:       last?.type ?? null,
      lastEngagementDate:       last?.date ?? null,
      stallScore,
    });
  }

  // Highest score (longest stall × biggest deal) first
  results.sort((a, b) => b.stallScore - a.stallScore);
  return results;
}
