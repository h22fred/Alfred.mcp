import { getAuthCookies, clearMemoryAuthCache, type ProgressFn } from "../auth/tokenExtractor.js";
import { clearCachedAuthFile } from "../auth/authFileCache.js";
import { userInfo } from "os";
import { DYNAMICS_HOST, ENGAGEMENT_TYPE_GUIDS, type EngagementType } from "../config.js";
import { FORECAST_NAMES, requireGuid, SN_INTERNAL_DOMAINS } from "../shared.js";

const DYNAMICS_BASE = `${DYNAMICS_HOST}/api/data/v9.2`;

// Startup diagnostic — log the resolved base URL to stderr so config issues are visible
process.stderr.write(`[alfred:config] Dynamics base URL: ${DYNAMICS_BASE}\n`);

// ---------------------------------------------------------------------------
// Security helpers
// ---------------------------------------------------------------------------

/**
 * Sanitize a user-supplied search string for safe use inside OData contains().
 * Strips characters that could break out of the string context (parentheses,
 * slashes, OData operators) and escapes single quotes.
 */
export function sanitizeODataSearch(input: string): string {
  // Allow only safe characters: letters, digits, spaces, hyphens, dots, @ _ #
  const stripped = input.replace(/[^a-zA-Z0-9 \-\.@_#]/g, "");
  // Escape remaining single quotes (belt-and-suspenders)
  return stripped.replace(/'/g, "''").slice(0, 100);
}

/** Structured audit log — written to stderr so it appears in log files but not MCP responses. */
function auditLog(action: string, details: Record<string, unknown> = {}): void {
  try {
    const entry = {
      timestamp: new Date().toISOString(),
      user: userInfo().username,
      instance: DYNAMICS_HOST,
      action,
      ...details,
    };
    process.stderr.write(`[alfred:audit] ${JSON.stringify(entry)}\n`);
  } catch (e) { process.stderr.write(`[alfred:warn] audit log failed: ${e instanceof Error ? e.message : String(e)}\n`); }
}

export type { EngagementType };

export interface Opportunity {
  opportunityid: string;
  sn_number?: string;   // OPTY#### format
  name: string;
  accountName: string;
  accountid: string;
  statuscode: number;
  statusName?: string;
  estimatedclosedate?: string;
  msdyn_forecastcategory?: number;
  forecastCategoryName?: string;
  ownerName?: string;
  scName?: string;
  totalamount?: number;  // Total ACV (full contract value)
  nnacv?: number;        // Net New ACV (sn_netnewacv) — the real NNACV
  // Extended fields
  opportunityType?: string;       // e.g. "Order Reset", "New Business"
  salesStage?: string;            // e.g. "7 - Deal Imminent"
  probability?: number;           // e.g. 90
  businessUnitList?: string;      // e.g. "Customer Service, ITSM, AI Platform Foundations, Impact"
  dealChampion?: string;
  industrySolution?: string;
  description?: string;           // opportunity description / notes
  isCompetitive?: boolean;
  winLossReason?: string;
  winLossNotes?: string;
  territoryName?: string;           // e.g. "CHLPC-TER-6" or "LUX-CPG-Switzerland"
  /** True when scName doesn't match the authenticated user — field may be stale */
  scNameMismatch?: boolean;
}

export interface Product {
  productid: string;
  name: string;
  productnumber?: string;
}

export interface Engagement {
  sn_engagementid?: string;
  sn_engagementnumber?: string;
  sn_name: string;
  engagementTypeName?: string;        // resolved from _sn_engagementtypeid_value
  _sn_engagementtypeid_value?: string;
  sn_description?: string;           // notes/description
  sn_completeddate?: string;
  sn_categorycode?: number;
  categoryName?: string;
  sn_salesstagecode?: number;
  salesStageName?: string;
  statecode?: number;
  statuscode?: number;
  statusName?: string;
  _sn_opportunityid_value?: string;
  opportunityName?: string;
  _sn_accountid_value?: string;
  accountName?: string;
  _sn_primaryproductid_value?: string;
  primaryProductName?: string;
  _ownerid_value?: string;
  ownerName?: string;
  createdon?: string;
  modifiedon?: string;
}

async function dynamicsFetch(path: string, options: RequestInit = {}, progress: ProgressFn = () => {}, _retryCount = 0): Promise<Response> {
  const cookieHeader = await getAuthCookies(progress);

  const headers: Record<string, string> = {
    Cookie: cookieHeader,
    "Content-Type": "application/json",
    Accept: "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    'Prefer': 'odata.include-annotations="*"',
    ...(options.headers as Record<string, string> ?? {}),
  };

  const url = path.startsWith("http") ? path : `${DYNAMICS_BASE}${path}`;
  const response = await fetch(url, { ...options, headers, signal: AbortSignal.timeout(30_000) });

  // Handle 429 throttling with exponential backoff
  if (response.status === 429 && _retryCount < 3) {
    const retryAfter = parseInt(response.headers.get("Retry-After") ?? "", 10);
    const delayMs = (retryAfter > 0 ? retryAfter * 1000 : 1000 * Math.pow(2, _retryCount));
    progress(`⏳ Throttled by Dynamics (429) — retrying in ${(delayMs / 1000).toFixed(0)}s (attempt ${_retryCount + 1}/3)...`);
    await new Promise(r => setTimeout(r, delayMs));
    return dynamicsFetch(path, options, progress, _retryCount + 1);
  }

  if (response.status === 401) {
    // Session expired — clear only Dynamics cookies (not Graph/Teams/Outlook tokens)
    clearMemoryAuthCache();
    clearCachedAuthFile("dynamics");
    progress("🔄 Dynamics session expired — re-acquiring cookies...");
    let freshCookie: string;
    try {
      freshCookie = await Promise.race([
        getAuthCookies(progress),
        new Promise<never>((_, reject) => setTimeout(() => reject(new Error("Auth refresh timed out after 30s — is Alfred running?")), 30_000)),
      ]);
    } catch (e) {
      throw new Error(`Auth refresh failed: ${e instanceof Error ? e.message : String(e)}`);
    }
    const retry = await fetch(url, {
      ...options,
      headers: { ...headers, Cookie: freshCookie },
      signal: AbortSignal.timeout(30_000),
    });
    if (!retry.ok) {
      let msg = `Dynamics API error: ${retry.status} ${retry.statusText}`;
      try { const b = await retry.json(); if (b?.error?.message) msg += ` — ${b.error.message}`; } catch (e) { process.stderr.write(`[alfred:warn] retry error body parse failed: ${e instanceof Error ? e.message : String(e)}\n`); }
      throw new Error(msg);
    }
    return retry;
  }

  if (!response.ok) {
    let msg = `Dynamics API error: ${response.status} ${response.statusText}`;
    let rawDetail = "";
    try {
      const ct = response.headers.get("content-type") ?? "";
      if (ct.includes("json")) {
        const body = await response.json();
        if (body?.error?.message) { rawDetail = body.error.message; msg += ` — ${rawDetail}`; }
      } else {
        const text = await response.text().catch(() => "");
        if (text) { rawDetail = text.slice(0, 300); msg += ` — ${rawDetail}`; }
      }
    } catch (e) { process.stderr.write(`[alfred:warn] error body parse failed: ${e instanceof Error ? e.message : String(e)}\n`); }

    // Translate known Dynamics errors to actionable messages
    if (response.status === 403 && rawDetail.includes("missing prv")) {
      const privMatch = rawDetail.match(/missing (prv\w+) privilege/);
      const entityMatch = rawDetail.match(/entity '(\w+)'/);
      msg = `❌ Permission denied: your Dynamics role doesn't have ${privMatch?.[1] ?? "the required"} privilege on ${entityMatch?.[1] ?? "this entity"}. Ask your CRM admin to grant this permission.`;
    }
    if (response.status === 400 && rawDetail.includes("already exists for this opportunity")) {
      const engMatch = rawDetail.match(/(ENG\d+)/);
      msg = `❌ This engagement type already exists on this opportunity. Collaborate on the existing one${engMatch ? ` (${engMatch[1]})` : ""} instead.`;
    }
    throw new Error(msg);
  }

  // Guard against HTML responses on success (e.g. auth redirects returning 200)
  const ct = response.headers.get("content-type") ?? "";
  if (ct && !ct.includes("json") && !ct.includes("octet") && options.method !== "DELETE" && response.status !== 204) {
    const snippet = await response.text().catch(() => "").then(t => t.slice(0, 150));
    throw new Error(
      `Dynamics returned ${ct} instead of JSON — likely a session redirect.\n` +
      `Open Alfred and confirm you are logged into Dynamics.\n` +
      `Response preview: ${snippet}`
    );
  }

  return response;
}

// ---------------------------------------------------------------------------
// Opportunities
// ---------------------------------------------------------------------------

function mapOpportunity(r: Record<string, unknown>): Opportunity {
  const account = r.parentaccountid as { accountid?: string; name?: string } | null;
  const forecastCode = r.msdyn_forecastcategory as number | undefined;
  return {
    opportunityid:       r.opportunityid as string,
    sn_number:           r.sn_number as string | undefined,
    name:                r.name as string,
    accountName:         account?.name ?? "—",
    accountid:           (account?.accountid ?? r._accountid_value) as string,
    statuscode:          r.statuscode as number,
    statusName:          r["statuscode@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    estimatedclosedate:  r.estimatedclosedate as string | undefined,
    msdyn_forecastcategory: forecastCode,
    forecastCategoryName: forecastCode ? (FORECAST_NAMES[forecastCode] ?? String(forecastCode)) : undefined,
    ownerName:           r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    scName:              r["_sn_solutionconsultant_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    totalamount:         r.totalamount as number | undefined,
    nnacv:               r.sn_netnewacv as number | undefined,
    // Extended fields (field names verified against Dynamics EntityDefinitions metadata)
    opportunityType:     r["sn_opportunitytype@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    salesStage:          r["stepname"] as string | undefined,
    probability:         r.closeprobability as number | undefined,
    businessUnitList:    r.sn_opportunitybusinessunitlist as string | undefined,
    dealChampion:        r["_sn_executivesponsor_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    industrySolution:    r["sn_industrysolution@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    description:         r.description as string | undefined,
    isCompetitive:       r.sn_noncompetitive != null ? !(r.sn_noncompetitive as boolean) : undefined,
    winLossReason:       r["sn_winlossnodecisionreason@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    winLossNotes:        r.sn_winlossnodecisionnotes as string | undefined,
    territoryName:       r["_sn_territory_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
  };
}

export interface OpportunityFilter {
  top?: number;        // max results (default 50)
  search?: string;     // filter by account/opportunity name (contains)
  minNnacv?: number;   // minimum NNACV (sn_netnewacv) in USD
  includeZeroValue?: boolean; // include $0 NNACV opps (default: excluded — too much noise)
  myOpportunitiesOnly?: boolean; // filter to current user's owned opportunities
  myOppsFilterField?: "sc" | "owner" | "collab"; // "sc" = _sn_solutionconsultant_value (default), "owner" = _ownerid_value (AEs), "collab" = collaboration team (SSC/Specialists)
  includeClosed?: boolean; // include won/lost/closed opps — default false (open only)
  excludeStale?: boolean; // exclude opps with close date > 6 months past — default true
  ownerSearch?: string; // filter by owner (AE) name — resolves to user IDs
  territoryCode?: string; // filter by territory code (e.g. "CHLPC-TER-6")
  excludeAppStoreRenewals?: boolean; // exclude $0 "App Store Renewal" opps — default false (set true for pipeline views)
}

export async function fetchCurrentUserId(progress: ProgressFn = () => {}): Promise<string> {
  progress("👤 Resolving current user...");
  const whoAmI = await dynamicsFetch("/WhoAmI", {}, progress);
  const { UserId } = await whoAmI.json() as { UserId: string };
  progress(`👤 User: ${UserId}`);
  return UserId;
}

export async function fetchOpportunities(filter: OpportunityFilter = {}, progress: ProgressFn = () => {}): Promise<Opportunity[]> {
  const top = filter.top ?? 50;
  auditLog("fetch_opportunities", { myOpportunitiesOnly: filter.myOpportunitiesOnly ?? true, search: filter.search ?? null, top });
  progress(`📡 Querying Dynamics for open opportunities (max ${top})...`);

  let filterClause = filter.includeClosed ? "statecode ge 0" : "statecode eq 0";
  if (filter.search) {
    const safe = sanitizeODataSearch(filter.search);
    filterClause += ` and (contains(name,'${safe}') or contains(sn_number,'${safe}'))`;
  }
  if (!filter.includeZeroValue) {
    filterClause += ` and sn_netnewacv ne 0`;
  }
  if (filter.minNnacv) {
    // Include opps >= threshold OR negative (negative NNACV is always important)
    filterClause += ` and (sn_netnewacv ge ${filter.minNnacv} or sn_netnewacv lt 0)`;
  }
  // Exclude zombie opps — close date > 6 months in the past (default: on)
  if (filter.excludeStale !== false) {
    const sixMonthsAgo = new Date();
    sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
    const staleDate = sixMonthsAgo.toISOString().slice(0, 10);
    filterClause += ` and (estimatedclosedate eq null or estimatedclosedate ge ${staleDate})`;
  }
  // Territory code filter — resolved to GUID for OData, or post-filtered if resolution fails
  let territoryGuid: string | null = null;
  if (filter.territoryCode) {
    const safeTerr = sanitizeODataSearch(filter.territoryCode);
    // Try to resolve territory code → GUID via the territory entity
    try {
      const terrRes = await dynamicsFetch(
        `/territories?$select=territoryid,name&$filter=contains(name,'${safeTerr}')&$top=3`,
        {}, progress
      );
      if (terrRes.ok) {
        const terrData = await terrRes.json() as { value: { territoryid: string; name: string }[] };
        if (terrData.value?.length === 1) {
          territoryGuid = terrData.value[0]!.territoryid;
          filterClause += ` and _sn_territory_value eq ${territoryGuid}`;
          progress(`📍 Territory: ${terrData.value[0]!.name} (${territoryGuid})`);
        } else if (terrData.value?.length > 1) {
          // Multiple matches — use all
          const ids = terrData.value.map(t => `_sn_territory_value eq ${t.territoryid}`).join(" or ");
          filterClause += ` and (${ids})`;
          progress(`📍 Matched ${terrData.value.length} territories`);
        } else {
          progress(`⚠️ Territory "${filter.territoryCode}" not found — will post-filter`);
        }
      }
    } catch { /* territory entity may not exist — will post-filter */ }
  }
  // Exclude $0 App Store Renewal opps — noise in pipeline views
  if (filter.excludeAppStoreRenewals) {
    filterClause += ` and not(contains(name,'App Store Renewal') and (sn_netnewacv eq null or sn_netnewacv eq 0))`;
  }
  // Track whether we already filtered by collab team (skip post-validation if so)
  let filteredByCollab = false;

  if (filter.myOpportunitiesOnly) {
    const userId = await fetchCurrentUserId(progress);
    requireGuid(userId, "currentUserId");

    if (filter.myOppsFilterField === "collab") {
      // SSC / Sales Specialist: filter to opportunities where user is on the collaboration team
      progress("📡 Finding your collaboration team opportunities...");
      const collabPath =
        `/sn_opportunitycollaborationteams` +
        `?$select=_sn_opportunity_value` +
        `&$filter=_sn_user_value eq ${userId} and statecode eq 0` +
        `&$top=200`;
      const collabRes = await dynamicsFetch(collabPath, {}, progress);
      const collabData = await collabRes.json();
      const oppIds = [...new Set(
        (collabData.value ?? []).map((r: Record<string, unknown>) => r._sn_opportunity_value as string).filter(Boolean)
      )] as string[];

      if (oppIds.length === 0) {
        progress("ℹ️ You are not on any opportunity collaboration teams");
        return [];
      }

      const idFilter = oppIds.map(id => `opportunityid eq ${id}`).join(" or ");
      filterClause += ` and (${idFilter})`;
      filteredByCollab = true;
      progress(`🔍 Found ${oppIds.length} collaboration team opportunities`);
    } else {
      const field = filter.myOppsFilterField === "owner" ? "_ownerid_value" : "_sn_solutionconsultant_value";
      filterClause += ` and (${field} eq '${userId}')`;
    }
  }
  if (filter.ownerSearch) {
    const users = await searchSystemUsers(filter.ownerSearch, progress);
    if (users.length === 0) {
      progress(`⚠️ No users found matching "${filter.ownerSearch}" — returning unfiltered results`);
    } else {
      const ownerConditions = users.map(u => { requireGuid(u.systemuserid, "ownerUserId"); return `_ownerid_value eq '${u.systemuserid}'`; }).join(" or ");
      filterClause += ` and (${ownerConditions})`;
      progress(`👥 Filtering by ${users.length} owner(s): ${users.map(u => u.fullname).join(", ")}`);
    }
  }

  const selectFields = "opportunityid,sn_number,name,_accountid_value,_ownerid_value,_sn_solutionconsultant_value,_sn_territory_value,statuscode,estimatedclosedate,totalamount,sn_netnewacv,msdyn_forecastcategory,stepname,closeprobability,sn_opportunitytype,sn_opportunitybusinessunitlist";

  const path =
    `/opportunities` +
    `?$select=${selectFields}` +
    `&$expand=parentaccountid($select=accountid,name)` +
    `&$filter=${encodeURIComponent(filterClause)}` +
    `&$orderby=estimatedclosedate asc` +
    `&$top=${top}`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  let results = (data.value ?? []).map(mapOpportunity);

  // When filtering to "my opps" by SC field, cross-reference against the collaboration
  // team table which is the authoritative source of SC assignment. The denormalised
  // _sn_solutionconsultant_value field can be stale/incorrect.
  // Skip for owner-based filtering (AEs) and collab-based filtering (already authoritative).
  if (filter.myOpportunitiesOnly && !filteredByCollab && filter.myOppsFilterField !== "owner" && results.length > 0) {
    progress("🔍 Validating against collaboration team...");
    const collabOpps = await fetchMyCollaborationOpportunities(progress);
    const collabIds = new Set(collabOpps.map(o => o.opportunityid));
    // Flag opps where the user isn't on the collab team (stale scName)
    for (const opp of results) {
      if (!collabIds.has(opp.opportunityid)) {
        opp.scNameMismatch = true;
      }
    }
    // Drop opps where user isn't on the collab team at all
    const before = results.length;
    results = results.filter((o: Opportunity) => !o.scNameMismatch);
    if (before > results.length) {
      progress(`⚠️ Removed ${before - results.length} opportunities with stale SC attribution (not on your collaboration team)`);
    }
  }

  // Post-filter: territory name fallback (when GUID resolution didn't find the territory)
  if (filter.territoryCode && !territoryGuid) {
    const terrLower = filter.territoryCode.toLowerCase();
    const beforeTerr = results.length;
    results = results.filter((o: Opportunity) => o.territoryName?.toLowerCase().includes(terrLower));
    if (beforeTerr > results.length) {
      progress(`📍 Territory post-filter: ${results.length} of ${beforeTerr} opps match "${filter.territoryCode}"`);
    }
  }

  progress(`✅ Found ${results.length} opportunities`);
  return results;
}

/** Resolve an opportunity identifier — accepts a GUID or an OPTY number (e.g. OPTY5328326). */
export async function resolveOpportunityId(input: string, progress: ProgressFn = () => {}): Promise<string> {
  // Already a GUID
  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(input)) return input;
  // Looks like an OPTY number — search by sn_number
  const safe = sanitizeODataSearch(input);
  progress(`🔍 Resolving OPTY number "${input}" to GUID...`);
  const res = await dynamicsFetch(
    `/opportunities?$select=opportunityid,sn_number,name&$filter=sn_number eq '${safe}'&$top=1`,
    {}, progress,
  );
  const data = await res.json();
  const match = (data.value as Record<string, unknown>[] ?? [])[0];
  if (!match?.opportunityid) throw new Error(`No opportunity found with number "${input}". Check the OPTY number and try again.`);
  progress(`✅ Resolved ${input} → ${match.name} (${match.opportunityid})`);
  return match.opportunityid as string;
}

export async function fetchOpportunityById(id: string, progress: ProgressFn = () => {}): Promise<Opportunity> {
  progress(`📡 Fetching opportunity ${id}...`);
  const baseFields = "opportunityid,sn_number,name,_accountid_value,_ownerid_value,_sn_solutionconsultant_value,_sn_territory_value,statuscode,estimatedclosedate,totalamount,sn_netnewacv,msdyn_forecastcategory,stepname,closeprobability,sn_opportunitytype,sn_opportunitybusinessunitlist";
  // Enrichment fields for single-opp detail view — may not exist in all instances
  // sn_industrysolution is Virtual — excluded from $select, may come through via annotations
  const enrichFields = ",_sn_executivesponsor_value,description,sn_noncompetitive,sn_winlossnodecisionreason,sn_winlossnodecisionnotes";
  const expand = "&$expand=parentaccountid($select=accountid,name)";

  let res: Response;
  try {
    res = await dynamicsFetch(`/opportunities(${id})?$select=${baseFields}${enrichFields}${expand}`, {}, progress);
  } catch {
    progress("⚠️ Some enrichment fields not available — fetching core fields");
    res = await dynamicsFetch(`/opportunities(${id})?$select=${baseFields}${expand}`, {}, progress);
  }
  return mapOpportunity(await res.json() as Record<string, unknown>);
}

export async function searchProducts(name: string, progress: ProgressFn = () => {}): Promise<Product[]> {
  progress(`🔍 Searching product families for "${name}"...`);
  const safe = sanitizeODataSearch(name);
  const path = `/sn_productfamilies?$select=sn_productfamilyid,sn_name&$filter=contains(sn_name,'${safe}')&$top=20`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const results = (data.value ?? []).map((r: Record<string, unknown>) => ({
    productid: r.sn_productfamilyid as string,
    name: r.sn_name as string,
  }));
  progress(`✅ Found ${results.length} matching products`);
  return results;
}

// ---------------------------------------------------------------------------
// Engagements
// ---------------------------------------------------------------------------

export async function getProductById(id: string, progress: ProgressFn = () => {}): Promise<Product> {
  progress(`🔍 Looking up product ${id}...`);
  const res = await dynamicsFetch(`/sn_productfamilies(${id})?$select=sn_productfamilyid,sn_name`, {}, progress);
  const r = await res.json() as Record<string, unknown>;
  return { productid: r.sn_productfamilyid as string, name: r.sn_name as string };
}

export async function fetchEngagementsByOpportunity(opportunityId: string, progress: ProgressFn = () => {}): Promise<Engagement[]> {
  requireGuid(opportunityId, "opportunityId");
  progress(`📡 Fetching engagements for opportunity ${opportunityId}...`);
  const path =
    `/sn_engagements` +
    `?$filter=_sn_opportunityid_value eq ${opportunityId}` +
    `&$select=sn_engagementid,sn_engagementnumber,sn_name,sn_description,sn_completeddate,sn_categorycode,sn_salesstagecode,statecode,statuscode,_sn_engagementtypeid_value,_sn_opportunityid_value,_sn_accountid_value,_sn_primaryproductid_value,_ownerid_value,createdon,modifiedon` +
    `&$expand=sn_engagementtypeid($select=sn_name),sn_primaryproductid($select=sn_name)` +
    `&$orderby=modifiedon desc`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const results = (data.value ?? []).map(mapEngagement);
  progress(`✅ Found ${results.length} engagements`);
  return results;
}

function mapEngagement(r: Record<string, unknown>): Engagement {
  const typeEntity    = r["sn_engagementtypeid"]    as { sn_name?: string } | null;
  const accountEntity = r["sn_accountid"]           as { name?: string } | null;
  const oppEntity     = r["sn_opportunityid"]       as { name?: string } | null;
  const productEntity = r["sn_primaryproductid"]    as { sn_name?: string } | null;
  return {
    ...r,
    engagementTypeName:  typeEntity?.sn_name ?? undefined,
    // ownerid is a polymorphic principal — use formatted value annotation
    ownerName:           r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    accountName:         accountEntity?.name ?? undefined,
    opportunityName:     oppEntity?.name ?? undefined,
    primaryProductName:  productEntity?.sn_name ?? undefined,
    salesStageName:      r["sn_salesstagecode@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    categoryName:        r["sn_categorycode@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    statusName:          r["statuscode@OData.Community.Display.V1.FormattedValue"] as string | undefined,
  } as Engagement;
}

export async function fetchEngagementById(id: string, progress: ProgressFn = () => {}): Promise<Engagement> {
  const path =
    `/sn_engagements(${id})` +
    `?$select=sn_engagementid,sn_engagementnumber,sn_name,sn_description,sn_completeddate,sn_categorycode,sn_salesstagecode,statecode,statuscode,_sn_engagementtypeid_value,_sn_opportunityid_value,_sn_accountid_value,_sn_primaryproductid_value,_ownerid_value,createdon,modifiedon` +
    `&$expand=sn_engagementtypeid($select=sn_name),sn_accountid($select=name),sn_opportunityid($select=name),sn_primaryproductid($select=sn_name)`;

  const res = await dynamicsFetch(path, {}, progress);
  return mapEngagement(await res.json() as Record<string, unknown>);
}

export interface CreateEngagementInput {
  opportunityId: string;
  accountId: string;
  primaryProductId: string;
  name: string;
  type: EngagementType;
  notes?: string;
  completedDate?: string;
}

export async function createEngagement(input: CreateEngagementInput, progress: ProgressFn = () => {}): Promise<Engagement> {
  auditLog("create_engagement", { type: input.type, opportunityId: input.opportunityId });
  const typeGuid = ENGAGEMENT_TYPE_GUIDS[input.type];
  if (!typeGuid) throw new Error(`Unknown engagement type: ${input.type}`);

  // Pre-check: Dynamics enforces at most 1 engagement per type per opportunity
  progress(`🔍 Checking for existing ${input.type} on this opportunity...`);
  const existing = await fetchEngagementsByOpportunity(input.opportunityId, progress);
  const duplicate = existing.find(e => e.engagementTypeName === input.type);
  if (duplicate) {
    const isCancelled = duplicate.statusName?.toLowerCase().includes("cancel");
    if (isCancelled) {
      throw new Error(
        `❌ A cancelled ${input.type} already exists on this opportunity: ${duplicate.sn_name} (${duplicate.sn_engagementnumber ?? duplicate.sn_engagementid}). ` +
        `Reopen the existing one instead of creating a duplicate.`
      );
    }
    throw new Error(
      `❌ A ${input.type} engagement already exists on this opportunity: ${duplicate.sn_name} (${duplicate.sn_engagementnumber ?? duplicate.sn_engagementid}). ` +
      `Collaborate on the existing one instead of creating a duplicate.`
    );
  }

  // Auto-complete if a completed date is provided and it's today or in the past
  const isCompleted = !!input.completedDate && new Date(input.completedDate) <= new Date();

  const payload: Record<string, unknown> = {
    sn_name: input.name,
    sn_description: input.notes,
    sn_completeddate: input.completedDate,
    "sn_engagementtypeid@odata.bind": `/sn_engagementtypes(${typeGuid})`,
    "sn_opportunityid@odata.bind": `/opportunities(${input.opportunityId})`,
    "sn_accountid@odata.bind": `/accounts(${input.accountId})`,
    "sn_primaryproductid@odata.bind": `/sn_productfamilies(${input.primaryProductId})`,
    ...(isCompleted ? { statecode: 1, statuscode: 2 } : {}),
  };

  progress(`📝 Creating "${input.name}" (${input.type}) engagement in Dynamics...`);
  const res = await dynamicsFetch("/sn_engagements", {
    method: "POST",
    body: JSON.stringify(payload),
    headers: { Prefer: "return=representation" },
  }, progress);

  let engagement: Engagement;
  if (res.status === 201) {
    engagement = mapEngagement(await res.json() as Record<string, unknown>);
  } else {
    const location = res.headers.get("OData-EntityId") ?? res.headers.get("Location");
    const match = location?.match(/sn_engagements\(([^)]+)\)/);
    if (!match?.[1]) throw new Error("Engagement created but could not retrieve ID from response");
    engagement = await fetchEngagementById(match[1], progress);
  }

  progress("✅ Engagement created successfully");

  // Auto-generate a timeline note on creation
  const engId = engagement.sn_engagementid;
  if (engId) {
    const today = new Date().toISOString().slice(0, 10);
    await createTimelineNote(
      engId,
      `${input.type} created – ${today}`,
      input.notes
        ? `Engagement created.\n\n${input.notes}`
        : `Engagement created.`,
      progress
    );
  }

  return engagement;
}

// ---------------------------------------------------------------------------
// Structured description builder (used for Technical Win and all engagements)
// ---------------------------------------------------------------------------

// Per-type label for the key_points bullet list
const KEY_POINTS_LABEL: Record<EngagementType, string> = {
  "Technical Win":            "Milestones achieved",
  "Discovery":                "Objectives / Key questions uncovered",
  "Demo":                     "Demo delivered / Customer feedback",
  "Business Case":            "Value drivers / Quantified benefits",
  "RFx":                      "Key requirements addressed",
  "Workshop":                 "Topics covered / Outcomes",
  "POV":                      "POV scope / Results",
  "EBC":                      "Topics discussed / Outcomes",
  "Customer Business Review": "Topics reviewed / Key outcomes",
  "Post Sale Engagement":     "Topics covered / Outcomes",
};

export interface EngagementDescription {
  engagementType?: EngagementType;
  useCase?: string;
  keyPoints?: string[];       // bullet list — label varies by type
  secondaryPoints?: string[]; // type-specific secondary list (e.g. "Customer feedback" for Demo, "Key questions" for Discovery)
  nextActions?: string[];     // bullet list
  risks?: string;
  stakeholders?: string;
  submissionDate?: string;    // RFx-specific
}

// Per-type label for the secondary bullet list (if applicable)
const SECONDARY_POINTS_LABEL: Partial<Record<EngagementType, string>> = {
  "Discovery":                "Key questions / requirements uncovered",
  "Demo":                     "Customer reactions / feedback",
  "Business Case":            "Quantified benefits",
  "Workshop":                 "Customer reactions / feedback",
  "POV":                      "Customer reactions / results",
  "EBC":                      "Customer reactions / feedback",
  "Customer Business Review": "Action items agreed",
};

/** Strip ALL leading bullet characters (•, -, *) so the tool can add its own consistently.
 *  Handles double-bullets (e.g. "• • text") and repeated markers. */
export function stripBullet(s: string): string {
  return s.replace(/^(?:[\s]*[•\-*][\s]*)+(.*)/s, "$1");
}

export function buildDescription(d: EngagementDescription): string {
  const lines: string[] = [];
  if (d.useCase) lines.push(`Use Case: ${d.useCase}`);
  if (d.keyPoints?.length) {
    const label = d.engagementType ? (KEY_POINTS_LABEL[d.engagementType] ?? "Key points") : "Key points";
    lines.push(`${label}:`);
    d.keyPoints.forEach(p => lines.push(`• ${stripBullet(p)}`));
  }
  if (d.secondaryPoints?.length && d.engagementType) {
    const label = SECONDARY_POINTS_LABEL[d.engagementType] ?? "Additional notes";
    lines.push(`${label}:`);
    d.secondaryPoints.forEach(p => lines.push(`• ${stripBullet(p)}`));
  }
  if (d.submissionDate && d.engagementType === "RFx") {
    lines.push(`Submission date: ${d.submissionDate}`);
  }
  if (d.nextActions?.length) {
    lines.push("Next actions:");
    d.nextActions.forEach(a => lines.push(`• ${stripBullet(a)}`));
  }
  lines.push(`Risks/Help Required: ${d.risks ?? "-"}`);
  if (d.stakeholders) lines.push(`Stakeholders: ${d.stakeholders}`);
  return lines.join("\n");
}

// ---------------------------------------------------------------------------
// Timeline notes (Dynamics annotations)
// ---------------------------------------------------------------------------

export async function createTimelineNote(
  engagementId: string,
  title: string,
  text: string,
  progress: ProgressFn = () => {}
): Promise<void> {
  // Dedup guard: skip if a note with the same title was created in the last 60s
  try {
    const recentNotes = await listTimelineNotes(engagementId, () => {});
    const now = Date.now();
    const isDuplicate = recentNotes.some(n =>
      n.subject === title && n.createdon && (now - new Date(n.createdon).getTime()) < 60_000
    );
    if (isDuplicate) {
      progress(`⏭️ Timeline note "${title}" already exists (dedup) — skipping`);
      return;
    }
  } catch (e) { process.stderr.write(`[alfred:warn] timeline dedup check failed: ${e instanceof Error ? e.message : String(e)}\n`); }

  progress(`📋 Adding timeline note: "${title}"...`);
  await dynamicsFetch("/annotations", {
    method: "POST",
    body: JSON.stringify({
      subject: title,
      notetext: text,
      "objectid_sn_engagement@odata.bind": `/sn_engagements(${engagementId})`,
    }),
  }, progress);
  progress("✅ Timeline note added");
}

// ---------------------------------------------------------------------------
// Update engagement
// ---------------------------------------------------------------------------

export async function updateEngagement(
  id: string,
  patch: {
    name?: string;
    type?: EngagementType;
    primaryProductId?: string;            // sn_productfamilies GUID
    completedDate?: string;
    markComplete?: boolean;               // sets statecode=1, statuscode=2
    description?: EngagementDescription; // structured — replaces full description
    notes?: string;                       // plain text fallback
    timelineTitle?: string;               // if set, creates a timeline note
    timelineText?: string;
  },
  progress: ProgressFn = () => {}
): Promise<Engagement> {
  progress(`📝 Updating engagement ${id}...`);

  const payload: Record<string, unknown> = {};
  if (patch.name) payload.sn_name = patch.name;
  if (patch.completedDate) payload.sn_completeddate = patch.completedDate;
  if (patch.markComplete === true)  { payload.statecode = 1; payload.statuscode = 2; }
  if (patch.markComplete === false) { payload.statecode = 0; payload.statuscode = 1; }
  if (patch.description) payload.sn_description = buildDescription(patch.description);
  else if (patch.notes) payload.sn_description = patch.notes;
  if (patch.type) {
    const typeGuid = ENGAGEMENT_TYPE_GUIDS[patch.type];
    if (!typeGuid) throw new Error(`Unknown engagement type: ${patch.type}`);
    payload["sn_engagementtypeid@odata.bind"] = `/sn_engagementtypes(${typeGuid})`;
  }
  if (patch.primaryProductId) {
    payload["sn_primaryproductid@odata.bind"] = `/sn_productfamilies(${requireGuid(patch.primaryProductId, "primaryProductId")})`;
  }

  await dynamicsFetch(`/sn_engagements(${id})`, {
    method: "PATCH",
    body: JSON.stringify(payload),
    headers: { "If-Match": "*" },
  }, progress);

  // Auto-create timeline notes on meaningful state changes
  const today = new Date().toISOString().slice(0, 10);
  if (patch.markComplete === true && !patch.timelineTitle) {
    await createTimelineNote(id, `Completed – ${today}`, "Engagement marked as complete.", progress);
  } else if (patch.markComplete === false && !patch.timelineTitle) {
    await createTimelineNote(id, `Reopened – ${today}`, "Engagement reopened.", progress);
  }

  if (patch.timelineTitle && patch.timelineText) {
    // Update-in-place: if a note with the same title exists, delete it first
    try {
      const existing = await listTimelineNotes(id, () => {});
      const match = existing.find(n => n.subject === patch.timelineTitle);
      if (match) {
        progress(`♻️ Replacing existing timeline note "${patch.timelineTitle}"...`);
        await deleteTimelineNote(match.annotationid, progress);
      }
    } catch (e) { process.stderr.write(`[alfred:warn] timeline note lookup for update-in-place failed: ${e instanceof Error ? e.message : String(e)}\n`); }
    await createTimelineNote(id, patch.timelineTitle, patch.timelineText, progress);
  }

  progress("✅ Engagement updated — fetching latest record...");
  return await fetchEngagementById(id, progress);
}

// ---------------------------------------------------------------------------
// Timeline note management
// ---------------------------------------------------------------------------

export interface TimelineNote {
  annotationid: string;
  subject: string;
  notetext?: string;
  createdon?: string;
}

export async function listTimelineNotes(
  entityId: string,
  progress: ProgressFn = () => {}
): Promise<TimelineNote[]> {
  progress(`📋 Fetching timeline notes for ${entityId}...`);
  const path =
    `/annotations?$filter=_objectid_value eq ${entityId}` +
    `&$select=annotationid,subject,notetext,createdon` +
    `&$orderby=createdon desc`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const notes = (data.value ?? []) as TimelineNote[];
  progress(`✅ Found ${notes.length} timeline note(s)`);
  return notes;
}

// ---------------------------------------------------------------------------
// Activities (appointments, phone calls, tasks)
// ---------------------------------------------------------------------------

export interface Activity {
  activityid: string;
  activitytype: string;   // "appointment", "phonecall", "task", etc.
  subject: string;
  description?: string;
  scheduledstart?: string;
  scheduledend?: string;
  actualstart?: string;
  actualend?: string;
  statecode: number;      // 0=Open, 1=Completed, 2=Canceled
  statusName: string;
  ownerName?: string;
  createdOn: string;
  regardingName?: string;
  regardingId?: string;
}

function mapActivity(r: Record<string, unknown>): Activity {
  const statecode = r.statecode as number;
  return {
    activityid: r.activityid as string ?? "",
    activitytype: r.activitytypecode as string ?? "",
    subject: r.subject as string ?? "",
    description: r.description as string,
    scheduledstart: r.scheduledstart as string,
    scheduledend: r.scheduledend as string,
    actualstart: r.actualstart as string,
    actualend: r.actualend as string,
    statecode,
    statusName: statecode === 0 ? "Open" : statecode === 1 ? "Completed" : "Canceled",
    ownerName: r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string,
    createdOn: r.createdon as string ?? "",
    regardingName: r["_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"] as string,
    regardingId: r._regardingobjectid_value as string,
  };
}

/** List all activities (appointments, calls, tasks, etc.) on an opportunity. */
export async function listActivities(
  opportunityId: string,
  progress: ProgressFn = () => {},
  options: { includeCompleted?: boolean; top?: number; activityType?: string } = {}
): Promise<Activity[]> {
  requireGuid(opportunityId, "opportunityId");
  progress(`📋 Fetching activities for opportunity ${opportunityId}...`);

  let filter = `_regardingobjectid_value eq ${opportunityId}`;
  if (!options.includeCompleted) {
    filter += ` and statecode eq 0`;
  }
  if (options.activityType) {
    const safe = sanitizeODataSearch(options.activityType);
    filter += ` and activitytypecode eq '${safe}'`;
  }

  const path =
    `/activitypointers` +
    `?$select=activityid,activitytypecode,subject,description,scheduledstart,scheduledend,actualstart,actualend,statecode,createdon,_ownerid_value,_regardingobjectid_value` +
    `&$filter=${encodeURIComponent(filter)}` +
    `&$orderby=scheduledstart desc` +
    `&$top=${options.top ?? 50}`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const activities = (data.value ?? []).map(mapActivity);
  progress(`✅ Found ${activities.length} activity/activities`);
  return activities;
}

export interface CreateAppointmentInput {
  opportunityId: string;
  subject: string;          // e.g. "#NBM Discovery call with Roche"
  startTime: string;        // ISO datetime
  endTime?: string;         // ISO datetime (defaults to startTime + 1h)
  description?: string;
  location?: string;
  requiredAttendees?: string[];  // email addresses
  optionalAttendees?: string[];
}

/** Create an appointment linked to an opportunity. */
export async function createAppointment(
  input: CreateAppointmentInput,
  progress: ProgressFn = () => {}
): Promise<Activity> {
  requireGuid(input.opportunityId, "opportunityId");
  auditLog("create_appointment", { opportunityId: input.opportunityId, subject: input.subject });
  progress(`📅 Creating appointment: ${input.subject}...`);

  // Calculate end time (default: 1 hour after start)
  const start = new Date(input.startTime);
  const endTime = input.endTime ?? new Date(start.getTime() + 60 * 60 * 1000).toISOString();

  const body: Record<string, unknown> = {
    subject: input.subject,
    scheduledstart: input.startTime,
    scheduledend: endTime,
    "regardingobjectid_opportunity@odata.bind": `/opportunities(${input.opportunityId})`,
  };
  if (input.description) body.description = input.description;
  if (input.location) body.location = input.location;

  // Add attendees as activity parties
  const parties: Array<{ "partyid_systemuser@odata.bind"?: string; addressused?: string; participationtypemask: number }> = [];
  for (const email of input.requiredAttendees ?? []) {
    parties.push({ addressused: email, participationtypemask: 5 }); // 5 = Required
  }
  for (const email of input.optionalAttendees ?? []) {
    parties.push({ addressused: email, participationtypemask: 6 }); // 6 = Optional
  }
  if (parties.length > 0) body.appointment_activity_parties = parties;

  const res = await dynamicsFetch("/appointments", {
    method: "POST",
    headers: { "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(body),
  }, progress);

  const created = await res.json();
  progress(`✅ Appointment created: ${input.subject}`);
  return mapActivity(created as Record<string, unknown>);
}

export interface CreatePhoneCallInput {
  opportunityId: string;
  subject: string;
  description?: string;
  phoneNumber?: string;
  directionCode?: boolean;  // true = Outgoing, false = Incoming
}

/** Create a phone call activity linked to an opportunity. */
export async function createPhoneCall(
  input: CreatePhoneCallInput,
  progress: ProgressFn = () => {}
): Promise<Activity> {
  requireGuid(input.opportunityId, "opportunityId");
  auditLog("create_phonecall", { opportunityId: input.opportunityId, subject: input.subject });
  progress(`📞 Logging phone call: ${input.subject}...`);

  const body: Record<string, unknown> = {
    subject: input.subject,
    phonenumber: input.phoneNumber ?? "",
    directioncode: input.directionCode ?? true, // default outgoing
    "regardingobjectid_opportunity@odata.bind": `/opportunities(${input.opportunityId})`,
  };
  if (input.description) body.description = input.description;

  const res = await dynamicsFetch("/phonecalls", {
    method: "POST",
    headers: { "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(body),
  }, progress);

  const created = await res.json();
  progress(`✅ Phone call logged: ${input.subject}`);
  return mapActivity(created as Record<string, unknown>);
}

export interface CreateTaskInput {
  opportunityId: string;
  subject: string;
  description?: string;
  dueDate?: string;         // ISO date
}

/** Create a follow-up task linked to an opportunity. */
export async function createTask(
  input: CreateTaskInput,
  progress: ProgressFn = () => {}
): Promise<Activity> {
  requireGuid(input.opportunityId, "opportunityId");
  auditLog("create_task", { opportunityId: input.opportunityId, subject: input.subject });
  progress(`📝 Creating task: ${input.subject}...`);

  const body: Record<string, unknown> = {
    subject: input.subject,
    "regardingobjectid_opportunity@odata.bind": `/opportunities(${input.opportunityId})`,
  };
  if (input.description) body.description = input.description;
  if (input.dueDate) body.scheduledend = input.dueDate;

  const res = await dynamicsFetch("/tasks", {
    method: "POST",
    headers: { "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(body),
  }, progress);

  const created = await res.json();
  progress(`✅ Task created: ${input.subject}`);
  return mapActivity(created as Record<string, unknown>);
}

/** Mark an activity (appointment, phone call, task) as complete. */
export async function completeActivity(
  activityType: string,
  activityId: string,
  progress: ProgressFn = () => {}
): Promise<void> {
  requireGuid(activityId, "activityId");
  const safe = sanitizeODataSearch(activityType);
  auditLog("complete_activity", { activityType: safe, activityId });
  progress(`✅ Marking ${safe} as complete...`);

  // Map activity type to correct OData endpoint (plural form)
  const entityMap: Record<string, string> = {
    appointment: "appointments",
    phonecall: "phonecalls",
    task: "tasks",
  };
  const entity = entityMap[safe.toLowerCase()] ?? `${safe.toLowerCase()}s`;

  await dynamicsFetch(`/${entity}(${activityId})`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ statecode: 1, statuscode: 5 }), // 1=Completed, 5=Completed status
  }, progress);

  progress(`✅ ${safe} marked as complete`);
}

// ---------------------------------------------------------------------------
// Collaboration Notes
// ---------------------------------------------------------------------------

/** Note type option set values for sn_collaborationnote.sn_notetype */
const COLLAB_NOTE_TYPE_MAP: Record<string, number> = {
  "general notes": 100000000,
  "sales ops update": 100000001,
  "renewal update": 100000002,
  "next steps": 100000003,
  "prime notes": 100000004,
};

const COLLAB_NOTE_TYPE_NAMES: Record<number, string> = Object.fromEntries(
  Object.entries(COLLAB_NOTE_TYPE_MAP).map(([k, v]) => [v, k.split(" ").map(w => w[0].toUpperCase() + w.slice(1)).join(" ")])
);

export type CollabNoteType = "General Notes" | "Sales Ops Update" | "Renewal Update" | "Next Steps" | "Prime Notes";

export interface CollaborationNote {
  id: string;
  noteType: string;
  noteTypeCode?: number;
  notes: string;
  owner: string;
  createdOn: string;
  modifiedOn?: string;
  opportunityId?: string;
  opportunityName?: string;
}

function mapCollabNote(r: Record<string, unknown>): CollaborationNote {
  const noteTypeCode = r.sn_notetype as number | undefined;
  return {
    id: r.sn_collaborationnoteid as string ?? r.activityid as string ?? "",
    noteType: (r["sn_notetype@OData.Community.Display.V1.FormattedValue"] as string)
              ?? COLLAB_NOTE_TYPE_NAMES[noteTypeCode ?? -1]
              ?? String(noteTypeCode ?? "Unknown"),
    noteTypeCode,
    notes: r.sn_notes as string ?? r.description as string ?? "",
    owner: r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string ?? "",
    createdOn: r.createdon as string ?? "",
    modifiedOn: r.modifiedon as string,
    opportunityId: r._regardingobjectid_value as string,
    opportunityName: r["_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"] as string,
  };
}

// Auto-discover the correct entity set name for Collaboration Notes (varies by Dynamics instance)
const COLLAB_NOTE_ENTITY_CANDIDATES = [
  "sn_collaborationnotes",        // plural (standard ServiceNow pattern)
  "sn_collaborationnote",         // singular
  "sn_salescollaborationnotes",   // alternate naming
  "msdyn_collaborationnotes",     // Microsoft Dynamics prefix
];
let resolvedCollabNoteEntity: string | null = null;

async function getCollabNoteEntity(progress: ProgressFn): Promise<string> {
  if (resolvedCollabNoteEntity) return resolvedCollabNoteEntity;

  // Try each candidate until one doesn't 404
  for (const entity of COLLAB_NOTE_ENTITY_CANDIDATES) {
    try {
      const res = await dynamicsFetch(`/${entity}?$top=1`, {}, () => {});
      if (res.ok) {
        resolvedCollabNoteEntity = entity;
        progress(`✅ Discovered collaboration notes entity: ${entity}`);
        return entity;
      }
    } catch { /* try next */ }
  }

  // Last resort: try metadata discovery
  try {
    progress("🔍 Searching Dynamics metadata for Collaboration Notes entity...");
    const metaRes = await dynamicsFetch(
      `/EntityDefinitions?$filter=contains(DisplayName/UserLocalizedLabel/Label,'Collaboration Note')&$select=LogicalName,EntitySetName&$top=3`,
      {}, () => {}
    );
    if (metaRes.ok) {
      const metaData = await metaRes.json() as { value: { LogicalName: string; EntitySetName: string }[] };
      if (metaData.value?.[0]?.EntitySetName) {
        resolvedCollabNoteEntity = metaData.value[0].EntitySetName;
        progress(`✅ Found via metadata: ${resolvedCollabNoteEntity}`);
        return resolvedCollabNoteEntity;
      }
    }
  } catch { /* metadata discovery failed */ }

  throw new Error(
    "Could not find the Collaboration Notes entity in Dynamics. " +
    "Please open a Collaboration Note in Dynamics 365, check the URL for the entity name, " +
    "and let Fred know so he can update Alfred."
  );
}

export async function listCollaborationNotes(
  opportunityId: string,
  progress: ProgressFn = () => {}
): Promise<CollaborationNote[]> {
  requireGuid(opportunityId, "opportunityId");
  progress(`📝 Fetching collaboration notes for opportunity ${opportunityId}...`);

  const entity = await getCollabNoteEntity(progress);
  const path =
    `/${entity}` +
    `?$select=sn_collaborationnoteid,activityid,sn_notetype,sn_notes,description,createdon,modifiedon,_ownerid_value,_regardingobjectid_value` +
    `&$filter=_regardingobjectid_value eq ${opportunityId}` +
    `&$orderby=createdon desc` +
    `&$top=50`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const notes = (data.value ?? []).map(mapCollabNote);
  progress(`✅ Found ${notes.length} collaboration note(s)`);
  return notes;
}

export interface CreateCollabNoteInput {
  opportunityId: string;
  noteType: string;      // "General Notes", "Next Steps", etc.
  notes: string;         // The note content
}

export async function createCollaborationNote(
  input: CreateCollabNoteInput,
  progress: ProgressFn = () => {}
): Promise<CollaborationNote> {
  requireGuid(input.opportunityId, "opportunityId");

  const noteTypeKey = input.noteType.toLowerCase();
  const noteTypeCode = COLLAB_NOTE_TYPE_MAP[noteTypeKey];
  if (noteTypeCode === undefined) {
    throw new Error(`Invalid note type "${input.noteType}". Valid types: ${Object.values(COLLAB_NOTE_TYPE_NAMES).join(", ")}`);
  }

  auditLog("create_collaboration_note", { opportunityId: input.opportunityId, noteType: input.noteType });
  progress(`📝 Creating ${input.noteType} collaboration note...`);

  const body: Record<string, unknown> = {
    sn_notetype: noteTypeCode,
    sn_notes: input.notes,
    "regardingobjectid_opportunity@odata.bind": `/opportunities(${input.opportunityId})`,
  };

  const entity = await getCollabNoteEntity(progress);
  const res = await dynamicsFetch(`/${entity}`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(body),
  }, progress);

  const created = await res.json();
  const note = mapCollabNote(created as Record<string, unknown>);
  progress(`✅ Created collaboration note: ${note.noteType}`);
  return note;
}

// ---------------------------------------------------------------------------
// Opportunity write operations (Sales MCP)
// ---------------------------------------------------------------------------

export interface CreateOpportunityInput {
  name: string;
  accountId: string;
  closeDate: string;               // ISO date e.g. "2026-12-31"
  opportunityType?: number;        // 1=New Business 2=Renewal 3=Existing Customer
  forecastCategory?: number;       // 100000001=Pipeline 100000002=Best Case 100000003=Committed
  ownerId?: string;                // Sales Rep systemuser GUID
  scId?: string;                   // SC systemuser GUID
  notes?: string;
}

export interface UpdateOpportunityInput {
  opportunityId: string;
  name?: string;
  closeDate?: string;
  forecastCategory?: number;
  ownerId?: string;
  scId?: string;
  notes?: string;
  winLossNotes?: string;        // sn_winlossnodecisionnotes — why the deal was won/lost
  probability?: number;         // closeprobability — e.g. 80
}

export async function createOpportunity(
  input: CreateOpportunityInput,
  progress: ProgressFn = () => {}
): Promise<Opportunity> {
  auditLog("create_opportunity", { name: input.name, accountId: input.accountId });
  progress(`📝 Creating opportunity "${input.name}"...`);

  const payload: Record<string, unknown> = {
    name: input.name,
    estimatedclosedate: input.closeDate,
    "parentaccountid@odata.bind": `/accounts(${input.accountId})`,
    ...(input.opportunityType !== undefined ? { opportunitytypecode: input.opportunityType } : {}),
    ...(input.forecastCategory !== undefined ? { msdyn_forecastcategory: input.forecastCategory } : {}),
    ...(input.ownerId ? { "ownerid@odata.bind": `/systemusers(${input.ownerId})` } : {}),
    ...(input.scId  ? { "sn_solutionconsultant@odata.bind": `/systemusers(${input.scId})` } : {}),
    ...(input.notes ? { description: input.notes } : {}),
  };

  const res = await dynamicsFetch("/opportunities", {
    method: "POST",
    body: JSON.stringify(payload),
    headers: { Prefer: "return=representation" },
  }, progress);

  let opp: Opportunity;
  if (res.status === 201) {
    opp = mapOpportunity(await res.json() as Record<string, unknown>);
  } else {
    const location = res.headers.get("OData-EntityId") ?? res.headers.get("Location");
    const match = location?.match(/opportunities\(([^)]+)\)/);
    if (!match?.[1]) throw new Error("Opportunity created but could not retrieve ID from response");
    opp = await fetchOpportunityById(match[1], progress);
  }

  progress(`✅ Opportunity created: ${opp.name} (${opp.sn_number ?? opp.opportunityid})`);
  return opp;
}

export async function updateOpportunity(
  input: UpdateOpportunityInput,
  progress: ProgressFn = () => {}
): Promise<Opportunity> {
  auditLog("update_opportunity", { opportunityId: input.opportunityId });
  progress(`📝 Updating opportunity ${input.opportunityId}...`);

  const payload: Record<string, unknown> = {
    ...(input.name        ? { name: input.name } : {}),
    ...(input.closeDate   ? { estimatedclosedate: input.closeDate } : {}),
    ...(input.forecastCategory !== undefined ? { msdyn_forecastcategory: input.forecastCategory } : {}),
    ...(input.ownerId     ? { "ownerid@odata.bind": `/systemusers(${input.ownerId})` } : {}),
    ...(input.scId        ? { "sn_solutionconsultant@odata.bind": `/systemusers(${input.scId})` } : {}),
    ...(input.notes       ? { description: input.notes } : {}),
    ...(input.winLossNotes ? { sn_winlossnodecisionnotes: input.winLossNotes } : {}),
    ...(input.probability !== undefined ? { closeprobability: input.probability } : {}),
  };

  await dynamicsFetch(`/opportunities(${input.opportunityId})`, {
    method: "PATCH",
    body: JSON.stringify(payload),
  }, progress);

  const updated = await fetchOpportunityById(input.opportunityId, progress);
  progress(`✅ Opportunity updated`);
  return updated;
}

export interface SystemUser {
  systemuserid: string;
  fullname: string;
  internalemailaddress?: string;
  title?: string;
}

export async function searchSystemUsers(
  name: string,
  progress: ProgressFn = () => {}
): Promise<SystemUser[]> {
  const safe = sanitizeODataSearch(name);
  progress(`👤 Searching users for "${safe}"...`);
  const res = await dynamicsFetch(
    `/systemusers?$select=systemuserid,fullname,internalemailaddress,title` +
    `&$filter=contains(fullname,'${safe}') and isdisabled eq false` +
    `&$top=10`,
    {}, progress
  );
  const data = await res.json() as { value: Record<string, unknown>[] };
  return (data.value ?? []).map(r => ({
    systemuserid: r.systemuserid as string,
    fullname:     r.fullname as string,
    internalemailaddress: r.internalemailaddress as string | undefined,
    title:        r.title as string | undefined,
  }));
}

export async function deleteEngagement(
  engagementId: string,
  progress: ProgressFn = () => {}
): Promise<void> {
  requireGuid(engagementId, "engagementId");
  auditLog("delete_engagement", { engagementId });
  progress(`🗑️ Deleting engagement ${engagementId}...`);
  await dynamicsFetch(`/sn_engagements(${engagementId})`, { method: "DELETE" }, progress);
  progress("✅ Engagement deleted");
}

export async function deleteTimelineNote(
  annotationId: string,
  progress: ProgressFn = () => {}
): Promise<void> {
  requireGuid(annotationId, "annotationId");
  auditLog("delete_timeline_note", { annotationId });
  progress(`🗑️ Deleting timeline note ${annotationId}...`);
  await dynamicsFetch(`/annotations(${annotationId})`, { method: "DELETE" }, progress);
  progress("✅ Timeline note deleted");
}

// ---------------------------------------------------------------------------
// Accounts
// ---------------------------------------------------------------------------

export interface Account {
  accountid: string;
  name: string;
  industryName?: string;
  websiteurl?: string;
  telephone1?: string;
  numberofemployees?: number;
  revenue?: number;
  address?: string;
  ownerName?: string;
  scName?: string;
  description?: string;
}

function mapAccount(r: Record<string, unknown>): Account {
  const city    = r.address1_city as string | undefined;
  const country = r.address1_country as string | undefined;
  const address = [city, country].filter(Boolean).join(", ") || undefined;
  return {
    accountid:         r.accountid as string,
    name:              r.name as string,
    industryName:      r["industrycode@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    websiteurl:        r.websiteurl as string | undefined,
    telephone1:        r.telephone1 as string | undefined,
    numberofemployees: r.numberofemployees as number | undefined,
    revenue:           r.revenue as number | undefined,
    address,
    ownerName:         r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    scName:            r["_sn_solutionconsultant_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    description:       r.description as string | undefined,
  };
}

const ACCOUNT_SELECT =
  "accountid,name,industrycode,websiteurl,telephone1,numberofemployees,revenue,address1_city,address1_country,description,_ownerid_value,_sn_solutionconsultant_value";

export async function fetchAccountById(id: string, progress: ProgressFn = () => {}): Promise<Account> {
  progress(`🏢 Fetching account ${id}...`);
  const res = await dynamicsFetch(
    `/accounts(${id})?$select=${ACCOUNT_SELECT}`,
    {}, progress
  );
  const account = mapAccount(await res.json() as Record<string, unknown>);
  progress(`✅ Account: ${account.name}`);
  return account;
}

export async function searchAccounts(name: string, progress: ProgressFn = () => {}): Promise<Account[]> {
  progress(`🔍 Searching accounts for "${name}"...`);
  const safe = sanitizeODataSearch(name);
  const path =
    `/accounts?$select=${ACCOUNT_SELECT}` +
    `&$filter=contains(name,'${safe}')` +
    `&$orderby=name asc&$top=10`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const results = (data.value ?? []).map(mapAccount);
  progress(`✅ Found ${results.length} account(s)`);
  return results;
}

// ---------------------------------------------------------------------------
// Contacts (CRM contacts — for external stakeholder enrichment)
// ---------------------------------------------------------------------------

export interface Contact {
  contactid: string;
  fullname: string;
  emailaddress1?: string;
  jobtitle?: string;
  telephone1?: string;
  accountName?: string;
}

export async function searchContacts(
  query: string,
  opts: { accountId?: string } = {},
  progress: ProgressFn = () => {}
): Promise<Contact[]> {
  const safe = sanitizeODataSearch(query);
  progress(`🔍 Searching contacts for "${safe}"...`);
  let filter = `(contains(fullname,'${safe}') or contains(emailaddress1,'${safe}'))`;
  if (opts.accountId) {
    requireGuid(opts.accountId, "accountId");
    filter += ` and _parentcustomerid_value eq '${opts.accountId}'`;
  }
  const path =
    `/contacts?$select=contactid,fullname,emailaddress1,jobtitle,telephone1,_parentcustomerid_value` +
    `&$expand=parentcustomerid_account($select=name)` +
    `&$filter=${encodeURIComponent(filter)}` +
    `&$orderby=fullname asc&$top=20`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json() as { value: Record<string, unknown>[] };
  const results = (data.value ?? []).map(r => ({
    contactid:     r.contactid as string,
    fullname:      r.fullname as string,
    emailaddress1: r.emailaddress1 as string | undefined,
    jobtitle:      r.jobtitle as string | undefined,
    telephone1:    r.telephone1 as string | undefined,
    accountName:   (r.parentcustomerid_account as Record<string, unknown> | undefined)?.name as string | undefined,
  }));
  progress(`✅ Found ${results.length} contact(s)`);
  return results;
}

export interface CreateContactInput {
  firstName: string;
  lastName: string;
  email?: string;
  jobTitle?: string;
  phone?: string;
  accountId?: string;     // Link to parent account
}

/** Create a new contact in Dynamics 365. */
export async function createContact(
  input: CreateContactInput,
  progress: ProgressFn = () => {}
): Promise<Contact> {
  auditLog("create_contact", { name: `${input.firstName} ${input.lastName}`, email: input.email });
  progress(`👤 Creating contact: ${input.firstName} ${input.lastName}...`);

  const body: Record<string, unknown> = {
    firstname: input.firstName,
    lastname: input.lastName,
  };
  if (input.email) body.emailaddress1 = input.email;
  if (input.jobTitle) body.jobtitle = input.jobTitle;
  if (input.phone) body.telephone1 = input.phone;
  if (input.accountId) {
    requireGuid(input.accountId, "accountId");
    body["parentcustomerid_account@odata.bind"] = `/accounts(${input.accountId})`;
  }

  const res = await dynamicsFetch("/contacts", {
    method: "POST",
    headers: { "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify(body),
  }, progress);

  const created = await res.json() as Record<string, unknown>;
  progress(`✅ Created contact: ${input.firstName} ${input.lastName}`);
  return {
    contactid: created.contactid as string,
    fullname: `${input.firstName} ${input.lastName}`,
    emailaddress1: input.email,
    jobtitle: input.jobTitle,
    telephone1: input.phone,
  };
}

/** Stakeholder role codes for opportunity connections. */
const STAKEHOLDER_ROLE_MAP: Record<string, number> = {
  "champion": 876130000,
  "economic buyer": 876130001,
  "technical buyer": 876130002,
  "coach": 876130003,
  "decision maker": 876130004,
  "influencer": 876130005,
  "end user": 876130006,
  "executive sponsor": 876130007,
};

export interface OpportunityContact {
  connectionid: string;
  contactId: string;
  contactName: string;
  role?: string;
  roleCode?: number;
}

/** List contacts associated with an opportunity via connections. */
export async function listOpportunityContacts(
  opportunityId: string,
  progress: ProgressFn = () => {}
): Promise<OpportunityContact[]> {
  requireGuid(opportunityId, "opportunityId");
  progress(`👥 Fetching contacts for opportunity ${opportunityId}...`);

  const path =
    `/connections` +
    `?$select=connectionid,_record1id_value,_record1roleid_value` +
    `&$filter=_record2id_value eq ${opportunityId} and _record1objecttypecode_value eq 2` +  // 2 = contact
    `&$top=50`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json() as { value: Record<string, unknown>[] };
  const contacts = (data.value ?? []).map(r => ({
    connectionid: r.connectionid as string,
    contactId: r._record1id_value as string ?? "",
    contactName: r["_record1id_value@OData.Community.Display.V1.FormattedValue"] as string ?? "",
    role: r["_record1roleid_value@OData.Community.Display.V1.FormattedValue"] as string,
    roleCode: r._record1roleid_value as number | undefined,
  }));
  progress(`✅ Found ${contacts.length} contact(s) on opportunity`);
  return contacts;
}

/** Add a contact to an opportunity via a connection. */
export async function addContactToOpportunity(
  contactId: string,
  opportunityId: string,
  role?: string,
  progress: ProgressFn = () => {}
): Promise<void> {
  requireGuid(contactId, "contactId");
  requireGuid(opportunityId, "opportunityId");
  auditLog("add_contact_to_opportunity", { contactId, opportunityId, role });
  progress(`🔗 Linking contact to opportunity...`);

  const body: Record<string, unknown> = {
    "record1id_contact@odata.bind": `/contacts(${contactId})`,
    "record2id_opportunity@odata.bind": `/opportunities(${opportunityId})`,
  };

  // Try to find a matching connection role if specified
  if (role) {
    const roleLower = role.toLowerCase();
    // Search for the role in Dynamics connection roles
    const safe = sanitizeODataSearch(role);
    const roleRes = await dynamicsFetch(
      `/connectionroles?$select=connectionroleid,name&$filter=contains(name,'${safe}')&$top=1`,
      {}, progress
    );
    const roleData = await roleRes.json() as { value: { connectionroleid: string; name: string }[] };
    if (roleData.value?.[0]) {
      body["record1roleid@odata.bind"] = `/connectionroles(${roleData.value[0].connectionroleid})`;
      progress(`📋 Role: ${roleData.value[0].name}`);
    } else {
      progress(`⚠️ Role "${role}" not found in Dynamics — linking without role`);
    }
  }

  await dynamicsFetch("/connections", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }, progress);

  progress(`✅ Contact linked to opportunity`);
}

// ---------------------------------------------------------------------------
// Collaboration team (opportunity-level) & Engagement participants
// ---------------------------------------------------------------------------

const COLLAB_ROLE_NAMES: Record<number, string> = {
  876130005: "Solution Consultant",
  876130023: "Renewal Account Manager",
};

const COLLAB_ACCESS_NAMES: Record<number, string> = {
  876130001: "Edit",
  876130002: "Read",
};

const COLLAB_PRIMARY_YES = 876130000;

export interface CollaborationTeamMember {
  id: string;
  userName: string;
  userId: string;
  collaborationRole: string;
  jobRole?: string;
  isPrimary: boolean;
  accessLevel: string;
}

function mapCollabMember(r: Record<string, unknown>): CollaborationTeamMember {
  return {
    id:                r.sn_opportunitycollaborationteamid as string,
    userName:          r["_sn_user_value@OData.Community.Display.V1.FormattedValue"] as string ?? "—",
    userId:            r._sn_user_value as string,
    collaborationRole: r["sn_collaborationrole@OData.Community.Display.V1.FormattedValue"] as string
                       ?? COLLAB_ROLE_NAMES[r.sn_collaborationrole as number] ?? String(r.sn_collaborationrole),
    jobRole:           r["_sn_jobroleid_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    isPrimary:         (r.sn_isprimary as number) === COLLAB_PRIMARY_YES,
    accessLevel:       r["sn_opportunityheader@OData.Community.Display.V1.FormattedValue"] as string
                       ?? COLLAB_ACCESS_NAMES[r.sn_opportunityheader as number] ?? "—",
  };
}

export async function fetchCollaborationTeam(
  opportunityId: string,
  progress: ProgressFn = () => {}
): Promise<CollaborationTeamMember[]> {
  progress(`👥 Fetching collaboration team for opportunity ${opportunityId}...`);
  const path =
    `/sn_opportunitycollaborationteams` +
    `?$select=sn_opportunitycollaborationteamid,_sn_user_value,sn_collaborationrole,sn_isprimary,sn_opportunityheader,_sn_jobroleid_value` +
    `&$filter=_sn_opportunity_value eq ${opportunityId} and statecode eq 0` +
    `&$orderby=sn_name asc` +
    `&$top=100`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const members = (data.value ?? []).map(mapCollabMember);
  progress(`✅ Found ${members.length} collaboration team member(s)`);
  return members;
}

export async function fetchMyCollaborationOpportunities(
  progress: ProgressFn = () => {}
): Promise<Opportunity[]> {
  const userId = await fetchCurrentUserId(progress);
  requireGuid(userId, "currentUserId");
  progress(`📡 Finding opportunities where you are on the collaboration team...`);

  // Step 1: Get all collab team entries for this user
  const collabPath =
    `/sn_opportunitycollaborationteams` +
    `?$select=_sn_opportunity_value` +
    `&$filter=_sn_user_value eq ${userId} and statecode eq 0` +
    `&$top=200`;

  const collabRes = await dynamicsFetch(collabPath, {}, progress);
  const collabData = await collabRes.json();
  const oppIds = [...new Set(
    (collabData.value ?? []).map((r: Record<string, unknown>) => r._sn_opportunity_value as string).filter(Boolean)
  )] as string[];

  if (oppIds.length === 0) {
    progress("ℹ️ You are not on any opportunity collaboration teams");
    return [];
  }

  // Step 2: Fetch full opportunity details for those IDs (batch in groups of 15 for OData filter length)
  const batchPromises: Promise<Opportunity[]>[] = [];
  for (let i = 0; i < oppIds.length; i += 15) {
    const batch = oppIds.slice(i, i + 15);
    const idFilter = batch.map(id => `opportunityid eq ${id}`).join(" or ");
    const selectFields = "opportunityid,sn_number,name,_accountid_value,_ownerid_value,_sn_solutionconsultant_value,_sn_territory_value,statuscode,estimatedclosedate,totalamount,sn_netnewacv,msdyn_forecastcategory,stepname,closeprobability,sn_opportunitytype,sn_opportunitybusinessunitlist";
    const path =
      `/opportunities?$select=${selectFields}` +
      `&$expand=parentaccountid($select=accountid,name)` +
      `&$filter=statecode eq 0 and (${idFilter})` +
      `&$orderby=estimatedclosedate asc&$top=50`;

    batchPromises.push(
      dynamicsFetch(path, {}, progress)
        .then(res => res.json())
        .then(data => (data.value ?? []).map(mapOpportunity))
    );
  }
  const allOpps: Opportunity[] = (await Promise.all(batchPromises)).flat();

  progress(`✅ Found ${allOpps.length} open opportunities where you are a collaborator`);
  return allOpps;
}

export interface EngagementParticipant {
  id: string;
  userName: string;
  userId: string;
  title?: string;
  isPrimary: boolean;
}

function mapParticipant(r: Record<string, unknown>): EngagementParticipant {
  return {
    id:        r.sn_engagementassigneeid as string,
    userName:  r["_sn_assigneeid_value@OData.Community.Display.V1.FormattedValue"] as string ?? r.sn_name as string ?? "—",
    userId:    r._sn_assigneeid_value as string,
    title:     r["a_title"] as string | undefined,
    isPrimary: r.sn_primary === true,
  };
}

export async function fetchEngagementParticipants(
  engagementId: string,
  progress: ProgressFn = () => {}
): Promise<EngagementParticipant[]> {
  progress(`👥 Fetching participants for engagement ${engagementId}...`);
  const path =
    `/sn_engagementassignees` +
    `?$select=sn_engagementassigneeid,_sn_assigneeid_value,sn_primary,sn_name` +
    `&$filter=_sn_engagementid_value eq ${engagementId} and statecode eq 0` +
    `&$top=50`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const participants = (data.value ?? []).map(mapParticipant);
  progress(`✅ Found ${participants.length} participant(s)`);
  return participants;
}

export interface MyEngagementFilter {
  search?: string;
  engagementType?: string;
  status?: "open" | "complete" | "all";
  top?: number;
}

export async function fetchMyEngagementAssignments(
  filter: MyEngagementFilter = {},
  progress: ProgressFn = () => {}
): Promise<Engagement[]> {
  const userId = await fetchCurrentUserId(progress);
  requireGuid(userId, "currentUserId");
  const top = filter.top ?? 50;
  auditLog("fetch_my_engagement_assignments", { userId, filter });
  progress(`📡 Finding engagements where you are a participant...`);

  // Step 1: Get engagement IDs from assignee table
  const assigneePath =
    `/sn_engagementassignees` +
    `?$select=_sn_engagementid_value` +
    `&$filter=_sn_assigneeid_value eq ${userId} and statecode eq 0` +
    `&$top=200`;

  const assigneeRes = await dynamicsFetch(assigneePath, {}, progress);
  const assigneeData = await assigneeRes.json();
  const engIds = [...new Set(
    (assigneeData.value ?? []).map((r: Record<string, unknown>) => r._sn_engagementid_value as string).filter(Boolean)
  )] as string[];

  if (engIds.length === 0) {
    progress("ℹ️ You are not a participant on any engagements");
    return [];
  }

  // Step 2: Fetch full engagement details (batch in groups of 15)
  const allEngagements: Engagement[] = [];
  for (let i = 0; i < engIds.length; i += 15) {
    const batch = engIds.slice(i, i + 15);
    const idFilter = batch.map(id => `sn_engagementid eq ${id}`).join(" or ");

    let statusFilter = "";
    const normStatus = filter.status?.toLowerCase();
    if (normStatus === "open") statusFilter = " and statecode eq 0";
    else if (normStatus === "complete") statusFilter = " and statecode eq 1";
    else if (normStatus && normStatus !== "all") {
      progress(`⚠️ Unknown status filter "${filter.status}" — expected open|complete|all, returning all`);
    }

    let searchFilter = "";
    if (filter.search) {
      const safe = sanitizeODataSearch(filter.search);
      searchFilter = ` and (contains(sn_name,'${safe}'))`;
    }

    const path =
      `/sn_engagements` +
      `?$select=sn_engagementid,sn_engagementnumber,sn_name,sn_description,sn_completeddate,sn_categorycode,sn_salesstagecode,statecode,statuscode,_sn_engagementtypeid_value,_sn_opportunityid_value,_sn_accountid_value,_sn_primaryproductid_value,_ownerid_value,createdon,modifiedon` +
      `&$expand=sn_engagementtypeid($select=sn_name),sn_accountid($select=name),sn_opportunityid($select=name),sn_primaryproductid($select=sn_name)` +
      `&$filter=(${idFilter})${statusFilter}${searchFilter}` +
      `&$orderby=modifiedon desc` +
      `&$top=${top}`;

    const res = await dynamicsFetch(path, {}, progress);
    const data = await res.json();
    allEngagements.push(...(data.value ?? []).map(mapEngagement));
  }

  // Step 3: Optional type filter (post-query since type is a lookup, not a simple field)
  let results = allEngagements;
  if (filter.engagementType) {
    const typeLower = filter.engagementType.toLowerCase();
    results = results.filter(e => e.engagementTypeName?.toLowerCase().includes(typeLower));
  }

  progress(`✅ Found ${results.length} engagements where you are a participant`);
  return results;
}

// ---------------------------------------------------------------------------
// Engagement attendees — Active Participants (internal) + Engagement Contacts (external)
// ---------------------------------------------------------------------------

// imported from shared.ts — SN_INTERNAL_DOMAINS

export interface AttendeeResult {
  email: string;
  name: string;
  type: "participant" | "contact" | "not_found";
  id?: string;
}

export async function addAttendeesToEngagement(
  engagementId: string,
  attendees: { name: string; email: string }[],
  progress: ProgressFn = () => {}
): Promise<AttendeeResult[]> {
  const results: AttendeeResult[] = [];
  const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  // Dedup by email (case-insensitive) — keep first occurrence
  const seen = new Set<string>();
  const dedupedAttendees = attendees.filter(a => {
    const key = a.email.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  if (dedupedAttendees.length < attendees.length) {
    progress(`ℹ️ Deduped ${attendees.length} → ${dedupedAttendees.length} unique attendees`);
  }

  auditLog("add_attendees", { engagementId, count: dedupedAttendees.length });

  // Pre-filter invalid emails synchronously
  const validAttendees: typeof dedupedAttendees = [];
  for (const attendee of dedupedAttendees) {
    if (!EMAIL_RE.test(attendee.email)) {
      progress(`⚠️ Skipping invalid email: ${attendee.email}`);
      results.push({ email: attendee.email, name: attendee.name, type: "not_found" });
    } else {
      validAttendees.push(attendee);
    }
  }

  // Process all valid attendees in parallel — each is independent (lookup + create)
  const attendeeResults = await Promise.allSettled(
    validAttendees.map(async (attendee) => {
      const domain = attendee.email.split("@")[1]!.toLowerCase();
      const isInternal = SN_INTERNAL_DOMAINS.has(domain);

      if (isInternal) {
        // Look up Dynamics systemuser by internal email
        const safeEmail = sanitizeODataSearch(attendee.email);
        const userRes = await dynamicsFetch(
          `/systemusers?$filter=internalemailaddress eq '${safeEmail}'&$select=systemuserid,fullname&$top=1`,
          {}, progress
        );
        const userData = await userRes.json() as { value: { systemuserid: string; fullname: string }[] };
        const user = userData.value?.[0];
        if (user) {
          await dynamicsFetch("/sn_engagementassignees", {
            method: "POST",
            body: JSON.stringify({
              sn_name: user.fullname,
              "sn_assigneeid@odata.bind":   `/systemusers(${user.systemuserid})`,
              "sn_engagementid@odata.bind": `/sn_engagements(${engagementId})`,
            }),
          }, progress);
          progress(`👤 Added participant: ${user.fullname}`);
          return { email: attendee.email, name: user.fullname, type: "participant" as const, id: user.systemuserid };
        } else {
          progress(`⚠️ Internal user not found in Dynamics: ${attendee.email}`);
          return { email: attendee.email, name: attendee.name, type: "not_found" as const };
        }
      } else {
        // Look up Dynamics contact by email
        const safeEmail = sanitizeODataSearch(attendee.email);
        const contactRes = await dynamicsFetch(
          `/contacts?$filter=emailaddress1 eq '${safeEmail}'&$select=contactid,fullname&$top=1`,
          {}, progress
        );
        const contactData = await contactRes.json() as { value: { contactid: string; fullname: string }[] };
        const contact = contactData.value?.[0];
        if (contact) {
          await dynamicsFetch("/sn_engagementcontacts", {
            method: "POST",
            body: JSON.stringify({
              sn_name: contact.fullname,
              "sn_contactid@odata.bind":    `/contacts(${contact.contactid})`,
              "sn_engagementid@odata.bind": `/sn_engagements(${engagementId})`,
            }),
          }, progress);
          progress(`👤 Added engagement contact: ${contact.fullname}`);
          return { email: attendee.email, name: contact.fullname, type: "contact" as const, id: contact.contactid };
        } else {
          progress(`⚠️ Contact not found in CRM: ${attendee.email}`);
          return { email: attendee.email, name: attendee.name, type: "not_found" as const };
        }
      }
    })
  );

  // Collect results — failed promises become not_found
  for (let i = 0; i < attendeeResults.length; i++) {
    const settled = attendeeResults[i]!;
    if (settled.status === "fulfilled") {
      results.push(settled.value);
    } else {
      results.push({ email: validAttendees[i]!.email, name: validAttendees[i]!.name, type: "not_found" });
    }
  }

  return results;
}

// ---------------------------------------------------------------------------
// Closing Plan — entity auto-discovery + CRUD
// ---------------------------------------------------------------------------

const CLOSING_PLAN_ENTITY_CANDIDATES = [
  "sn_closingplans",             // plural (standard SN pattern)
  "sn_closingplan",              // singular
  "sn_closingplanmilestones",    // might be milestone entity
  "sn_closingplanmilestone",     // singular milestone
  "sn_closingplanactivities",    // alternate naming
  "sn_closingplanactivity",
];
let resolvedClosingPlanEntity: string | null = null;

async function getClosingPlanEntity(progress: ProgressFn): Promise<string> {
  if (resolvedClosingPlanEntity) return resolvedClosingPlanEntity;

  for (const entity of CLOSING_PLAN_ENTITY_CANDIDATES) {
    try {
      const res = await dynamicsFetch(`/${entity}?$top=1`, {}, () => {});
      if (res.ok) {
        resolvedClosingPlanEntity = entity;
        progress(`✅ Discovered closing plan entity: ${entity}`);
        return entity;
      }
    } catch { /* try next */ }
  }

  // Metadata discovery fallback
  try {
    progress("🔍 Searching Dynamics metadata for Closing Plan entity...");
    const metaRes = await dynamicsFetch(
      `/EntityDefinitions?$filter=contains(DisplayName/UserLocalizedLabel/Label,'Closing Plan')&$select=LogicalName,EntitySetName&$top=5`,
      {}, () => {}
    );
    if (metaRes.ok) {
      const metaData = await metaRes.json() as { value: { LogicalName: string; EntitySetName: string }[] };
      if (metaData.value?.[0]?.EntitySetName) {
        resolvedClosingPlanEntity = metaData.value[0].EntitySetName;
        progress(`✅ Found via metadata: ${resolvedClosingPlanEntity}`);
        return resolvedClosingPlanEntity;
      }
    }
  } catch { /* metadata discovery failed */ }

  throw new Error(
    "Could not find the Closing Plan entity in Dynamics. " +
    "Please open a Closing Plan in Dynamics 365, check the URL for the entity name, " +
    "and let Fred know so he can update Alfred."
  );
}

export interface ClosingPlanMilestone {
  id: string;
  title: string;
  dueDate?: string;
  status: string;         // e.g. "Open", "Complete", "At Risk"
  statusCode: number;
  stateCode: number;
  owner?: string;
  description?: string;
  createdOn?: string;
  modifiedOn?: string;
  opportunityId?: string;
  opportunityName?: string;
}

function mapClosingPlanMilestone(r: Record<string, unknown>): ClosingPlanMilestone {
  // Adapt to whichever fields the entity uses — try common patterns
  const id = (r.sn_closingplanid ?? r.sn_closingplanmilestoneid ?? r.activityid ?? r[Object.keys(r).find(k => k.endsWith("id") && !k.startsWith("_")) ?? ""] ?? "") as string;
  return {
    id,
    title: (r.sn_name ?? r.sn_title ?? r.subject ?? r.sn_milestonename ?? "") as string,
    dueDate: (r.sn_duedate ?? r.scheduledend ?? r.sn_targetdate ?? "") as string || undefined,
    status: (r["statecode@OData.Community.Display.V1.FormattedValue"] ?? r["statuscode@OData.Community.Display.V1.FormattedValue"] ?? (r.statecode === 0 ? "Open" : r.statecode === 1 ? "Complete" : "Unknown")) as string,
    statusCode: (r.statuscode ?? 0) as number,
    stateCode: (r.statecode ?? 0) as number,
    owner: (r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] ?? "") as string || undefined,
    description: (r.sn_description ?? r.description ?? "") as string || undefined,
    createdOn: (r.createdon ?? "") as string || undefined,
    modifiedOn: (r.modifiedon ?? "") as string || undefined,
    opportunityId: (r._regardingobjectid_value ?? r._sn_opportunityid_value ?? "") as string || undefined,
    opportunityName: (r["_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"] ?? r["_sn_opportunityid_value@OData.Community.Display.V1.FormattedValue"] ?? "") as string || undefined,
  };
}

export async function listClosingPlan(
  opportunityId: string,
  progress: ProgressFn = () => {},
  options: { includeCompleted?: boolean } = {}
): Promise<ClosingPlanMilestone[]> {
  requireGuid(opportunityId, "opportunityId");
  progress(`📋 Fetching closing plan for opportunity ${opportunityId}...`);

  const entity = await getClosingPlanEntity(progress);

  // Try common lookup patterns — the entity may link to opportunity via different fields
  const lookupFields = [
    `_regardingobjectid_value eq ${opportunityId}`,
    `_sn_opportunityid_value eq ${opportunityId}`,
    `_sn_opportunity_value eq ${opportunityId}`,
  ];

  let milestones: ClosingPlanMilestone[] = [];

  for (const filter of lookupFields) {
    try {
      const stateFilter = options.includeCompleted ? "" : " and statecode eq 0";
      const path =
        `/${entity}` +
        `?$filter=${encodeURIComponent(filter + stateFilter)}` +
        `&$orderby=sn_duedate asc,createdon asc` +
        `&$top=100`;

      const res = await dynamicsFetch(path, {}, progress);
      if (!res.ok) continue;
      const data = await res.json() as { value: Record<string, unknown>[] };
      if (data.value?.length > 0) {
        milestones = data.value.map(mapClosingPlanMilestone);
        break;
      }
    } catch { /* try next lookup field */ }
  }

  progress(`✅ Found ${milestones.length} closing plan milestone(s)`);
  return milestones;
}

export interface CreateMilestoneInput {
  opportunityId: string;
  title: string;
  dueDate?: string;       // ISO date string
  description?: string;
}

export async function createClosingPlanMilestone(
  input: CreateMilestoneInput,
  progress: ProgressFn = () => {}
): Promise<ClosingPlanMilestone> {
  requireGuid(input.opportunityId, "opportunityId");
  auditLog("create_closing_plan_milestone", { opportunityId: input.opportunityId, title: input.title });
  progress(`📌 Creating closing plan milestone: ${input.title}...`);

  const entity = await getClosingPlanEntity(progress);

  // Build body — try multiple binding patterns
  const body: Record<string, unknown> = {
    sn_name: input.title,
  };
  if (input.dueDate) body.sn_duedate = input.dueDate;
  if (input.description) body.sn_description = input.description;

  // Try regarding object binding first (activity pattern), then direct lookup
  const bindingAttempts = [
    { ...body, "regardingobjectid_opportunity@odata.bind": `/opportunities(${input.opportunityId})` },
    { ...body, "sn_opportunityid@odata.bind": `/opportunities(${input.opportunityId})` },
    { ...body, "sn_opportunity@odata.bind": `/opportunities(${input.opportunityId})` },
  ];

  let created: Record<string, unknown> | null = null;
  for (const attempt of bindingAttempts) {
    try {
      const res = await dynamicsFetch(`/${entity}`, {
        method: "POST",
        headers: { "Content-Type": "application/json", Prefer: "return=representation" },
        body: JSON.stringify(attempt),
      }, progress);

      if (res.ok) {
        created = await res.json() as Record<string, unknown>;
        break;
      }
    } catch { /* try next binding */ }
  }

  if (!created) {
    throw new Error(`Failed to create closing plan milestone. The entity "${entity}" may use a different opportunity binding.`);
  }

  progress(`✅ Milestone created: ${input.title}`);
  return mapClosingPlanMilestone(created);
}

export async function updateClosingPlanMilestone(
  milestoneId: string,
  updates: { complete?: boolean; atRisk?: boolean; title?: string; dueDate?: string; description?: string },
  progress: ProgressFn = () => {}
): Promise<void> {
  requireGuid(milestoneId, "milestoneId");
  auditLog("update_closing_plan_milestone", { milestoneId, updates });

  const entity = await getClosingPlanEntity(progress);
  const body: Record<string, unknown> = {};

  if (updates.complete) {
    body.statecode = 1;    // Complete / Inactive
    body.statuscode = 2;   // Completed
    progress(`✅ Marking milestone as complete...`);
  }
  if (updates.atRisk !== undefined) {
    body.sn_atrisk = updates.atRisk;
    progress(updates.atRisk ? `⚠️ Flagging milestone as at risk...` : `✅ Removing risk flag...`);
  }
  if (updates.title) body.sn_name = updates.title;
  if (updates.dueDate) body.sn_duedate = updates.dueDate;
  if (updates.description) body.sn_description = updates.description;

  const res = await dynamicsFetch(`/${entity}(${milestoneId})`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }, progress);

  if (!res.ok && res.status !== 204) {
    throw new Error(`Failed to update milestone ${milestoneId}: ${res.status} ${res.statusText}`);
  }

  progress(`✅ Milestone updated`);
}

// ---------------------------------------------------------------------------
// Forecast Summary — aggregate pipeline by forecast category
// ---------------------------------------------------------------------------

export interface ForecastSummary {
  totalPipeline: number;
  committed: number;
  bestCase: number;
  pipeline: number;
  omitted: number;
  oppCount: number;
  byCategory: Array<{
    category: string;
    categoryCode: number;
    count: number;
    nnacv: number;
    opps: Array<{ name: string; account: string; nnacv: number; closeDate?: string; owner?: string }>;
  }>;
  closingThisQuarter: number;
  closingNextQuarter: number;
  atRiskCount: number;
}

export async function getForecastSummary(
  options: {
    myOppsOnly?: boolean;
    myOppsFilterField?: "owner" | "collab";
    ownerSearch?: string;
    accountSearch?: string;
    quarter?: string;          // e.g. "Q2 2026" — filters by close date within that quarter
  },
  progress: ProgressFn = () => {}
): Promise<ForecastSummary> {
  progress("📊 Building forecast summary...");

  const opps = await fetchOpportunities({
    search: options.accountSearch,
    myOpportunitiesOnly: options.myOppsOnly ?? true,
    myOppsFilterField: options.myOppsFilterField ?? "owner",
    ownerSearch: options.ownerSearch,
    includeClosed: false,
    includeZeroValue: true,
    top: 500,
  }, progress);

  // Quarter filtering
  let filtered = opps;
  if (options.quarter) {
    const qMatch = options.quarter.match(/Q([1-4])\s*(\d{4})/i);
    if (qMatch) {
      const q = parseInt(qMatch[1]!);
      const y = parseInt(qMatch[2]!);
      const qStart = new Date(y, (q - 1) * 3, 1);
      const qEnd = new Date(y, q * 3, 0); // last day of quarter
      filtered = opps.filter(o => {
        if (!o.estimatedclosedate) return false;
        const d = new Date(o.estimatedclosedate);
        return d >= qStart && d <= qEnd;
      });
      progress(`📅 Filtered to ${options.quarter}: ${filtered.length} opps`);
    }
  }

  const today = new Date();
  const soon = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);

  // Current quarter boundaries
  const currentQ = Math.floor(today.getMonth() / 3);
  const currentQEnd = new Date(today.getFullYear(), (currentQ + 1) * 3, 0);
  const nextQEnd = new Date(today.getFullYear(), (currentQ + 2) * 3, 0);

  let committed = 0, bestCase = 0, pipelineVal = 0, omitted = 0;
  let closingThisQ = 0, closingNextQ = 0, atRisk = 0;

  const catMap: Record<string, { code: number; count: number; nnacv: number; opps: ForecastSummary["byCategory"][0]["opps"] }> = {};

  for (const o of filtered) {
    const val = o.nnacv ?? 0;
    const catName = o.forecastCategoryName ?? "Unknown";
    const catCode = o.msdyn_forecastcategory ?? 0;

    switch (catCode) {
      case 100000003: committed += val; break;
      case 100000002: bestCase += val; break;
      case 100000001: pipelineVal += val; break;
      case 100000004: omitted += val; break;
    }

    if (!catMap[catName]) catMap[catName] = { code: catCode, count: 0, nnacv: 0, opps: [] };
    catMap[catName].count++;
    catMap[catName].nnacv += val;
    catMap[catName].opps.push({
      name: o.name,
      account: o.accountName,
      nnacv: val,
      closeDate: o.estimatedclosedate,
      owner: o.ownerName,
    });

    // Timing checks
    if (o.estimatedclosedate) {
      const close = new Date(o.estimatedclosedate);
      if (close <= currentQEnd) closingThisQ++;
      else if (close <= nextQEnd) closingNextQ++;
      if (close < today || (close < soon && close >= today)) atRisk++;
    } else {
      atRisk++; // No close date = at risk
    }
  }

  const byCategory = Object.entries(catMap)
    .sort((a, b) => {
      const order = [100000003, 100000002, 100000001, 100000004, 0];
      return order.indexOf(a[1].code) - order.indexOf(b[1].code);
    })
    .map(([category, data]) => ({
      category,
      categoryCode: data.code,
      count: data.count,
      nnacv: data.nnacv,
      opps: data.opps.sort((a, b) => (b.nnacv ?? 0) - (a.nnacv ?? 0)),
    }));

  progress(`✅ Forecast: ${filtered.length} opps | Committed $${committed.toLocaleString()} | Best Case $${bestCase.toLocaleString()} | Pipeline $${pipelineVal.toLocaleString()}`);

  return {
    totalPipeline: committed + bestCase + pipelineVal,
    committed,
    bestCase,
    pipeline: pipelineVal,
    omitted,
    oppCount: filtered.length,
    byCategory,
    closingThisQuarter: closingThisQ,
    closingNextQuarter: closingNextQ,
    atRiskCount: atRisk,
  };
}

// ---------------------------------------------------------------------------
// Opportunity Summary — read/write the summary tab (annotations or custom entity)
// ---------------------------------------------------------------------------

const OPP_SUMMARY_ENTITY_CANDIDATES = [
  "sn_opportunitysummaries",
  "sn_opportunitysummary",
  "sn_dealreviews",
  "sn_dealreview",
];
let resolvedOppSummaryEntity: string | null = null;

async function getOppSummaryEntity(progress: ProgressFn): Promise<string | null> {
  if (resolvedOppSummaryEntity) return resolvedOppSummaryEntity;

  for (const entity of OPP_SUMMARY_ENTITY_CANDIDATES) {
    try {
      const res = await dynamicsFetch(`/${entity}?$top=1`, {}, () => {});
      if (res.ok) {
        resolvedOppSummaryEntity = entity;
        progress(`✅ Discovered opportunity summary entity: ${entity}`);
        return entity;
      }
    } catch { /* try next */ }
  }

  // Metadata discovery fallback
  try {
    const metaRes = await dynamicsFetch(
      `/EntityDefinitions?$filter=contains(DisplayName/UserLocalizedLabel/Label,'Opportunity Summary') or contains(DisplayName/UserLocalizedLabel/Label,'Deal Review')&$select=LogicalName,EntitySetName&$top=5`,
      {}, () => {}
    );
    if (metaRes.ok) {
      const metaData = await metaRes.json() as { value: { LogicalName: string; EntitySetName: string }[] };
      if (metaData.value?.[0]?.EntitySetName) {
        resolvedOppSummaryEntity = metaData.value[0].EntitySetName;
        progress(`✅ Found via metadata: ${resolvedOppSummaryEntity}`);
        return resolvedOppSummaryEntity;
      }
    }
  } catch { /* metadata failed */ }

  return null; // No custom entity — fall back to annotations
}

export interface OpportunitySummary {
  id: string;
  title: string;
  content: string;
  createdOn?: string;
  modifiedOn?: string;
  owner?: string;
}

export async function getOpportunitySummary(
  opportunityId: string,
  progress: ProgressFn = () => {}
): Promise<OpportunitySummary[]> {
  requireGuid(opportunityId, "opportunityId");
  progress(`📄 Fetching opportunity summary for ${opportunityId}...`);

  // Try custom summary entity first
  const entity = await getOppSummaryEntity(progress);
  if (entity) {
    const lookupFields = [
      `_regardingobjectid_value eq ${opportunityId}`,
      `_sn_opportunityid_value eq ${opportunityId}`,
      `_sn_opportunity_value eq ${opportunityId}`,
    ];
    for (const filter of lookupFields) {
      try {
        const res = await dynamicsFetch(
          `/${entity}?$filter=${encodeURIComponent(filter)}&$orderby=modifiedon desc&$top=20`,
          {}, progress
        );
        if (!res.ok) continue;
        const data = await res.json() as { value: Record<string, unknown>[] };
        if (data.value?.length > 0) {
          const summaries = data.value.map((r): OpportunitySummary => ({
            id: (r[Object.keys(r).find(k => k.endsWith("id") && !k.startsWith("_") && !k.includes("value")) ?? ""] ?? "") as string,
            title: (r.sn_name ?? r.sn_title ?? r.subject ?? "") as string,
            content: (r.sn_summary ?? r.sn_description ?? r.sn_notes ?? r.description ?? r.notetext ?? "") as string,
            createdOn: r.createdon as string | undefined,
            modifiedOn: r.modifiedon as string | undefined,
            owner: r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
          }));
          progress(`✅ Found ${summaries.length} summary record(s)`);
          return summaries;
        }
      } catch { /* try next lookup */ }
    }
  }

  // Fallback: read annotations (standard Dynamics notes on the opportunity)
  progress("📝 No custom summary entity — reading opportunity notes (annotations)...");
  const annotRes = await dynamicsFetch(
    `/annotations?$select=annotationid,subject,notetext,createdon,modifiedon,_ownerid_value` +
    `&$filter=_objectid_value eq ${opportunityId} and isdocument eq false` +
    `&$orderby=modifiedon desc&$top=20`,
    {}, progress
  );
  const annotData = await annotRes.json() as { value: Record<string, unknown>[] };
  const summaries = (annotData.value ?? []).map((r): OpportunitySummary => ({
    id: r.annotationid as string,
    title: (r.subject ?? "") as string,
    content: (r.notetext ?? "") as string,
    createdOn: r.createdon as string | undefined,
    modifiedOn: r.modifiedon as string | undefined,
    owner: r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
  }));
  progress(`✅ Found ${summaries.length} note(s)`);
  return summaries;
}

export async function updateOpportunitySummary(
  opportunityId: string,
  summary: string,
  title?: string,
  progress: ProgressFn = () => {}
): Promise<OpportunitySummary> {
  requireGuid(opportunityId, "opportunityId");
  auditLog("update_opportunity_summary", { opportunityId });
  progress(`📝 Writing opportunity summary...`);

  // Try custom entity first
  const entity = await getOppSummaryEntity(progress);
  if (entity) {
    // Check for existing summary to update in place
    const existing = await getOpportunitySummary(opportunityId, progress);
    if (existing.length > 0 && existing[0]!.id) {
      // Update existing
      const body: Record<string, unknown> = {
        sn_summary: summary,
        sn_description: summary,
        sn_notes: summary,
      };
      if (title) { body.sn_name = title; body.sn_title = title; }
      await dynamicsFetch(`/${entity}(${existing[0]!.id})`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      }, progress);
      progress(`✅ Updated existing summary`);
      return { ...existing[0]!, content: summary, title: title ?? existing[0]!.title };
    }
    // Create new
    const body: Record<string, unknown> = {
      sn_summary: summary,
      sn_description: summary,
      sn_notes: summary,
      sn_name: title ?? "Opportunity Summary",
      "regardingobjectid_opportunity@odata.bind": `/opportunities(${opportunityId})`,
    };
    try {
      const res = await dynamicsFetch(`/${entity}`, {
        method: "POST",
        headers: { "Content-Type": "application/json", Prefer: "return=representation" },
        body: JSON.stringify(body),
      }, progress);
      if (res.ok) {
        progress(`✅ Created summary on custom entity`);
        const created = await res.json() as Record<string, unknown>;
        return {
          id: Object.values(created).find(v => typeof v === "string" && /^[0-9a-f-]{36}$/.test(v)) as string ?? "",
          title: title ?? "Opportunity Summary",
          content: summary,
          createdOn: created.createdon as string | undefined,
          modifiedOn: created.modifiedon as string | undefined,
        };
      }
    } catch { /* fall through to annotation */ }
  }

  // Fallback: create/update as annotation (standard note)
  const existing = await getOpportunitySummary(opportunityId, progress);
  const summaryNote = existing.find(n => n.title.toLowerCase().includes("summary"));

  if (summaryNote?.id) {
    // Update existing note
    await dynamicsFetch(`/annotations(${summaryNote.id})`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        subject: title ?? summaryNote.title ?? "Opportunity Summary",
        notetext: summary,
      }),
    }, progress);
    progress(`✅ Updated existing summary note`);
    return { ...summaryNote, content: summary, title: title ?? summaryNote.title };
  }

  // Create new annotation
  const res = await dynamicsFetch("/annotations", {
    method: "POST",
    headers: { "Content-Type": "application/json", Prefer: "return=representation" },
    body: JSON.stringify({
      subject: title ?? "Opportunity Summary",
      notetext: summary,
      "objectid_opportunity@odata.bind": `/opportunities(${opportunityId})`,
    }),
  }, progress);
  const created = await res.json() as Record<string, unknown>;
  progress(`✅ Created summary note`);
  return {
    id: created.annotationid as string,
    title: title ?? "Opportunity Summary",
    content: summary,
    createdOn: created.createdon as string | undefined,
    modifiedOn: created.modifiedon as string | undefined,
  };
}

// ---------------------------------------------------------------------------
// Quotes — read quotes linked to an opportunity
// ---------------------------------------------------------------------------

export interface Quote {
  quoteid: string;
  name: string;
  quoteNumber?: string;
  status: string;
  statusCode: number;
  totalAmount?: number;
  createdOn?: string;
  modifiedOn?: string;
  effectiveFrom?: string;
  effectiveTo?: string;
  description?: string;
  owner?: string;
}

export async function listQuotes(
  opportunityId: string,
  progress: ProgressFn = () => {}
): Promise<Quote[]> {
  requireGuid(opportunityId, "opportunityId");
  progress(`📋 Fetching quotes for opportunity ${opportunityId}...`);

  const path =
    `/quotes` +
    `?$select=quoteid,name,quotenumber,statuscode,statecode,totalamount,createdon,modifiedon,effectivefrom,effectiveto,description,_ownerid_value` +
    `&$filter=_opportunityid_value eq ${opportunityId}` +
    `&$orderby=modifiedon desc` +
    `&$top=50`;

  let quotes: Quote[] = [];
  try {
    const res = await dynamicsFetch(path, {}, progress);
    if (res.ok) {
      const data = await res.json() as { value: Record<string, unknown>[] };
      quotes = (data.value ?? []).map((r): Quote => ({
        quoteid: r.quoteid as string,
        name: (r.name ?? "") as string,
        quoteNumber: r.quotenumber as string | undefined,
        status: (r["statuscode@OData.Community.Display.V1.FormattedValue"] ?? (r.statecode === 0 ? "Draft" : r.statecode === 1 ? "Active" : r.statecode === 2 ? "Won" : r.statecode === 3 ? "Closed" : "Unknown")) as string,
        statusCode: (r.statuscode ?? 0) as number,
        totalAmount: r.totalamount as number | undefined,
        createdOn: r.createdon as string | undefined,
        modifiedOn: r.modifiedon as string | undefined,
        effectiveFrom: r.effectivefrom as string | undefined,
        effectiveTo: r.effectiveto as string | undefined,
        description: r.description as string | undefined,
        owner: r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
      }));
    }
  } catch (e) {
    // Quotes entity may not exist or be accessible
    progress(`⚠️ Could not fetch quotes: ${e instanceof Error ? e.message : String(e)}`);
    return [];
  }

  progress(`✅ Found ${quotes.length} quote(s)`);
  return quotes;
}
