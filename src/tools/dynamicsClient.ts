import { getAuthCookies, clearAuthCache, type ProgressFn } from "../auth/tokenExtractor.js";
import { userInfo } from "os";
import { DYNAMICS_HOST, ENGAGEMENT_TYPE_GUIDS, type EngagementType } from "../config.js";
import { FORECAST_NAMES, requireGuid } from "../shared.js";

const DYNAMICS_BASE = `${DYNAMICS_HOST}/api/data/v9.2`;

// ---------------------------------------------------------------------------
// Security helpers
// ---------------------------------------------------------------------------

/**
 * Sanitize a user-supplied search string for safe use inside OData contains().
 * Strips characters that could break out of the string context (parentheses,
 * slashes, OData operators) and escapes single quotes.
 */
function sanitizeODataSearch(input: string): string {
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
  } catch { /* non-fatal */ }
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
  totalamount?: number;
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
    // Session expired — clear cache and retry once with fresh cookies
    clearAuthCache();
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
      try { const b = await retry.json(); if (b?.error?.message) msg += ` — ${b.error.message}`; } catch { /* ignore */ }
      throw new Error(msg);
    }
    return retry;
  }

  if (!response.ok) {
    let msg = `Dynamics API error: ${response.status} ${response.statusText}`;
    try {
      const ct = response.headers.get("content-type") ?? "";
      if (ct.includes("json")) {
        const body = await response.json();
        if (body?.error?.message) msg += ` — ${body.error.message}`;
      } else {
        const text = await response.text().catch(() => "");
        if (text) msg += ` — ${text.slice(0, 200)}`;
      }
    } catch { /* ignore */ }
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
    // ownerid is a polymorphic principal — use formatted value annotation (free with odata.include-annotations=*)
    ownerName:           r["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    scName:              r["_sn_solutionconsultant_value@OData.Community.Display.V1.FormattedValue"] as string | undefined,
    totalamount:         r.totalamount as number | undefined,
  };
}

export interface OpportunityFilter {
  top?: number;        // max results (default 50)
  search?: string;     // filter by account/opportunity name (contains)
  minNnacv?: number;   // minimum totalamount (NNACV) in USD
  myOpportunitiesOnly?: boolean; // filter to current user's owned opportunities
  includeClosed?: boolean; // include won/lost/closed opps — default false (open only)
  ownerSearch?: string; // filter by owner (AE) name — resolves to user IDs
}

interface CurrentUser {
  userId: string;
  territoryId?: string;
}

export async function fetchCurrentUserId(progress: ProgressFn = () => {}): Promise<string> {
  const user = await fetchCurrentUser(progress);
  return user.userId;
}

async function fetchCurrentUser(progress: ProgressFn = () => {}): Promise<CurrentUser> {
  progress("👤 Resolving current user...");
  const whoAmI = await dynamicsFetch("/WhoAmI", {}, progress);
  const { UserId } = await whoAmI.json() as { UserId: string };

  // Fetch user's territory GUID
  let territoryId: string | undefined;
  try {
    const userRes = await dynamicsFetch(
      `/systemusers(${UserId})?$select=_sn_fieldterritory_value`,
      {}, progress
    );
    const userData = await userRes.json() as Record<string, unknown>;
    territoryId = userData._sn_fieldterritory_value as string | undefined || undefined;
  } catch {
    // Territory field may not exist — non-fatal
  }

  progress(`👤 User: ${UserId}${territoryId ? ` | Territory: ${territoryId}` : ""}`);
  return { userId: UserId, territoryId };
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
  if (filter.minNnacv) {
    filterClause += ` and totalamount ge ${filter.minNnacv}`;
  }
  if (filter.myOpportunitiesOnly) {
    const { userId, territoryId } = await fetchCurrentUser(progress);
    requireGuid(userId, "currentUserId");
    // Match opps where user is SC OR in the user's territory
    const scFilter = `_sn_solutionconsultant_value eq '${userId}'`;
    const terrFilter = territoryId ? (requireGuid(territoryId, "territoryId"), ` or _sn_fieldterritory_value eq '${territoryId}'`) : "";
    filterClause += ` and (${scFilter}${terrFilter})`;
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

  const path =
    "/opportunities" +
    `?$select=opportunityid,sn_number,name,_accountid_value,_ownerid_value,_sn_solutionconsultant_value,statuscode,estimatedclosedate,totalamount,msdyn_forecastcategory` +
    `&$expand=parentaccountid($select=accountid,name)` +
    `&$filter=${encodeURIComponent(filterClause)}` +
    `&$orderby=estimatedclosedate asc` +
    `&$top=${top}`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const results = (data.value ?? []).map(mapOpportunity);
  progress(`✅ Found ${results.length} opportunities`);
  return results;
}

export async function fetchOpportunityById(id: string, progress: ProgressFn = () => {}): Promise<Opportunity> {
  progress(`📡 Fetching opportunity ${id}...`);
  const path =
    `/opportunities(${id})` +
    "?$select=opportunityid,sn_number,name,_accountid_value,_ownerid_value,_sn_solutionconsultant_value,statuscode,estimatedclosedate,totalamount,msdyn_forecastcategory" +
    "&$expand=parentaccountid($select=accountid,name)";

  const res = await dynamicsFetch(path, {}, progress);
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
  keyPoints?: string[];    // bullet list — label varies by type
  nextActions?: string[];  // bullet list
  risks?: string;
  stakeholders?: string;
}

export function buildDescription(d: EngagementDescription): string {
  const lines: string[] = [];
  if (d.useCase) lines.push(`Use Case: ${d.useCase}`);
  if (d.keyPoints?.length) {
    const label = d.engagementType ? (KEY_POINTS_LABEL[d.engagementType] ?? "Key points") : "Key points";
    lines.push(`${label}:`);
    d.keyPoints.forEach(p => lines.push(`• ${p}`));
  }
  if (d.nextActions?.length) {
    lines.push("Next actions:");
    d.nextActions.forEach(a => lines.push(`• ${a}`));
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
  } catch { /* non-fatal — proceed with creation */ }

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
  engagementId: string,
  progress: ProgressFn = () => {}
): Promise<TimelineNote[]> {
  progress(`📋 Fetching timeline notes for engagement ${engagementId}...`);
  const path =
    `/annotations?$filter=_objectid_value eq ${engagementId}` +
    `&$select=annotationid,subject,notetext,createdon` +
    `&$orderby=createdon desc`;

  const res = await dynamicsFetch(path, {}, progress);
  const data = await res.json();
  const notes = (data.value ?? []) as TimelineNote[];
  progress(`✅ Found ${notes.length} timeline note(s)`);
  return notes;
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
  auditLog("delete_engagement", { engagementId });
  progress(`🗑️ Deleting engagement ${engagementId}...`);
  await dynamicsFetch(`/sn_engagements(${engagementId})`, { method: "DELETE" }, progress);
  progress("✅ Engagement deleted");
}

export async function deleteTimelineNote(
  annotationId: string,
  progress: ProgressFn = () => {}
): Promise<void> {
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
  const allOpps: Opportunity[] = [];
  for (let i = 0; i < oppIds.length; i += 15) {
    const batch = oppIds.slice(i, i + 15);
    const idFilter = batch.map(id => `opportunityid eq ${id}`).join(" or ");
    const path =
      `/opportunities` +
      `?$select=opportunityid,sn_number,name,_accountid_value,_ownerid_value,_sn_solutionconsultant_value,statuscode,estimatedclosedate,totalamount,msdyn_forecastcategory` +
      `&$expand=parentaccountid($select=accountid,name)` +
      `&$filter=statecode eq 0 and (${idFilter})` +
      `&$orderby=estimatedclosedate asc` +
      `&$top=50`;

    const res = await dynamicsFetch(path, {}, progress);
    const data = await res.json();
    allOpps.push(...(data.value ?? []).map(mapOpportunity));
  }

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

const INTERNAL_ATTENDEE_DOMAINS = new Set(["servicenow.com", "now.com"]);

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

  for (const attendee of dedupedAttendees) {
    if (!EMAIL_RE.test(attendee.email)) {
      progress(`⚠️ Skipping invalid email: ${attendee.email}`);
      results.push({ email: attendee.email, name: attendee.name, type: "not_found" });
      continue;
    }
    const domain = attendee.email.split("@")[1]!.toLowerCase();
    const isInternal = INTERNAL_ATTENDEE_DOMAINS.has(domain);

    try {
      if (isInternal) {
        // Look up Dynamics systemuser by internal email
        const safeEmail = attendee.email.replace(/'/g, "''");
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
          results.push({ email: attendee.email, name: user.fullname, type: "participant", id: user.systemuserid });
        } else {
          progress(`⚠️ Internal user not found in Dynamics: ${attendee.email}`);
          results.push({ email: attendee.email, name: attendee.name, type: "not_found" });
        }
      } else {
        // Look up Dynamics contact by email
        const safeEmail = attendee.email.replace(/'/g, "''");
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
          results.push({ email: attendee.email, name: contact.fullname, type: "contact", id: contact.contactid });
        } else {
          progress(`⚠️ Contact not found in CRM: ${attendee.email}`);
          results.push({ email: attendee.email, name: attendee.name, type: "not_found" });
        }
      }
    } catch {
      results.push({ email: attendee.email, name: attendee.name, type: "not_found" });
    }
  }

  return results;
}
