import { getAuthCookies, clearAuthCache, type ProgressFn } from "../auth/tokenExtractor.js";

const DYNAMICS_BASE = "https://servicenow.crm.dynamics.com/api/data/v9.2";

export type EngagementType =
  | "Business Case"
  | "Customer Business Review"
  | "Demo"
  | "Discovery"
  | "EBC"
  | "Post Sale Engagement"
  | "POV"
  | "RFx"
  | "Technical Win"
  | "Workshop";

// Hardcoded GUIDs from sn_engagementtypes lookup table
const ENGAGEMENT_TYPE_GUIDS: Record<EngagementType, string> = {
  "Business Case":            "e7cadf53-6e73-eb11-a812-000d3a1c68be",
  "Customer Business Review": "e8cadf53-6e73-eb11-a812-000d3a1c68be",
  "Demo":                     "e9cadf53-6e73-eb11-a812-000d3a1c68be",
  "Discovery":                "43d14916-aa9c-ec11-b400-0022483026eb",
  "EBC":                      "eacadf53-6e73-eb11-a812-000d3a1c68be",
  "Post Sale Engagement":     "7a12ddba-aaba-eb11-8236-000d3a9d0356",
  "POV":                      "ebcadf53-6e73-eb11-a812-000d3a1c68be",
  "RFx":                      "eccadf53-6e73-eb11-a812-000d3a1c68be",
  "Technical Win":            "edcadf53-6e73-eb11-a812-000d3a1c68be",
  "Workshop":                 "eecadf53-6e73-eb11-a812-000d3a1c68be",
};

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

async function dynamicsFetch(path: string, options: RequestInit = {}, progress: ProgressFn = () => {}): Promise<Response> {
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
  const response = await fetch(url, { ...options, headers });

  if (response.status === 401) {
    // Session expired — clear cache and retry once with fresh cookies
    clearAuthCache();
    progress("🔄 Dynamics session expired — re-acquiring cookies...");
    const freshCookie = await getAuthCookies(progress);
    const retry = await fetch(url, {
      ...options,
      headers: { ...headers, Cookie: freshCookie },
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
      const body = await response.json();
      if (body?.error?.message) msg += ` — ${body.error.message}`;
    } catch { /* ignore */ }
    throw new Error(msg);
  }

  return response;
}

// ---------------------------------------------------------------------------
// Opportunities
// ---------------------------------------------------------------------------

const FORECAST_NAMES: Record<number, string> = {
  100000001: "Pipeline",
  100000002: "Best Case",
  100000003: "Committed",
  100000004: "Omitted",
};

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
  progress(`📡 Querying Dynamics for open opportunities (max ${top})...`);

  let filterClause = filter.includeClosed ? "statecode ge 0" : "statecode eq 0";
  if (filter.search) {
    const safe = filter.search.replace(/'/g, "''");
    filterClause += ` and (contains(name,'${safe}') or contains(sn_number,'${safe}'))`;
  }
  if (filter.minNnacv) {
    filterClause += ` and totalamount ge ${filter.minNnacv}`;
  }
  if (filter.myOpportunitiesOnly) {
    const { userId, territoryId } = await fetchCurrentUser(progress);
    // Match opps where user is SC OR in the user's territory
    const scFilter = `_sn_solutionconsultant_value eq '${userId}'`;
    const terrFilter = territoryId ? ` or _sn_fieldterritory_value eq '${territoryId}'` : "";
    filterClause += ` and (${scFilter}${terrFilter})`;
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
  const safe = name.replace(/'/g, "''");
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
    engagement = match?.[1]
      ? await fetchEngagementById(match[1], progress)
      : (payload as unknown as Engagement);
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

export async function deleteTimelineNote(
  annotationId: string,
  progress: ProgressFn = () => {}
): Promise<void> {
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
const ACCOUNT_EXPAND = "";

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
  const safe = name.replace(/'/g, "''");
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

  for (const attendee of attendees) {
    const domain = attendee.email.split("@")[1]?.toLowerCase() ?? "";
    const isInternal = INTERNAL_ATTENDEE_DOMAINS.has(domain);

    try {
      if (isInternal) {
        // Look up Dynamics systemuser by internal email
        const userRes = await dynamicsFetch(
          `/systemusers?$filter=internalemailaddress eq '${attendee.email}'&$select=systemuserid,fullname&$top=1`,
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
        const contactRes = await dynamicsFetch(
          `/contacts?$filter=emailaddress1 eq '${attendee.email}'&$select=contactid,fullname&$top=1`,
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
