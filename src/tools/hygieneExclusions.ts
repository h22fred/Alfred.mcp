import { existsSync, mkdirSync, readFileSync, writeFileSync } from "fs";
import { homedir } from "os";
import { join } from "path";

const EXCLUSIONS_DIR  = join(homedir(), ".alfred");
const EXCLUSIONS_FILE = join(EXCLUSIONS_DIR, "hygiene-exclusions.json");

// ---------------------------------------------------------------------------
// Schema
// ---------------------------------------------------------------------------

export interface OppExclusion {
  opp: string;       // OPTY#### sn_number
  reason?: string;
}

export interface MilestoneExclusion {
  opp?: string;      // OPTY#### — match against sn_number
  account?: string;  // case-insensitive exact match against accountName
  milestones: string[];
  reason?: string;
}

export interface ParentChildLink {
  child: string;     // OPTY#### — child inherits parent's completed milestones
  parent: string;    // OPTY####
  reason?: string;
}

export interface HygieneExclusions {
  version: 1;
  excludedOpps: OppExclusion[];
  excludedMilestones: MilestoneExclusion[];
  parentChildLinks: ParentChildLink[];
}

const EMPTY: HygieneExclusions = {
  version: 1,
  excludedOpps: [],
  excludedMilestones: [],
  parentChildLinks: [],
};

// ---------------------------------------------------------------------------
// Load / Save
// ---------------------------------------------------------------------------

export function loadExclusions(): HygieneExclusions {
  if (!existsSync(EXCLUSIONS_FILE)) return { ...EMPTY, excludedOpps: [], excludedMilestones: [], parentChildLinks: [] };
  try {
    const raw = JSON.parse(readFileSync(EXCLUSIONS_FILE, "utf8")) as HygieneExclusions;
    return {
      version: 1,
      excludedOpps:       Array.isArray(raw.excludedOpps)       ? raw.excludedOpps       : [],
      excludedMilestones: Array.isArray(raw.excludedMilestones) ? raw.excludedMilestones : [],
      parentChildLinks:   Array.isArray(raw.parentChildLinks)   ? raw.parentChildLinks   : [],
    };
  } catch {
    return { ...EMPTY, excludedOpps: [], excludedMilestones: [], parentChildLinks: [] };
  }
}

function saveExclusions(ex: HygieneExclusions): void {
  mkdirSync(EXCLUSIONS_DIR, { recursive: true });
  writeFileSync(EXCLUSIONS_FILE, JSON.stringify(ex, null, 2), "utf8");
}

// ---------------------------------------------------------------------------
// Mutation helpers
// ---------------------------------------------------------------------------

export function addExcludedOpp(opp: string, reason?: string): void {
  const ex = loadExclusions();
  const normalized = opp.trim().toUpperCase();
  if (!ex.excludedOpps.some(e => e.opp.toUpperCase() === normalized)) {
    ex.excludedOpps.push({ opp: normalized, ...(reason ? { reason } : {}) });
    saveExclusions(ex);
  }
}

export function addExcludedMilestones(
  target: { opp?: string; account?: string },
  milestones: string[],
  reason?: string
): void {
  const ex = loadExclusions();
  const entry: MilestoneExclusion = {
    ...(target.opp     ? { opp:     target.opp.trim().toUpperCase() }   : {}),
    ...(target.account ? { account: target.account.trim() }             : {}),
    milestones,
    ...(reason ? { reason } : {}),
  };
  ex.excludedMilestones.push(entry);
  saveExclusions(ex);
}

export function addParentChildLink(child: string, parent: string, reason?: string): void {
  const ex = loadExclusions();
  const c = child.trim().toUpperCase();
  const p = parent.trim().toUpperCase();
  if (!ex.parentChildLinks.some(l => l.child.toUpperCase() === c)) {
    ex.parentChildLinks.push({ child: c, parent: p, ...(reason ? { reason } : {}) });
    saveExclusions(ex);
  }
}

export function removeExcludedOpp(opp: string): boolean {
  const ex = loadExclusions();
  const normalized = opp.trim().toUpperCase();
  const before = ex.excludedOpps.length;
  ex.excludedOpps = ex.excludedOpps.filter(e => e.opp.toUpperCase() !== normalized);
  if (ex.excludedOpps.length !== before) { saveExclusions(ex); return true; }
  return false;
}

export function removeExcludedMilestones(target: { opp?: string; account?: string }): boolean {
  const ex = loadExclusions();
  const before = ex.excludedMilestones.length;
  ex.excludedMilestones = ex.excludedMilestones.filter(e => {
    if (target.opp     && e.opp?.toUpperCase()    === target.opp.trim().toUpperCase())     return false;
    if (target.account && e.account?.toLowerCase() === target.account.trim().toLowerCase()) return false;
    return true;
  });
  if (ex.excludedMilestones.length !== before) { saveExclusions(ex); return true; }
  return false;
}

export function removeParentChildLink(child: string): boolean {
  const ex = loadExclusions();
  const normalized = child.trim().toUpperCase();
  const before = ex.parentChildLinks.length;
  ex.parentChildLinks = ex.parentChildLinks.filter(l => l.child.toUpperCase() !== normalized);
  if (ex.parentChildLinks.length !== before) { saveExclusions(ex); return true; }
  return false;
}

// ---------------------------------------------------------------------------
// Apply exclusions
// ---------------------------------------------------------------------------

/** Returns true if this opp should be skipped entirely. */
export function isOppExcluded(
  opp: { opportunityid: string; sn_number?: string; accountName?: string },
  ex: HygieneExclusions
): boolean {
  if (!opp.sn_number) return false;
  const num = opp.sn_number.toUpperCase();
  return ex.excludedOpps.some(e => e.opp.toUpperCase() === num);
}

/** Returns the milestones that should be ignored for a given opp (from both opp- and account-level rules). */
export function getExcludedMilestones(
  opp: { sn_number?: string; accountName?: string },
  ex: HygieneExclusions
): Set<string> {
  const excluded = new Set<string>();
  const num = opp.sn_number?.toUpperCase();
  const acct = opp.accountName?.toLowerCase();

  for (const rule of ex.excludedMilestones) {
    const oppMatch  = rule.opp     && num  && rule.opp.toUpperCase()    === num;
    const acctMatch = rule.account && acct && rule.account.toLowerCase() === acct;
    if (oppMatch || acctMatch) {
      for (const m of rule.milestones) excluded.add(m);
    }
  }
  return excluded;
}

/** Returns the parent OPTY number for this child, if a link is configured. */
export function getParentLink(
  opp: { sn_number?: string },
  ex: HygieneExclusions
): string | null {
  if (!opp.sn_number) return null;
  const num = opp.sn_number.toUpperCase();
  return ex.parentChildLinks.find(l => l.child.toUpperCase() === num)?.parent ?? null;
}

// ---------------------------------------------------------------------------
// Summary text (for list display)
// ---------------------------------------------------------------------------

export function formatExclusions(ex: HygieneExclusions): string {
  if (
    ex.excludedOpps.length === 0 &&
    ex.excludedMilestones.length === 0 &&
    ex.parentChildLinks.length === 0
  ) {
    return "No hygiene exclusions configured. Every open opportunity will be evaluated normally.";
  }

  const lines: string[] = ["## Hygiene exclusions", ""];

  if (ex.excludedOpps.length > 0) {
    lines.push("### Fully excluded opportunities");
    for (const e of ex.excludedOpps) {
      lines.push(`- **${e.opp}**${e.reason ? ` — ${e.reason}` : ""}`);
    }
    lines.push("");
  }

  if (ex.excludedMilestones.length > 0) {
    lines.push("### Milestone exclusions");
    for (const e of ex.excludedMilestones) {
      const target = e.opp ? `opp ${e.opp}` : `account "${e.account}"`;
      lines.push(`- **${target}** — skip: ${e.milestones.join(", ")}${e.reason ? ` _(${e.reason})_` : ""}`);
    }
    lines.push("");
  }

  if (ex.parentChildLinks.length > 0) {
    lines.push("### Parent → child links (child inherits parent milestones)");
    for (const l of ex.parentChildLinks) {
      lines.push(`- **${l.child}** inherits from **${l.parent}**${l.reason ? ` — ${l.reason}` : ""}`);
    }
    lines.push("");
  }

  return lines.join("\n");
}
