/**
 * Architecture tests — verify structural invariants, tool parity between SC and Sales,
 * field schema correctness, and configuration consistency.
 */
import { describe, it, expect } from "vitest";
import { readFileSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { ALL_ENGAGEMENT_TYPES, ENGAGEMENT_TYPE_GUIDS } from "../src/config.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, "..");

function readSource(path: string): string {
  return readFileSync(join(ROOT, path), "utf8");
}

function extractToolNames(src: string): string[] {
  const re = /server\.tool\(\s*\n?\s*"([^"]+)"/g;
  const names: string[] = [];
  let m;
  while ((m = re.exec(src)) !== null) names.push(m[1]);
  return names;
}

// ---------------------------------------------------------------------------
// SC / Sales tool parity
// ---------------------------------------------------------------------------
describe("SC / Sales tool parity", () => {
  const scSrc = readSource("src/sc/index.ts");
  const salesSrc = readSource("src/sales/index.ts");
  const scTools = extractToolNames(scSrc);
  const salesTools = extractToolNames(salesSrc);

  // Shared tools that must exist in both servers
  const REQUIRED_SHARED = [
    "get_opportunity",
    "list_engagements",
    "get_engagement",
    "create_engagement",
    "update_engagement",
    "list_timeline_notes",
    "delete_timeline_note",
    "delete_engagement",
    "add_engagement_attendees",
    "get_engagement_participants",
    "search_my_engagements",
    "get_calendar_events",
    "search_emails",
    "list_mail_folders",
    "get_teams_transcript",
    "get_teams_chats",
    "configure_teams_webhook",
    "post_teams_notification",
    "run_hygiene_sweep",
    "detect_post_meeting_engagements",
    "search_accounts",
    "get_account",
    "search_products",
    "search_contacts",
    "get_collaboration_team",
    "update_alfred",
  ];

  for (const tool of REQUIRED_SHARED) {
    it(`shared tool "${tool}" exists in SC server`, () => {
      expect(scTools).toContain(tool);
    });

    it(`shared tool "${tool}" exists in Sales server`, () => {
      expect(salesTools).toContain(tool);
    });
  }

  // Sales-specific tools
  const SALES_ONLY = ["create_opportunity", "update_opportunity", "get_territory_pipeline", "search_users"];
  for (const tool of SALES_ONLY) {
    it(`sales-only tool "${tool}" exists in Sales server`, () => {
      expect(salesTools).toContain(tool);
    });
  }

  // SC-specific tools
  const SC_ONLY = ["assess_tech_win"];
  for (const tool of SC_ONLY) {
    it(`SC-only tool "${tool}" exists in SC server`, () => {
      expect(scTools).toContain(tool);
    });
  }
});

// ---------------------------------------------------------------------------
// Dynamics field schema compliance
// ---------------------------------------------------------------------------
describe("Dynamics field schema", () => {
  const dynamicsSrc = readSource("src/tools/dynamicsClient.ts");

  // Fields that DON'T EXIST and must NEVER appear in $select
  const FORBIDDEN_FIELDS = [
    "sn_salestage",
    "sn_businessunitlist",
    "_sn_dealchampion_value",
    "sn_iscompetitive",
    "sn_winlossreason",
    "sn_winlossnotes",
    "sn_industrysolution",
    "sn_type",
  ];

  for (const field of FORBIDDEN_FIELDS) {
    it(`forbidden field "${field}" is not used in $select queries`, () => {
      // Check for usage in $select strings — not in comments or type definitions
      const selectPattern = new RegExp(`\\$select[^"]*${field}`, "g");
      expect(dynamicsSrc).not.toMatch(selectPattern);
    });
  }

  // Correct field names that MUST be used
  const REQUIRED_FIELDS = [
    "stepname",                       // sales stage (not sn_salestage)
    "sn_opportunitybusinessunitlist", // BU list (not sn_businessunitlist)
    "sn_noncompetitive",              // competitive flag (inverted)
    "sn_winlossnodecisionreason",     // win/loss reason (not sn_winlossreason)
    "sn_netnewacv",                   // NNACV field
  ];

  for (const field of REQUIRED_FIELDS) {
    it(`correct field "${field}" is used in dynamicsClient`, () => {
      expect(dynamicsSrc).toContain(field);
    });
  }

  it("isCompetitive mapping inverts sn_noncompetitive", () => {
    expect(dynamicsSrc).toContain("!(r.sn_noncompetitive as boolean)");
  });

  it("NNACV comes from sn_netnewacv, not totalamount", () => {
    expect(dynamicsSrc).toMatch(/nnacv:\s+r\.sn_netnewacv/);
  });
});

// ---------------------------------------------------------------------------
// NNACV usage in display code
// ---------------------------------------------------------------------------
describe("NNACV display consistency", () => {
  it("sales/index.ts uses nnacv not totalamount for display", () => {
    const salesSrc = readSource("src/sales/index.ts");
    // Check pipeline display code doesn't use totalamount
    const displayLines = salesSrc.split("\n").filter(l =>
      l.includes("toLocaleString") && (l.includes("totalamount") || l.includes("nnacv"))
    );
    for (const line of displayLines) {
      expect(line).not.toContain("totalamount");
      expect(line).toContain("nnacv");
    }
  });

  it("sc/index.ts uses nnacv not totalamount for display", () => {
    const scSrc = readSource("src/sc/index.ts");
    const displayLines = scSrc.split("\n").filter(l =>
      l.includes("toLocaleString") && (l.includes("totalamount") || l.includes("nnacv"))
    );
    for (const line of displayLines) {
      expect(line).not.toContain("totalamount");
      expect(line).toContain("nnacv");
    }
  });

  it("hygieneClient.ts uses nnacv not totalamount", () => {
    const hygieneSrc = readSource("src/tools/hygieneClient.ts");
    // All opportunity value references should be nnacv
    const valueRefs = hygieneSrc.split("\n").filter(l =>
      l.includes("opportunity.") && (l.includes("totalamount") || l.includes("nnacv"))
    );
    for (const line of valueRefs) {
      expect(line).not.toContain("totalamount");
    }
  });
});

// ---------------------------------------------------------------------------
// Engagement type consistency
// ---------------------------------------------------------------------------
describe("engagement type consistency", () => {
  it("all engagement types have GUIDs", () => {
    for (const type of ALL_ENGAGEMENT_TYPES) {
      expect(ENGAGEMENT_TYPE_GUIDS).toHaveProperty(type);
    }
  });

  it("no extra GUIDs beyond defined types", () => {
    const guidKeys = Object.keys(ENGAGEMENT_TYPE_GUIDS);
    for (const key of guidKeys) {
      expect(ALL_ENGAGEMENT_TYPES).toContain(key);
    }
  });

  it("SC server create_engagement uses ALL_ENGAGEMENT_TYPES enum", () => {
    const scSrc = readSource("src/sc/index.ts");
    expect(scSrc).toContain("z.enum(ENGAGEMENT_TYPES)");
  });

  it("Sales server create_engagement uses ALL_ENGAGEMENT_TYPES enum", () => {
    const salesSrc = readSource("src/sales/index.ts");
    expect(salesSrc).toContain("z.enum(ENGAGEMENT_TYPES)");
  });
});

// ---------------------------------------------------------------------------
// Security patterns in tool descriptions
// ---------------------------------------------------------------------------
describe("tool description safety", () => {
  const scSrc = readSource("src/sc/index.ts");
  const salesSrc = readSource("src/sales/index.ts");

  it("destructive tools require confirmation in SC server", () => {
    // delete_engagement, delete_timeline_note, create_engagement should have confirmed param
    expect(scSrc).toMatch(/delete_engagement[\s\S]*?confirmed.*boolean/);
    expect(scSrc).toMatch(/delete_timeline_note[\s\S]*?confirmed.*boolean/);
  });

  it("destructive tools require confirmation in Sales server", () => {
    expect(salesSrc).toMatch(/delete_engagement[\s\S]*?confirmed.*boolean/);
    expect(salesSrc).toMatch(/delete_timeline_note[\s\S]*?confirmed.*boolean/);
    expect(salesSrc).toMatch(/create_opportunity[\s\S]*?confirmed.*boolean/);
    expect(salesSrc).toMatch(/update_opportunity[\s\S]*?confirmed.*boolean/);
  });

  it("create_engagement includes mandatory link instruction in both servers", () => {
    for (const src of [scSrc, salesSrc]) {
      expect(src).toContain("AFTER EVERY SUCCESSFUL CREATE");
      expect(src).toContain("Never omit the link");
    }
  });

  it("external data sources use externalData wrapper", () => {
    // Calendar, email, transcript, chats should all use externalData()
    for (const src of [scSrc, salesSrc]) {
      const calendarHandler = src.includes('externalData("Outlook calendar"');
      const emailHandler = src.includes('externalData("Outlook emails"');
      if (src.includes("get_calendar_events")) expect(calendarHandler).toBe(true);
      if (src.includes("search_emails")) expect(emailHandler).toBe(true);
    }
  });
});

// ---------------------------------------------------------------------------
// Import consistency
// ---------------------------------------------------------------------------
describe("import consistency", () => {
  it("sales server imports clearAuthCache for open_chrome_debug", () => {
    const salesSrc = readSource("src/sales/index.ts");
    expect(salesSrc).toContain("clearAuthCache");
  });

  it("sales server imports clearGraphTokenCache", () => {
    const salesSrc = readSource("src/sales/index.ts");
    expect(salesSrc).toContain("clearGraphTokenCache");
  });

  it("both servers import from shared module", () => {
    const scSrc = readSource("src/sc/index.ts");
    const salesSrc = readSource("src/sales/index.ts");
    expect(scSrc).toContain("from \"../shared.js\"");
    expect(salesSrc).toContain("from \"../shared.js\"");
  });
});
