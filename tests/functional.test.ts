/**
 * Functional tests — verify business logic produces correct output.
 * Covers: buildDescription, stripBullet, formatHygieneReport, engagement types, NNACV display.
 */
import { describe, it, expect } from "vitest";
import { buildDescription, stripBullet, sanitizeODataSearch } from "../src/tools/dynamicsClient.js";
import { formatHygieneReport, type HygieneResult } from "../src/tools/hygieneClient.js";
import { requireGuid, WriteRateLimiter, stripHtml, FORECAST_NAMES, SN_INTERNAL_DOMAINS, PERSONAL_EMAIL_DOMAINS, NON_CUSTOMER_DOMAINS } from "../src/shared.js";
import { ALL_ENGAGEMENT_TYPES, ENGAGEMENT_TYPE_GUIDS } from "../src/config.js";

// ---------------------------------------------------------------------------
// stripBullet
// ---------------------------------------------------------------------------
describe("stripBullet", () => {
  it("strips single bullet •", () => {
    expect(stripBullet("• some text")).toBe("some text");
  });

  it("strips double bullet • • (the reported bug)", () => {
    expect(stripBullet("• • some text")).toBe("some text");
  });

  it("strips triple+ bullets", () => {
    expect(stripBullet("• • • triple")).toBe("triple");
  });

  it("strips dash bullet", () => {
    expect(stripBullet("- some text")).toBe("some text");
  });

  it("strips asterisk bullet", () => {
    expect(stripBullet("* some text")).toBe("some text");
  });

  it("strips double dashes", () => {
    expect(stripBullet("- - double dash")).toBe("double dash");
  });

  it("strips leading whitespace + bullet", () => {
    expect(stripBullet("  • some text")).toBe("some text");
  });

  it("preserves plain text", () => {
    expect(stripBullet("plain text")).toBe("plain text");
  });

  it("preserves mid-text bullets", () => {
    expect(stripBullet("Text with • bullet in middle")).toBe("Text with • bullet in middle");
  });

  it("handles empty string", () => {
    expect(stripBullet("")).toBe("");
  });

  it("handles bullet only", () => {
    expect(stripBullet("•")).toBe("");
  });

  it("handles bullet with trailing space", () => {
    expect(stripBullet("• ")).toBe("");
  });

  it("handles bullet without space", () => {
    expect(stripBullet("•text")).toBe("text");
  });

  it("strips tab + bullet", () => {
    expect(stripBullet("\t• tabbed")).toBe("tabbed");
  });

  it("strips mixed bullet styles (• then -)", () => {
    expect(stripBullet("• - mixed")).toBe("mixed");
  });

  it("strips mixed bullet styles (- then •)", () => {
    expect(stripBullet("- • mixed")).toBe("mixed");
  });
});

// ---------------------------------------------------------------------------
// buildDescription
// ---------------------------------------------------------------------------
describe("buildDescription", () => {
  it("builds a structured description with all fields", () => {
    const result = buildDescription({
      engagementType: "Discovery",
      useCase: "ITSM",
      keyPoints: ["Scope confirmed", "Timeline agreed"],
      nextActions: ["Send proposal"],
      risks: "None",
      stakeholders: "Ahmed (VP IT), Fredrik",
    });
    expect(result).toContain("Use Case: ITSM");
    expect(result).toContain("Objectives / Key questions uncovered:");
    expect(result).toContain("• Scope confirmed");
    expect(result).toContain("• Timeline agreed");
    expect(result).toContain("Next actions:");
    expect(result).toContain("• Send proposal");
    expect(result).toContain("Risks/Help Required: None");
    expect(result).toContain("Stakeholders: Ahmed (VP IT), Fredrik");
  });

  it("uses type-specific labels for key_points", () => {
    const tw = buildDescription({ engagementType: "Technical Win", keyPoints: ["Done"] });
    expect(tw).toContain("Milestones achieved:");

    const demo = buildDescription({ engagementType: "Demo", keyPoints: ["Shown"] });
    expect(demo).toContain("Demo delivered / Customer feedback:");

    const bc = buildDescription({ engagementType: "Business Case", keyPoints: ["ROI"] });
    expect(bc).toContain("Value drivers / Quantified benefits:");
  });

  it("includes secondary points with type-specific labels", () => {
    const result = buildDescription({
      engagementType: "Demo",
      keyPoints: ["Feature shown"],
      secondaryPoints: ["Customer loved it"],
    });
    expect(result).toContain("Customer reactions / feedback:");
    expect(result).toContain("• Customer loved it");
  });

  it("includes RFx submission date", () => {
    const result = buildDescription({
      engagementType: "RFx",
      submissionDate: "2026-04-15",
    });
    expect(result).toContain("Submission date: 2026-04-15");
  });

  it("does not include submission date for non-RFx", () => {
    const result = buildDescription({
      engagementType: "Demo",
      submissionDate: "2026-04-15",
    });
    expect(result).not.toContain("Submission date");
  });

  it("never produces double bullets regardless of input", () => {
    const result = buildDescription({
      engagementType: "Discovery",
      keyPoints: ["• already bulleted", "• • double", "plain", "- dashed"],
      nextActions: ["• • • triple bullet action"],
    });
    expect(result).not.toMatch(/• •/);
    expect(result).not.toMatch(/• -/);
    // Each line with bullet should have exactly one •
    const bulletLines = result.split("\n").filter(l => l.startsWith("•"));
    for (const line of bulletLines) {
      expect(line).toMatch(/^• [^•]/);
    }
  });

  it("defaults risks to dash when not provided", () => {
    const result = buildDescription({});
    expect(result).toContain("Risks/Help Required: -");
  });

  it("falls back to 'Key points' when no engagement type", () => {
    const result = buildDescription({ keyPoints: ["item"] });
    expect(result).toContain("Key points:");
  });

  it("produces minimal output with no fields", () => {
    const result = buildDescription({});
    expect(result).toContain("Risks/Help Required: -");
    // Should not crash or produce empty string
    expect(result.length).toBeGreaterThan(0);
  });

  it("handles all engagement types without crashing", () => {
    for (const type of ALL_ENGAGEMENT_TYPES) {
      const result = buildDescription({
        engagementType: type,
        keyPoints: ["test point"],
        nextActions: ["test action"],
      });
      expect(result).toContain("test point");
      expect(result).toContain("test action");
    }
  });

  it("handles unicode and special characters in key points", () => {
    const result = buildDescription({
      keyPoints: ["Zürich meeting — confirmed ✓", "日本語テスト"],
      stakeholders: "Müller (CTO), O'Brien (VP)",
    });
    expect(result).toContain("Zürich meeting — confirmed ✓");
    expect(result).toContain("日本語テスト");
    expect(result).toContain("Müller (CTO), O'Brien (VP)");
  });
});

// ---------------------------------------------------------------------------
// sanitizeODataSearch
// ---------------------------------------------------------------------------
describe("sanitizeODataSearch", () => {
  it("passes through safe characters", () => {
    expect(sanitizeODataSearch("SITA")).toBe("SITA");
    expect(sanitizeODataSearch("PMI-2026")).toBe("PMI-2026");
  });

  it("strips SQL injection characters", () => {
    expect(sanitizeODataSearch("'; DROP TABLE--")).toBe(" DROP TABLE--");
  });

  it("strips special OData characters", () => {
    expect(sanitizeODataSearch("test(eq)")).toBe("testeq");
  });

  it("escapes single quotes", () => {
    expect(sanitizeODataSearch("O'Brien")).toBe("OBrien");
  });

  it("truncates to 100 characters", () => {
    const long = "A".repeat(200);
    expect(sanitizeODataSearch(long).length).toBe(100);
  });

  it("handles empty string", () => {
    expect(sanitizeODataSearch("")).toBe("");
  });

  it("strips unicode characters (only ASCII allowed)", () => {
    // sanitizeODataSearch allows only [a-zA-Z0-9 \-\.@_#]
    expect(sanitizeODataSearch("Zürich")).toBe("Zrich");
    expect(sanitizeODataSearch("Straße")).toBe("Strae");
  });

  it("preserves email addresses", () => {
    expect(sanitizeODataSearch("user@servicenow.com")).toBe("user@servicenow.com");
  });

  it("preserves hashes and numbers", () => {
    expect(sanitizeODataSearch("OPTY#5299816")).toBe("OPTY#5299816");
  });
});

// ---------------------------------------------------------------------------
// formatHygieneReport
// ---------------------------------------------------------------------------
describe("formatHygieneReport", () => {
  const makeResult = (overrides: Partial<HygieneResult> = {}): HygieneResult => ({
    opportunity: {
      opportunityid: "test-guid",
      name: "Test Opp",
      accountName: "Test Account",
      accountid: "acc-guid",
      statuscode: 1,
      nnacv: 500000,
    },
    engagements: [],
    missingRequired: [],
    missingOptional: [],
    status: "green",
    ...overrides,
  });

  it("generates a report with counts", () => {
    const results = [
      makeResult({ status: "red", missingRequired: ["Discovery"] }),
      makeResult({ status: "green" }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toContain("1 critical");
    expect(report).toContain("1 complete");
  });

  it("groups by account", () => {
    const results = [
      makeResult({ opportunity: { ...makeResult().opportunity, accountName: "SITA" } }),
      makeResult({ opportunity: { ...makeResult().opportunity, accountName: "SITA" } }),
      makeResult({ opportunity: { ...makeResult().opportunity, accountName: "PMI" } }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toContain("SITA");
    expect(report).toContain("PMI");
  });

  it("uses nnacv not totalamount for display", () => {
    const results = [
      makeResult({
        opportunity: { ...makeResult().opportunity, nnacv: 1500000, totalamount: 3000000 },
      }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toContain("$1.5M");
    expect(report).not.toContain("$3.0M");
  });

  it("shows missing engagements for red items", () => {
    const results = [
      makeResult({ status: "red", missingRequired: ["Discovery", "Demo"] }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toContain("missing: Discovery, Demo");
  });

  it("handles empty results array", () => {
    const report = formatHygieneReport([]);
    expect(report).toBeDefined();
    expect(report.length).toBeGreaterThan(0);
  });

  it("formats NNACV correctly: millions and thousands", () => {
    const results = [
      makeResult({ opportunity: { ...makeResult().opportunity, nnacv: 2500000 } }),
      makeResult({ opportunity: { ...makeResult().opportunity, name: "Small Opp", nnacv: 150000 } }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toContain("$2.5M");
    expect(report).toContain("$150K");
  });

  it("handles zero and null NNACV gracefully", () => {
    const results = [
      makeResult({ opportunity: { ...makeResult().opportunity, nnacv: 0 } }),
      makeResult({ opportunity: { ...makeResult().opportunity, nnacv: undefined as unknown as number } }),
    ];
    // Should not throw
    const report = formatHygieneReport(results);
    expect(report).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// Engagement type configuration
// ---------------------------------------------------------------------------
describe("engagement types", () => {
  it("has a GUID for every engagement type", () => {
    for (const type of ALL_ENGAGEMENT_TYPES) {
      expect(ENGAGEMENT_TYPE_GUIDS[type]).toBeDefined();
      expect(ENGAGEMENT_TYPE_GUIDS[type]).toMatch(
        /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/
      );
    }
  });

  it("has no duplicate GUIDs", () => {
    const guids = Object.values(ENGAGEMENT_TYPE_GUIDS);
    expect(new Set(guids).size).toBe(guids.length);
  });

  it("list is sorted alphabetically (case-insensitive)", () => {
    const sorted = [...ALL_ENGAGEMENT_TYPES].sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: "base" })
    );
    expect(ALL_ENGAGEMENT_TYPES).toEqual(sorted);
  });
});

// ---------------------------------------------------------------------------
// Forecast category names
// ---------------------------------------------------------------------------
describe("forecast categories", () => {
  it("maps known category codes", () => {
    expect(FORECAST_NAMES[100000001]).toBe("Pipeline");
    expect(FORECAST_NAMES[100000002]).toBe("Best Case");
    expect(FORECAST_NAMES[100000003]).toBe("Committed");
    expect(FORECAST_NAMES[100000004]).toBe("Omitted");
  });
});

// ---------------------------------------------------------------------------
// Email domain classification
// ---------------------------------------------------------------------------
describe("email domain classification", () => {
  it("SN_INTERNAL_DOMAINS includes servicenow.com and now.com", () => {
    expect(SN_INTERNAL_DOMAINS.has("servicenow.com")).toBe(true);
    expect(SN_INTERNAL_DOMAINS.has("now.com")).toBe(true);
  });

  it("PERSONAL_EMAIL_DOMAINS excludes servicenow.com", () => {
    expect(PERSONAL_EMAIL_DOMAINS.has("servicenow.com")).toBe(false);
  });

  it("PERSONAL_EMAIL_DOMAINS includes common providers", () => {
    expect(PERSONAL_EMAIL_DOMAINS.has("gmail.com")).toBe(true);
    expect(PERSONAL_EMAIL_DOMAINS.has("outlook.com")).toBe(true);
    expect(PERSONAL_EMAIL_DOMAINS.has("hotmail.com")).toBe(true);
  });

  it("NON_CUSTOMER_DOMAINS is union of internal + personal", () => {
    for (const d of SN_INTERNAL_DOMAINS) expect(NON_CUSTOMER_DOMAINS.has(d)).toBe(true);
    for (const d of PERSONAL_EMAIL_DOMAINS) expect(NON_CUSTOMER_DOMAINS.has(d)).toBe(true);
    expect(NON_CUSTOMER_DOMAINS.size).toBe(SN_INTERNAL_DOMAINS.size + PERSONAL_EMAIL_DOMAINS.size);
  });

  it("customer domains are not in NON_CUSTOMER_DOMAINS", () => {
    expect(NON_CUSTOMER_DOMAINS.has("sita.aero")).toBe(false);
    expect(NON_CUSTOMER_DOMAINS.has("pmi.com")).toBe(false);
    expect(NON_CUSTOMER_DOMAINS.has("straumann.com")).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// WriteRateLimiter edge cases
// ---------------------------------------------------------------------------
describe("WriteRateLimiter edge cases", () => {
  it("resets after window expires", () => {
    const limiter = new WriteRateLimiter(1, 50); // 50ms window
    limiter.check("test");
    expect(() => limiter.check("test")).toThrow("Rate limit");
    // Wait for window to expire
    return new Promise<void>((resolve) => {
      setTimeout(() => {
        expect(() => limiter.check("test")).not.toThrow();
        resolve();
      }, 60);
    });
  });

  it("different limiters are independent", () => {
    const a = new WriteRateLimiter(1, 60_000);
    const b = new WriteRateLimiter(1, 60_000);
    a.check("a");
    expect(() => b.check("b")).not.toThrow();
  });
});
