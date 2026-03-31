/**
 * Functional tests — verify business logic produces correct output.
 * Covers: buildDescription, stripBullet, formatHygieneReport, engagement types, NNACV display.
 */
import { describe, it, expect } from "vitest";
import { buildDescription, stripBullet, sanitizeODataSearch } from "../src/tools/dynamicsClient.js";
import { formatHygieneReport, type HygieneResult } from "../src/tools/hygieneClient.js";
import { requireGuid, WriteRateLimiter, stripHtml, FORECAST_NAMES } from "../src/shared.js";
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
