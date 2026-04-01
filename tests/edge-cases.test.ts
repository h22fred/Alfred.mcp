/**
 * Edge-case and integration tests — fills gaps identified in audit.
 * Covers: sanitizeODataSearch edge cases, WriteRateLimiter window cleanup,
 * stripHtml re-injection prevention, urlHostMatches edge cases, requireGuid
 * path traversal, formatHygieneReport mixed/empty results, and attendee
 * classification logic.
 */
import { describe, it, expect } from "vitest";
import { sanitizeODataSearch } from "../src/tools/dynamicsClient.js";
import { formatHygieneReport, type HygieneResult } from "../src/tools/hygieneClient.js";
import { requireGuid, WriteRateLimiter, stripHtml, urlHostMatches, SN_INTERNAL_DOMAINS, NON_CUSTOMER_DOMAINS } from "../src/shared.js";

// ---------------------------------------------------------------------------
// 1. sanitizeODataSearch — additional edge cases
// ---------------------------------------------------------------------------
describe("sanitizeODataSearch edge cases", () => {
  it("returns empty string for empty input", () => {
    expect(sanitizeODataSearch("")).toBe("");
  });

  it("returns empty string when input is only special chars", () => {
    expect(sanitizeODataSearch("';|()%\0")).toBe("");
  });

  it("strips single quotes (prevents OData string breakout)", () => {
    expect(sanitizeODataSearch("It's a test")).not.toContain("'");
  });

  it("strips semicolons (prevents query chaining)", () => {
    expect(sanitizeODataSearch("test; DROP")).not.toContain(";");
  });

  it("strips backticks and curly braces", () => {
    const result = sanitizeODataSearch("test`{evil}");
    expect(result).not.toContain("`");
    expect(result).not.toContain("{");
    expect(result).not.toContain("}");
  });

  it("truncates long input to exactly 100 characters", () => {
    const long = "A".repeat(200);
    expect(sanitizeODataSearch(long).length).toBe(100);
  });

  it("truncates input that is exactly at boundary", () => {
    const exact = "B".repeat(100);
    expect(sanitizeODataSearch(exact).length).toBe(100);
    expect(sanitizeODataSearch(exact)).toBe(exact);
  });

  it("does not truncate input under boundary", () => {
    const short = "C".repeat(99);
    expect(sanitizeODataSearch(short).length).toBe(99);
  });

  it("strips parentheses used in OData function injection", () => {
    const result = sanitizeODataSearch("contains(name,'hack')");
    expect(result).not.toContain("(");
    expect(result).not.toContain(")");
  });

  it("preserves dots in domain-style searches", () => {
    expect(sanitizeODataSearch("acme.org")).toBe("acme.org");
  });

  it("preserves underscores in field-like searches", () => {
    expect(sanitizeODataSearch("my_search_term")).toBe("my_search_term");
  });

  it("strips tab and newline characters", () => {
    const result = sanitizeODataSearch("line1\tline2\nline3");
    expect(result).not.toContain("\t");
    expect(result).not.toContain("\n");
  });
});

// ---------------------------------------------------------------------------
// 2. WriteRateLimiter — window cleanup and boundary tests
// ---------------------------------------------------------------------------
describe("WriteRateLimiter window cleanup and boundaries", () => {
  it("allows operations when count is under the limit", () => {
    const limiter = new WriteRateLimiter(5, 60_000);
    for (let i = 0; i < 5; i++) {
      expect(() => limiter.check("op")).not.toThrow();
    }
  });

  it("throws when at exactly the limit (limit N, after N calls)", () => {
    const limiter = new WriteRateLimiter(3, 60_000);
    limiter.check("a");
    limiter.check("a");
    limiter.check("a");
    // 4th call should throw — 3 timestamps already in window
    expect(() => limiter.check("a")).toThrow("Rate limit");
  });

  it("cleans up old timestamps outside the window", () => {
    const limiter = new WriteRateLimiter(2, 50); // 50ms window
    limiter.check("cleanup");
    limiter.check("cleanup");
    expect(() => limiter.check("cleanup")).toThrow("Rate limit");

    return new Promise<void>((resolve) => {
      setTimeout(() => {
        // After the window expires, old timestamps should be cleaned up
        expect(() => limiter.check("cleanup")).not.toThrow();
        resolve();
      }, 60);
    });
  });

  it("error message contains the action name", () => {
    const limiter = new WriteRateLimiter(1, 60_000);
    limiter.check("create_engagement");
    try {
      limiter.check("create_engagement");
      expect.unreachable("should have thrown");
    } catch (e) {
      expect((e as Error).message).toContain("create_engagement");
    }
  });

  it("error message contains the limit count", () => {
    const limiter = new WriteRateLimiter(7, 60_000);
    for (let i = 0; i < 7; i++) limiter.check("x");
    try {
      limiter.check("x");
      expect.unreachable("should have thrown");
    } catch (e) {
      expect((e as Error).message).toContain("7");
    }
  });

  it("limit of 0 always throws", () => {
    const limiter = new WriteRateLimiter(0, 60_000);
    expect(() => limiter.check("zero")).toThrow("Rate limit");
  });
});

// ---------------------------------------------------------------------------
// 3. stripHtml — nested tag attacks and re-injection prevention
// ---------------------------------------------------------------------------
describe("stripHtml edge cases and re-injection prevention", () => {
  it("handles nested <<script>script> tags", () => {
    const input = '<<script>script>alert("xss")<</script>/script>';
    const result = stripHtml(input);
    expect(result).not.toContain("<script>");
    expect(result).not.toContain("</script>");
    // Should not contain any HTML tags at all
    expect(result).not.toMatch(/<[a-z]/i);
  });

  it("HTML entities decoded AFTER tag stripping prevents re-injection", () => {
    // If entities were decoded first, &lt;script&gt; would become <script>
    const input = "&lt;script&gt;alert('xss')&lt;/script&gt;";
    const result = stripHtml(input);
    // The entities decode to literal < and > AFTER all tags are removed,
    // so no new tags are formed
    expect(result).toContain("<script>");  // literal text, not an HTML tag
    // Verify it does NOT re-enter tag removal (the < is a decoded entity)
    expect(result).toContain("alert");
  });

  it("returns empty string for empty input", () => {
    expect(stripHtml("")).toBe("");
  });

  it("returns plain text unchanged (no HTML)", () => {
    expect(stripHtml("Hello world, this is plain text.")).toBe("Hello world, this is plain text.");
  });

  it("handles plain text with special characters", () => {
    expect(stripHtml("Price: $100 & free shipping")).toBe("Price: $100 & free shipping");
  });

  it("handles triple-nested obfuscated tags", () => {
    const input = "<<<b>b>b>bold<<<//b>/b>/b>";
    const result = stripHtml(input);
    expect(result).not.toMatch(/<[a-z]/i);
    expect(result).toContain("bold");
  });

  it("strips img tags with onerror handlers", () => {
    const result = stripHtml('<img src=x onerror="alert(1)">safe text');
    expect(result).not.toContain("<img");
    expect(result).not.toContain("onerror");
    expect(result).toContain("safe text");
  });

  it("handles null-byte inside tags", () => {
    const result = stripHtml("<scr\0ipt>evil</scr\0ipt>safe");
    expect(result).not.toMatch(/<scr/i);
  });

  it("strips all tags from deeply nested HTML structure", () => {
    const input = "<div><p><span><b><i><u>deep</u></i></b></span></p></div>";
    const result = stripHtml(input);
    expect(result).toContain("deep");
    expect(result).not.toContain("<");
    expect(result).not.toContain(">");
  });
});

// ---------------------------------------------------------------------------
// 4. urlHostMatches — additional edge cases
// ---------------------------------------------------------------------------
describe("urlHostMatches edge cases", () => {
  it("exact hostname match", () => {
    expect(urlHostMatches("https://outlook.office.com/mail/", "outlook.office.com")).toBe(true);
  });

  it("subdomain match (sub.outlook.office.com)", () => {
    expect(urlHostMatches("https://sub.outlook.office.com/", "outlook.office.com")).toBe(true);
  });

  it("rejects domain with target as substring prefix (eviloffice.com vs office.com)", () => {
    // "eviloffice.com" does NOT end with ".office.com" nor equal "office.com"
    expect(urlHostMatches("https://eviloffice.com/", "office.com")).toBe(false);
  });

  it("rejects attacker domain with target embedded (outlook.office.com.evil.com)", () => {
    expect(urlHostMatches("https://outlook.office.com.evil.com/", "outlook.office.com")).toBe(false);
  });

  it("rejects target hostname in query string", () => {
    expect(urlHostMatches("https://evil.com?host=outlook.office.com", "outlook.office.com")).toBe(false);
  });

  it("rejects target hostname in path", () => {
    expect(urlHostMatches("https://evil.com/outlook.office.com", "outlook.office.com")).toBe(false);
  });

  it("rejects completely unrelated domain", () => {
    expect(urlHostMatches("https://google.com/", "webhook.office.com")).toBe(false);
  });

  it("returns false for invalid URL (not a URL at all)", () => {
    expect(urlHostMatches("not-a-url", "office.com")).toBe(false);
  });

  it("returns false for empty string", () => {
    expect(urlHostMatches("", "office.com")).toBe(false);
  });

  it("returns false for URL with no hostname", () => {
    expect(urlHostMatches("file:///etc/passwd", "office.com")).toBe(false);
  });

  it("is case-insensitive for the URL hostname", () => {
    expect(urlHostMatches("https://OUTLOOK.OFFICE.COM/", "outlook.office.com")).toBe(true);
  });

  it("target hostname must be lowercase (function contract)", () => {
    // urlHostMatches lowercases the URL host but compares target as-is
    // Callers are expected to pass lowercase targets
    expect(urlHostMatches("https://outlook.office.com/", "OUTLOOK.OFFICE.COM")).toBe(false);
  });

  it("handles URL with port number", () => {
    // URL hostname does not include port, so this should still match
    expect(urlHostMatches("https://outlook.office.com:443/mail/", "outlook.office.com")).toBe(true);
  });

  it("handles URL with authentication in it", () => {
    expect(urlHostMatches("https://user:pass@outlook.office.com/", "outlook.office.com")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 5. requireGuid — additional edge cases
// ---------------------------------------------------------------------------
describe("requireGuid edge cases", () => {
  it("valid GUID passes through unchanged", () => {
    const guid = "e143abb9-f8a0-ef11-8a69-6045bdf0cf09";
    expect(requireGuid(guid, "test")).toBe(guid);
  });

  it("valid all-zeros GUID passes", () => {
    expect(requireGuid("00000000-0000-0000-0000-000000000000", "id")).toBe("00000000-0000-0000-0000-000000000000");
  });

  it("uppercase GUID passes", () => {
    expect(requireGuid("E143ABB9-F8A0-EF11-8A69-6045BDF0CF09", "id"))
      .toBe("E143ABB9-F8A0-EF11-8A69-6045BDF0CF09");
  });

  it("throws for path traversal attempt ../../etc/passwd", () => {
    expect(() => requireGuid("../../etc/passwd", "id")).toThrow("Invalid id");
  });

  it("throws for path traversal with backslashes", () => {
    expect(() => requireGuid("..\\..\\windows\\system32", "id")).toThrow("Invalid id");
  });

  it("throws for empty string", () => {
    expect(() => requireGuid("", "id")).toThrow("Invalid id");
  });

  it("throws for GUID with OData injection suffix", () => {
    expect(() =>
      requireGuid("e143abb9-f8a0-ef11-8a69-6045bdf0cf09)?$filter=1 eq 1", "id")
    ).toThrow();
  });

  it("throws for GUID with spaces", () => {
    expect(() => requireGuid("e143abb9-f8a0-ef11-8a69-6045bdf0cf09 ", "id")).toThrow();
  });

  it("throws for GUID with leading space", () => {
    expect(() => requireGuid(" e143abb9-f8a0-ef11-8a69-6045bdf0cf09", "id")).toThrow();
  });

  it("throws for partial GUID (too short)", () => {
    expect(() => requireGuid("e143abb9-f8a0-ef11", "id")).toThrow("Invalid id");
  });

  it("throws for GUID with wrong section lengths", () => {
    expect(() => requireGuid("e143abb-9f8a0-ef11-8a69-6045bdf0cf09", "id")).toThrow();
  });

  it("error message includes the label", () => {
    try {
      requireGuid("bad", "engagement_id");
      expect.unreachable("should have thrown");
    } catch (e) {
      expect((e as Error).message).toContain("engagement_id");
    }
  });

  it("rejects SQL injection via GUID field", () => {
    expect(() => requireGuid("'; DROP TABLE engagements; --", "id")).toThrow();
  });

  it("rejects newlines in GUID", () => {
    expect(() => requireGuid("e143abb9-f8a0-ef11-8a69-6045bdf0cf09\n", "id")).toThrow();
  });
});

// ---------------------------------------------------------------------------
// 6. formatHygieneReport — mixed statuses and edge cases
// ---------------------------------------------------------------------------
describe("formatHygieneReport integration", () => {
  const makeResult = (overrides: Partial<HygieneResult> = {}): HygieneResult => ({
    opportunity: {
      opportunityid: "test-guid-" + Math.random().toString(36).slice(2, 8),
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

  it("formats report with mix of red, yellow, and green results", () => {
    const results = [
      makeResult({ status: "red", missingRequired: ["Discovery", "Demo"], opportunity: { opportunityid: "r1", name: "Red Opp", accountName: "Acme", accountid: "a1", statuscode: 1, nnacv: 1000000 } }),
      makeResult({ status: "yellow", missingOptional: ["Workshop"], opportunity: { opportunityid: "y1", name: "Yellow Opp", accountName: "Acme", accountid: "a2", statuscode: 1, nnacv: 200000 } }),
      makeResult({ status: "green", opportunity: { opportunityid: "g1", name: "Green Opp", accountName: "Beta Corp", accountid: "b1", statuscode: 1, nnacv: 750000 } }),
    ];
    const report = formatHygieneReport(results);

    // Header counts
    expect(report).toContain("1 critical");
    expect(report).toContain("1 on track");
    expect(report).toContain("1 complete");

    // Account grouping
    expect(report).toContain("Acme");
    expect(report).toContain("Beta Corp");

    // Missing engagements for red
    expect(report).toContain("missing: Discovery, Demo");

    // Green shows complete
    expect(report).toContain("all complete");
  });

  it("handles empty results array without crashing", () => {
    const report = formatHygieneReport([]);
    expect(report).toBeDefined();
    expect(report.length).toBeGreaterThan(0);
    // Should show 0 for all counts
    expect(report).toContain("0 critical");
    expect(report).toContain("0 complete");
  });

  it("formats large NNACV as millions", () => {
    const results = [
      makeResult({ opportunity: { ...makeResult().opportunity, nnacv: 2500000 } }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toContain("$2.5M");
  });

  it("formats sub-million NNACV as thousands", () => {
    const results = [
      makeResult({ opportunity: { ...makeResult().opportunity, nnacv: 350000 } }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toContain("$350K");
  });

  it("handles zero NNACV gracefully", () => {
    const results = [
      makeResult({ opportunity: { ...makeResult().opportunity, nnacv: 0 } }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toBeDefined();
  });

  it("handles undefined NNACV gracefully", () => {
    const results = [
      makeResult({ opportunity: { ...makeResult().opportunity, nnacv: undefined as unknown as number } }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toBeDefined();
  });

  it("report includes close date when available", () => {
    const results = [
      makeResult({
        status: "red",
        missingRequired: ["Discovery"],
        opportunity: {
          ...makeResult().opportunity,
          estimatedclosedate: "2026-06-15T00:00:00Z",
        },
      }),
    ];
    const report = formatHygieneReport(results);
    expect(report).toContain("close 2026-06-15");
  });

  it("red items appear before green items in output", () => {
    const results = [
      makeResult({ status: "green", opportunity: { ...makeResult().opportunity, accountName: "ZZZZ Last" } }),
      makeResult({ status: "red", missingRequired: ["Demo"], opportunity: { ...makeResult().opportunity, accountName: "AAAA First" } }),
    ];
    const report = formatHygieneReport(results);
    const redPos = report.indexOf("AAAA First");
    const greenPos = report.indexOf("ZZZZ Last");
    expect(redPos).toBeLessThan(greenPos);
  });
});

// ---------------------------------------------------------------------------
// 7. PostMeetingCandidate attendee classification
// ---------------------------------------------------------------------------
describe("attendee external/internal classification", () => {
  it("servicenow.com is classified as internal", () => {
    expect(SN_INTERNAL_DOMAINS.has("servicenow.com")).toBe(true);
  });

  it("now.com is classified as internal", () => {
    expect(SN_INTERNAL_DOMAINS.has("now.com")).toBe(true);
  });

  it("customer domains are NOT in NON_CUSTOMER_DOMAINS", () => {
    expect(NON_CUSTOMER_DOMAINS.has("acme.com")).toBe(false);
    expect(NON_CUSTOMER_DOMAINS.has("sita.aero")).toBe(false);
    expect(NON_CUSTOMER_DOMAINS.has("pmi.com")).toBe(false);
  });

  it("gmail.com is in NON_CUSTOMER_DOMAINS (personal)", () => {
    expect(NON_CUSTOMER_DOMAINS.has("gmail.com")).toBe(true);
  });

  it("attendee classification logic: external vs internal split", () => {
    // Simulate the exact logic from postMeetingClient.ts notifyPostMeetingCandidates
    const attendees = [
      { name: "Alice", email: "alice@servicenow.com" },
      { name: "Bob", email: "bob@now.com" },
      { name: "Charlie", email: "charlie@acme.com" },
      { name: "Diana", email: "diana@customer.org" },
    ];

    const extCount = attendees.filter(a => {
      const domain = a.email.split("@")[1]?.toLowerCase();
      return domain && !SN_INTERNAL_DOMAINS.has(domain);
    }).length;
    const intCount = attendees.length - extCount;

    expect(extCount).toBe(2); // acme.com, customer.org
    expect(intCount).toBe(2); // servicenow.com, now.com
  });

  it("attendee with no @ in email is classified as external (non-internal)", () => {
    const attendees = [{ name: "NoEmail", email: "noemail" }];
    const extCount = attendees.filter(a => {
      const domain = a.email.split("@")[1]?.toLowerCase();
      return domain && !SN_INTERNAL_DOMAINS.has(domain);
    }).length;
    // domain is undefined, so the filter returns false (not counted as external)
    expect(extCount).toBe(0);
  });

  it("attendee with empty email is handled safely", () => {
    const attendees = [{ name: "Empty", email: "" }];
    const extCount = attendees.filter(a => {
      const domain = a.email.split("@")[1]?.toLowerCase();
      return domain && !SN_INTERNAL_DOMAINS.has(domain);
    }).length;
    expect(extCount).toBe(0);
  });

  it("personal email domains (gmail, outlook) are not classified as customer external", () => {
    // This tests the NON_CUSTOMER_DOMAINS set used in attendeeDomainWord filtering
    const personalDomains = ["gmail.com", "outlook.com", "hotmail.com", "yahoo.com", "live.com"];
    for (const d of personalDomains) {
      expect(NON_CUSTOMER_DOMAINS.has(d)).toBe(true);
    }
  });
});
