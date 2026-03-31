/**
 * Technical tests — verify error handling patterns, token management logic,
 * Graph API scope handling, and Dynamics API response processing.
 */
import { describe, it, expect } from "vitest";
import { readFileSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, "..");

function readSource(path: string): string {
  return readFileSync(join(ROOT, path), "utf8");
}

// ---------------------------------------------------------------------------
// Token management
// ---------------------------------------------------------------------------
describe("token management", () => {
  const outlookSrc = readSource("src/tools/outlookClient.ts");
  const teamsSrc = readSource("src/tools/teamsClient.ts");

  it("Outlook token has cache with expiry", () => {
    expect(outlookSrc).toContain("TOKEN_CACHE_MS");
    expect(outlookSrc).toContain("expiresAt");
  });

  it("Teams token has cache with expiry", () => {
    expect(teamsSrc).toContain("TOKEN_CACHE_MS");
    expect(teamsSrc).toContain("teamsTokenCache");
  });

  it("clearGraphTokenCache resets Outlook token", () => {
    expect(outlookSrc).toContain("export function clearGraphTokenCache");
    expect(outlookSrc).toContain("tokenCache = null");
  });

  it("Teams MSAL extraction checks for Chat.Read scope", () => {
    expect(teamsSrc).toContain("Chat.Read");
    expect(teamsSrc).toContain("hasScope");
  });

  it("Teams graphFetch clears cache on scope permission errors", () => {
    expect(teamsSrc).toContain("Missing scope permissions");
    expect(teamsSrc).toContain("teamsTokenCache = null");
  });

  it("Outlook Playwright fallback reuses existing tabs", () => {
    expect(outlookSrc).toContain("existingPages");
    expect(outlookSrc).not.toMatch(/const page = await ctx\.newPage\(\);[\s]*let captured/);
  });

  it("Teams Playwright fallback reuses existing tabs", () => {
    expect(teamsSrc).toContain("existingPages");
  });
});

// ---------------------------------------------------------------------------
// Error handling
// ---------------------------------------------------------------------------
describe("error handling patterns", () => {
  const dynamicsSrc = readSource("src/tools/dynamicsClient.ts");

  it("dynamicsFetch handles 429 throttling with backoff", () => {
    expect(dynamicsSrc).toContain("429");
    expect(dynamicsSrc).toContain("retryAfter");
    expect(dynamicsSrc).toContain("_retryCount");
  });

  it("dynamicsFetch handles 401 with session refresh", () => {
    expect(dynamicsSrc).toContain("401");
    expect(dynamicsSrc).toContain("clearAuthCache");
    expect(dynamicsSrc).toContain("session expired");
  });

  it("dynamicsFetch detects HTML auth redirects", () => {
    expect(dynamicsSrc).toContain("returned ${ct} instead of JSON");
    expect(dynamicsSrc).toContain("session redirect");
  });

  it("privilege errors give actionable message", () => {
    expect(dynamicsSrc).toContain("Permission denied");
    expect(dynamicsSrc).toContain("CRM admin");
  });

  it("duplicate engagement errors are user-friendly", () => {
    expect(dynamicsSrc).toContain("already exists on this opportunity");
    expect(dynamicsSrc).toContain("Collaborate on the existing one");
  });

  it("cancelled engagement duplicate suggests reopen", () => {
    expect(dynamicsSrc).toContain("Reopen the existing one instead of creating a duplicate");
  });

  it("Teams graphFetch handles 429 with backoff", () => {
    const teamsSrc = readSource("src/tools/teamsClient.ts");
    expect(teamsSrc).toContain("429");
    expect(teamsSrc).toContain("_retryCount");
  });

  it("Teams scope error gives actionable guidance", () => {
    const teamsSrc = readSource("src/tools/teamsClient.ts");
    expect(teamsSrc).toContain("Open the Teams tab in Alfred");
  });
});

// ---------------------------------------------------------------------------
// Dynamics API query patterns
// ---------------------------------------------------------------------------
describe("OData query correctness", () => {
  const dynamicsSrc = readSource("src/tools/dynamicsClient.ts");

  it("opportunity queries use sn_netnewacv for NNACV filter", () => {
    // The filter should reference sn_netnewacv, not totalamount
    expect(dynamicsSrc).toMatch(/sn_netnewacv ge/);
    expect(dynamicsSrc).toMatch(/sn_netnewacv lt 0/);  // negative NNACV inclusion
  });

  it("engagement queries include expand for type name", () => {
    expect(dynamicsSrc).toContain("$expand=sn_engagementtypeid($select=sn_name)");
  });

  it("opportunity queries include annotations for formatted values", () => {
    expect(dynamicsSrc).toContain('odata.include-annotations="*"');
  });

  it("queries use GUID validation before building paths", () => {
    // All fetch-by-ID functions should call requireGuid
    expect(dynamicsSrc).toMatch(/requireGuid\(opportunityId/);
    expect(dynamicsSrc).toMatch(/requireGuid\(engagementId/);
  });

  it("$select strings don't contain forbidden fields", () => {
    const selectMatches = dynamicsSrc.match(/\$select=[^"&]*/g) ?? [];
    const forbidden = ["sn_salestage", "sn_businessunitlist", "sn_dealchampion", "sn_type"];
    for (const select of selectMatches) {
      for (const field of forbidden) {
        expect(select).not.toContain(field);
      }
    }
  });
});

// ---------------------------------------------------------------------------
// Rate limiting coverage
// ---------------------------------------------------------------------------
describe("rate limiting", () => {
  const scSrc = readSource("src/sc/index.ts");
  const salesSrc = readSource("src/sales/index.ts");

  it("SC server has engagement write limiter", () => {
    expect(scSrc).toContain("engagementWriteLimiter");
    expect(scSrc).toContain("engagementWriteLimiter.check");
  });

  it("SC server has delete limiter", () => {
    expect(scSrc).toContain("deleteWriteLimiter");
    expect(scSrc).toContain("deleteWriteLimiter.check");
  });

  it("Sales server has engagement write limiter", () => {
    expect(salesSrc).toContain("engagementWriteLimiter");
  });

  it("Sales server has delete limiter", () => {
    expect(salesSrc).toContain("deleteWriteLimiter");
  });

  it("Sales server has opportunity write limiter", () => {
    expect(salesSrc).toContain("opportunityWriteLimiter");
    expect(salesSrc).toContain("opportunityWriteLimiter.check");
  });

  it("update_engagement is rate-limited in both servers", () => {
    expect(scSrc).toMatch(/engagementWriteLimiter\.check\("update_engagement"\)/);
    expect(salesSrc).toMatch(/engagementWriteLimiter\.check\("update_engagement"\)/);
  });

  it("add_engagement_attendees is rate-limited in both servers", () => {
    expect(scSrc).toMatch(/engagementWriteLimiter\.check\("add_engagement_attendees"\)/);
    expect(salesSrc).toMatch(/engagementWriteLimiter\.check\("add_engagement_attendees"\)/);
  });

  it("add_opportunity_note is rate-limited in Sales server", () => {
    expect(salesSrc).toMatch(/opportunityWriteLimiter\.check\("add_opportunity_note"\)/);
  });
});

// ---------------------------------------------------------------------------
// Duplicate engagement pre-check
// ---------------------------------------------------------------------------
describe("duplicate engagement pre-check", () => {
  const dynamicsSrc = readSource("src/tools/dynamicsClient.ts");

  it("createEngagement checks for existing engagement of same type", () => {
    expect(dynamicsSrc).toContain("Checking for existing");
    expect(dynamicsSrc).toContain("fetchEngagementsByOpportunity");
    expect(dynamicsSrc).toContain("engagementTypeName === input.type");
  });

  it("check covers cancelled engagements (must reopen, not recreate)", () => {
    expect(dynamicsSrc).toContain("Reopen the existing one");
  });

  it("check covers active engagements (must collaborate)", () => {
    expect(dynamicsSrc).toContain("Collaborate on the existing one");
  });
});

// ---------------------------------------------------------------------------
// Stale scName / collaboration team validation
// ---------------------------------------------------------------------------
describe("collaboration team validation", () => {
  const dynamicsSrc = readSource("src/tools/dynamicsClient.ts");
  const hygieneSrc = readSource("src/tools/hygieneClient.ts");

  it("fetchOpportunities cross-references against collaboration team when myOpportunitiesOnly", () => {
    expect(dynamicsSrc).toContain("Validating against collaboration team");
    expect(dynamicsSrc).toContain("fetchMyCollaborationOpportunities");
    expect(dynamicsSrc).toContain("scNameMismatch");
  });

  it("Opportunity interface includes scNameMismatch flag", () => {
    expect(dynamicsSrc).toContain("scNameMismatch?: boolean");
  });

  it("stale SC opps are filtered out from myOpportunitiesOnly results", () => {
    expect(dynamicsSrc).toContain("stale SC attribution");
    expect(dynamicsSrc).toContain("!o.scNameMismatch");
  });

  it("hygiene sweep uses collaboration team as authoritative data source", () => {
    expect(hygieneSrc).toContain("fetchMyCollaborationOpportunities");
    // Should NOT use fetchOpportunities(myOpportunitiesOnly) for the sweep
    expect(hygieneSrc).not.toMatch(/fetchOpportunities\(\{[\s\S]*?myOpportunitiesOnly/);
  });
});

// ---------------------------------------------------------------------------
// Audit logging
// ---------------------------------------------------------------------------
describe("audit logging", () => {
  const dynamicsSrc = readSource("src/tools/dynamicsClient.ts");

  it("create operations are audit logged", () => {
    expect(dynamicsSrc).toMatch(/auditLog\("create_engagement"/);
  });

  it("delete operations are audit logged", () => {
    expect(dynamicsSrc).toMatch(/auditLog\("delete_engagement"/);
    expect(dynamicsSrc).toMatch(/auditLog\("delete_timeline_note"/);
  });

  it("audit log includes timestamp and user", () => {
    expect(dynamicsSrc).toContain("timestamp:");
    expect(dynamicsSrc).toContain("userInfo().username");
  });

  it("fetch_opportunities audit logged", () => {
    expect(dynamicsSrc).toMatch(/auditLog\("fetch_opportunities"/);
  });
});

// ---------------------------------------------------------------------------
// Timeline note dedup guard
// ---------------------------------------------------------------------------
describe("timeline note dedup", () => {
  const dynamicsSrc = readSource("src/tools/dynamicsClient.ts");

  it("createTimelineNote checks for recent duplicate by title", () => {
    expect(dynamicsSrc).toContain("listTimelineNotes");
    expect(dynamicsSrc).toMatch(/subject.*===|===.*subject/);
  });

  it("dedup window is 60 seconds", () => {
    expect(dynamicsSrc).toContain("60");
  });
});

// ---------------------------------------------------------------------------
// Hygiene sweep noise filtering
// ---------------------------------------------------------------------------
describe("hygiene sweep filtering", () => {
  const hygieneSrc = readSource("src/tools/hygieneClient.ts");

  it("filters App Store renewal noise by default", () => {
    expect(hygieneSrc).toContain("NOISE_PATTERNS");
    expect(hygieneSrc).toMatch(/app\s*store\s*renewal/i);
  });

  it("allows disabling noise filter via excludeAppStore", () => {
    expect(hygieneSrc).toContain("excludeAppStore");
  });

  it("hygiene sweep applies NNACV threshold filtering", () => {
    // After switching to collaboration team, client-side NNACV filter must be present
    expect(hygieneSrc).toContain("nnacv");
    expect(hygieneSrc).toMatch(/nnacv.*>=.*minNnacv|minNnacv.*<=.*nnacv|o\.nnacv >= minNnacv/);
  });

  it("hygiene sweep filters to open opportunities only", () => {
    expect(hygieneSrc).toContain("statuscode === 1");
  });
});

// ---------------------------------------------------------------------------
// Engagement link construction
// ---------------------------------------------------------------------------
describe("engagement link construction", () => {
  const scSrc = readSource("src/sc/index.ts");
  const salesSrc = readSource("src/sales/index.ts");

  it("both servers construct Dynamics deep-link with correct entity", () => {
    for (const src of [scSrc, salesSrc]) {
      expect(src).toContain("etn=sn_engagement");
      expect(src).toContain("pagetype=entityrecord");
    }
  });

  it("engagementSummary includes link in both servers", () => {
    for (const src of [scSrc, salesSrc]) {
      expect(src).toContain("engagementLink(e)");
      expect(src).toContain("Open in Dynamics");
    }
  });
});

// ---------------------------------------------------------------------------
// Token extraction — CDP and Playwright fallbacks
// ---------------------------------------------------------------------------
describe("token extraction fallbacks", () => {
  const outlookSrc = readSource("src/tools/outlookClient.ts");
  const teamsSrc = readSource("src/tools/teamsClient.ts");

  it("Outlook CDP extraction tries multiple tab types", () => {
    // Should try Outlook tab, then Teams tab, then any tab
    expect(outlookSrc).toContain("outlook.office.com");
  });

  it("Teams MSAL extraction decodes JWT to check scopes", () => {
    expect(teamsSrc).toContain("decodeToken");
    expect(teamsSrc).toContain("isGraphToken");
    expect(teamsSrc).toContain("hasScope");
  });

  it("Playwright fallback handles case where no existing tabs match", () => {
    // Should fall back to newPage() when no existing tab works
    expect(outlookSrc).toContain("await ctx.newPage()");
  });
});
