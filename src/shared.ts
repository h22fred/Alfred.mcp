import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Validate a Dynamics GUID. Throws if invalid — prevents path manipulation in API URLs. */
export function requireGuid(value: string, label: string): string {
  if (!GUID_RE.test(value)) throw new Error(`Invalid ${label}: expected a Dynamics GUID.`);
  return value;
}

/** Create a progress callback that logs to stderr and sends MCP logging messages. */
export function makeProgress(srv: McpServer) {
  return (msg: string) => {
    console.error(`[progress] ${msg}`);
    srv.server.sendLoggingMessage({ level: "info", data: msg });
  };
}

/** Simple in-process write rate limiter — max N writes per rolling window. */
export class WriteRateLimiter {
  private timestamps: number[] = [];
  constructor(private readonly max: number, private readonly windowMs: number) {}
  check(action: string): void {
    const now = Date.now();
    this.timestamps = this.timestamps.filter(t => now - t < this.windowMs);
    if (this.timestamps.length >= this.max) {
      throw new Error(
        `Rate limit: no more than ${this.max} ${action} operations per ${this.windowMs / 60000} minutes. ` +
        `Please review what you are doing before continuing.`
      );
    }
    this.timestamps.push(now);
  }
}

/**
 * Strip HTML tags and decode common entities to readable plain text.
 * Uses iterative stripping to handle nested/obfuscated tags like <<script>script>.
 */
export function stripHtml(html: string): string {
  let text = html;

  // Iteratively remove script/style blocks (handles nested obfuscation)
  let prev: string;
  do {
    prev = text;
    text = text.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
    text = text.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "");
  } while (text !== prev);

  // Convert structural tags to whitespace
  text = text
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<\/th>/gi, " | ")
    .replace(/<\/td>/gi, " | ");

  // Iteratively strip all remaining tags (handles <<tag>tag> nesting)
  do {
    prev = text;
    text = text.replace(/<[^>]+>/g, "");
  } while (text !== prev);

  // Decode HTML entities AFTER all tags are removed (prevents re-injection via &lt;script&gt;)
  text = text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&nbsp;/g, " ")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");

  return text.replace(/\n{3,}/g, "\n\n").trim();
}

/** Check if a URL's hostname ends with the given suffix (prevents substring bypass attacks). */
export function urlHostMatches(url: string, hostname: string): boolean {
  try {
    const host = new URL(url).hostname.toLowerCase();
    return host === hostname || host.endsWith("." + hostname);
  } catch {
    return false;
  }
}

/** ServiceNow internal email domains — used to classify attendees as internal vs external. */
export const SN_INTERNAL_DOMAINS = new Set(["servicenow.com", "now.com"]);

/** Personal/generic email domains to skip when matching attendees to customer accounts. */
export const PERSONAL_EMAIL_DOMAINS = new Set([
  "gmail.com", "outlook.com", "hotmail.com", "yahoo.com", "live.com",
  "microsoft.com", "google.com",
]);

/** Combined set: all domains that are NOT customer-specific. */
export const NON_CUSTOMER_DOMAINS = new Set([...SN_INTERNAL_DOMAINS, ...PERSONAL_EMAIL_DOMAINS]);

export const FORECAST_NAMES: Record<number, string> = {
  100000001: "Pipeline",
  100000002: "Best Case",
  100000003: "Committed",
  100000004: "Omitted",
};
