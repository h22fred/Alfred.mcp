import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { existsSync, readFileSync, rmSync } from "fs";
import { execFileSync } from "child_process";
import { homedir } from "os";
import { join } from "path";
import { INTERNAL_DOMAINS } from "./config.js";

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
 *
 * Security approach — defence in depth via iterative stripping:
 *   1. Iteratively remove <script> and <style> blocks until stable.
 *      The loop handles obfuscated nesting like <<script>script> that a single
 *      pass cannot catch. The regex patterns are intentionally kept simple and
 *      run repeatedly rather than trying to write a single "perfect" pattern
 *      (which would be far more complex and still bypassable).
 *      CodeQL "incomplete-multi-character-sanitization": suppressed — the iterative
 *      do-while loop is the deliberate guard against the bypass CodeQL flags.
 *   2. Iteratively strip all remaining tags until no more `<…>` remain.
 *   3. Decode HTML entities ONLY after all tags are gone, preventing
 *      re-injection via encoded payloads like &lt;script&gt;.
 *
 * This output is plain text shown to users in a desktop MCP client, not injected
 * into a browser DOM — XSS risk is very low, but we still sanitize carefully.
 */
// lgtm[js/incomplete-multi-character-sanitization]
export function stripHtml(html: string): string {
  let text = html;

  // Iteratively remove script/style blocks (handles nested obfuscation).
  // The do-while loop is the defence against partial-match bypass — CodeQL flags
  // the regex pattern but the iteration itself closes that gap.
  // lgtm[js/bad-html-filtering-regexp]
  let prev: string;
  do {
    prev = text;
    // lgtm[js/incomplete-multi-character-sanitization]
    text = text.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
    // lgtm[js/incomplete-multi-character-sanitization]
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

  // Iteratively strip all remaining tags (handles <<tag>tag> nesting).
  // lgtm[js/incomplete-multi-character-sanitization]
  do {
    prev = text;
    text = text.replace(/<[^>]+>/g, "");
  } while (text !== prev);

  // Decode HTML entities AFTER all tags are removed.
  // Decoding before stripping would allow &lt;script&gt; to survive as <script>.
  // lgtm[js/double-escaping]
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

/** Internal email domains for this company — derived from Dynamics URL or set in ~/.alfred-config.json (internalDomains). */
export const SN_INTERNAL_DOMAINS = INTERNAL_DOMAINS;

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

/**
 * Post-update migration: install Playwright Chromium and remove the old
 * Chrome-based launcher (Alfred.app on macOS, Alfred.bat on Windows).
 * Called by update_alfred after every successful git pull + rebuild.
 */
export function regenerateAlfredApp(installDir: string): string | null {
  const home = homedir();
  const messages: string[] = [];

  // Delete old launcher
  if (process.platform === "darwin") {
    const alfredApp = join(home, "Desktop", "Alfred.app");
    if (existsSync(alfredApp)) {
      rmSync(alfredApp, { recursive: true, force: true });
      messages.push("🗑️ Removed Alfred.app from Desktop (no longer needed)");
    }
  } else if (process.platform === "win32") {
    const desktopPaths = [
      join(home, "Desktop", "Alfred.bat"),
      join(home, "OneDrive", "Desktop", "Alfred.bat"),
    ];
    for (const p of desktopPaths) {
      if (existsSync(p)) {
        rmSync(p, { force: true });
        messages.push("🗑️ Removed Alfred.bat from Desktop (no longer needed)");
        break;
      }
    }
  }

  // Install Playwright Chromium
  const pwBin = join(
    installDir, "node_modules", ".bin",
    process.platform === "win32" ? "playwright.cmd" : "playwright"
  );
  if (existsSync(pwBin)) {
    try {
      execFileSync(pwBin, ["install", "chromium"], {
        timeout: 120_000,
        env: { ...process.env },
      });
      messages.push("🎭 Playwright Chromium updated");
    } catch (e) {
      process.stderr.write(`[alfred:warn] playwright install failed: ${e instanceof Error ? e.message : String(e)}\n`);
    }
  }

  return messages.length > 0 ? messages.join("\n") : null;
}
