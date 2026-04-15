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

/**
 * Regenerate the Alfred.app shell script on macOS.
 * Called by update_alfred so the .app bundle stays in sync with the repo
 * (e.g. update-check logic, Chrome flags, notification text).
 * No-op on Windows or if Alfred.app doesn't exist.
 */
export function regenerateAlfredApp(installDir: string): string | null {
  if (process.platform !== "darwin") return null;

  const { existsSync: ex, readFileSync: rf, writeFileSync: wf, copyFileSync: cf } = require("fs") as typeof import("fs");
  const { homedir: hd } = require("os") as typeof import("os");
  const { join: pj } = require("path") as typeof import("path");

  const home = hd();
  const appScript = pj(home, "Desktop", "Alfred.app", "Contents", "MacOS", "Alfred");
  if (!ex(appScript)) return null;

  const configPath = pj(home, ".alfred-config.json");
  const cfg = ex(configPath) ? JSON.parse(rf(configPath, "utf8")) : {};
  const company = cfg.dynamicsCompany ?? "servicenow";
  if (!/^[a-zA-Z0-9-]+$/.test(company)) {
    throw new Error(
      `[alfred] Invalid dynamicsCompany "${company}" in ~/.alfred-config.json: must contain only alphanumeric characters and hyphens.`
    );
  }
  const dynamicsUrl = `https://${company}.crm.dynamics.com`;

  const script = `#!/bin/bash
notify() { osascript -e "display notification \\"\$1\\" with title \\"Alfred\\"" 2>/dev/null; }

# Already running?
if pgrep -f "alfred-profile" > /dev/null 2>&1; then
  notify "Already running — you're good to use Claude!"
  open -a "Claude" 2>/dev/null || true
  exit 0
fi

CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
if [ ! -x "\$CHROME" ]; then
  osascript -e 'display alert "Alfred" message "Google Chrome not found. Please install Chrome first." as critical' 2>/dev/null
  exit 1
fi

mkdir -p "\$HOME/.alfred-profile"

# Pre-launch notification
PROFILE_SIZE=\$(du -sk "\$HOME/.alfred-profile" 2>/dev/null | cut -f1)
if [ -z "\$PROFILE_SIZE" ] || [ "\$PROFILE_SIZE" -lt 500 ]; then
  notify "First time setup: log into Dynamics, Outlook and Teams in this window. You only do this once!"
else
  notify "Launching — ready for Claude!"
fi

# Open Claude after a short delay (runs in background before exec)
(sleep 2 && open -a "Claude" 2>/dev/null) &

# Background update check — uses git directly, never blocks startup
(
  ALFRED_DIR="\$(cd "\$(dirname "\$0")/../.." && pwd)"
  INSTALLED=\$(git -C "\$ALFRED_DIR" rev-parse --short HEAD 2>/dev/null)
  if [ -z "\$INSTALLED" ]; then exit 0; fi
  git -C "\$ALFRED_DIR" fetch --quiet 2>/dev/null || exit 0
  REMOTE=\$(git -C "\$ALFRED_DIR" rev-parse --short origin/main 2>/dev/null)
  if [ -n "\$REMOTE" ] && [ "\$INSTALLED" != "\$REMOTE" ]; then
    osascript -e "display notification \\"A new version of Alfred is available. Ask Claude: update Alfred\\" with title \\"Alfred Update Available 🆕\\" sound name \\"Ping\\"" 2>/dev/null
  fi
) &

# Become Chrome — Alfred.app's Dock icon persists on the process
exec "\$CHROME" \\
  --remote-debugging-port=9222 \\
  --remote-debugging-address=127.0.0.1 \\
  --user-data-dir="\$HOME/.alfred-profile" \\
  --no-first-run \\
  --no-default-browser-check \\
  --disable-extensions \\
  --disable-sync \\
  --disable-default-apps \\
  --disable-translate \\
  --disable-component-update \\
  --disable-domain-reliability \\
  --disable-client-side-phishing-detection \\
  "${dynamicsUrl}" \\
  "https://outlook.office.com" \\
  "https://teams.microsoft.com/v2/"
`;

  wf(appScript, script, { mode: 0o755 });

  // Refresh icon in case it changed
  const iconSrc = pj(installDir, "setup", "assets", "alfred.icns");
  const iconDst = pj(home, "Desktop", "Alfred.app", "Contents", "Resources", "alfred.icns");
  if (ex(iconSrc)) cf(iconSrc, iconDst);

  // Regenerate Info.plist (removes LSUIElement so Alfred shows in Dock with its icon)
  const plistPath = pj(home, "Desktop", "Alfred.app", "Contents", "Info.plist");
  wf(plistPath, `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleExecutable</key><string>Alfred</string>
  <key>CFBundleIdentifier</key><string>com.servicenow.alfred</string>
  <key>CFBundleName</key><string>Alfred</string>
  <key>CFBundleIconFile</key><string>alfred</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleVersion</key><string>1.0</string>
  <key>NSHighResolutionCapable</key><true/>
</dict>
</plist>
`);

  return "🔄 Regenerated Alfred.app with latest updates";
}
