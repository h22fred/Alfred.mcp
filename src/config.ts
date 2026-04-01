import { readFileSync, existsSync, statSync } from "fs";
import { homedir } from "os";
import { join } from "path";

interface AlfredConfig {
  dynamicsUrl?: string;
  teamsWebhook?: string;
  role?: "sc" | "ssc" | "manager";
  engagementTypes?: string[];
  installedVersion?: string;
}

const configPath = join(homedir(), ".alfred-config.json");
let raw: AlfredConfig = {};
if (existsSync(configPath)) {
  try {
    raw = JSON.parse(readFileSync(configPath, "utf8")) as AlfredConfig;
  } catch (e) {
    const msg = `[alfred] FATAL: ${configPath} exists but is not valid JSON: ${e instanceof Error ? e.message : String(e)}`;
    process.stderr.write(msg + "\n");
    throw new Error(msg + `\nFix or delete the file and restart Alfred.`);
  }

  // Warn if config file is readable by group or others (Unix/macOS only)
  if (process.platform !== "win32") {
    try {
      const mode = statSync(configPath).mode;
      if (mode & 0o044) {
        process.stderr.write(
          `[alfred] \u26a0\ufe0f WARNING: ~/.alfred-config.json is readable by other users. Run: chmod 600 ~/.alfred-config.json\n`
        );
      }
    } catch (e) {
      process.stderr.write(`[alfred:warn] stat check on config file failed: ${e instanceof Error ? e.message : String(e)}\n`);
    }
  }
}

/** Base URL of the customer's Dynamics 365 instance, e.g. https://acme.crm.dynamics.com */
export const DYNAMICS_HOST: string =
  raw.dynamicsUrl ?? "https://servicenow.crm.dynamics.com";

export const alfredConfig = raw;

// ---------------------------------------------------------------------------
// Canonical engagement type list — single source of truth
// ---------------------------------------------------------------------------

export const ALL_ENGAGEMENT_TYPES = [
  "Business Case",
  "Customer Business Review",
  "Demo",
  "Discovery",
  "EBC",
  "Post Sale Engagement",
  "POV",
  "RFx",
  "Technical Win",
  "Workshop",
] as const;

export type EngagementType = (typeof ALL_ENGAGEMENT_TYPES)[number];

/** Hardcoded GUIDs from sn_engagementtypes lookup table */
export const ENGAGEMENT_TYPE_GUIDS: Record<EngagementType, string> = {
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
