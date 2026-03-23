#!/usr/bin/env node
/**
 * Standalone hygiene sweep runner — called by the Monday 9:30am cron job.
 *
 * Behaviour:
 *  1. Auto-launches Alfred if not already running
 *  2. Tries to run the hygiene sweep using your existing Dynamics session
 *  3. If auth fails (not logged in) → posts a Teams reminder + macOS notification
 *  4. Posts results to Teams on success
 *
 * Usage: node scripts/hygiene-sweep.mjs
 * Cron:  30 9 * * 1  (every Monday at 9:30am)
 */

import { runHygieneSweep, formatHygieneReport } from "../dist/tools/hygieneClient.js";
import { setTeamsWebhook } from "../dist/tools/teamsClient.js";
import { ensureAlfred } from "../dist/auth/tokenExtractor.js";
import { readFileSync, existsSync } from "fs";
import { homedir } from "os";
import { join } from "path";
import { execFileSync } from "child_process";

const log = (msg) => console.log(`[hygiene] ${msg}`);
const err = (msg) => console.error(`[hygiene] ${msg}`);

// ---------------------------------------------------------------------------
// Load config
// ---------------------------------------------------------------------------
const configPath = join(homedir(), ".alfred-config.json");
if (!existsSync(configPath)) {
  err("No config at ~/.alfred-config.json — run Setup.command first.");
  process.exit(1);
}
const config = JSON.parse(readFileSync(configPath, "utf8"));

if (config.teamsWebhook) {
  setTeamsWebhook(config.teamsWebhook);
} else {
  log("⚠️  No Teams webhook in config — results will only appear in the log. Re-run Setup.command to add one.");
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function postTeamsRaw(webhookUrl, title, body) {
  try {
    await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "message",
        attachments: [{
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
            type: "AdaptiveCard",
            version: "1.4",
            body: [
              { type: "TextBlock", text: title, weight: "Bolder", size: "Large", wrap: true },
              { type: "TextBlock", text: body, wrap: true, spacing: "Medium" },
            ],
          },
        }],
      }),
    });
  } catch { /* webhook failure is non-fatal */ }
}

function macosNotify(title, message) {
  try {
    execFileSync("osascript", ["-e",
      `display notification "${message}" with title "${title}" sound name "Ping"`
    ]);
  } catch { /* non-fatal */ }
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

log("Starting hygiene sweep...");

// Step 1 — ensure Alfred is running
try {
  log("Checking Alfred...");
  await ensureAlfred(log);
  log("Alfred is running");
} catch (e) {
  err(`Could not launch Alfred: ${e.message}`);
  const msg = "Alfred could not start. Open Alfred.app, log into Dynamics, then ask Claude to run hygiene sweep.";
  if (config.teamsWebhook) await postTeamsRaw(config.teamsWebhook, "⚠️ Weekly CRM Hygiene — Action Required", msg);
  macosNotify("CRM Hygiene Sweep", "Alfred failed to start — run manually");
  process.exit(1);
}

// Step 2 — run sweep (getAuthCookies will use CDP automatically)
try {
  const results = await runHygieneSweep({
    postToTeams: !!config.teamsWebhook,
    minNnacv: 100_000,
    engagementTypes: config.engagementTypes,
  }, log);

  const report = formatHygieneReport(results);
  console.log("\n" + report);
  log("Sweep complete ✅");

} catch (e) {
  err(`Sweep failed: ${e.message}`);

  const isAuthError = e.message.includes("cookie") || e.message.includes("auth") ||
                      e.message.includes("401") || e.message.includes("logged in");

  const isTeamsError = e.message.includes("Teams rejected") || e.message.includes("HTTP error 400") || e.message.includes("webhook");
  const title = isAuthError ? "⚠️ Weekly CRM Hygiene — Login Required" : "⚠️ Weekly CRM Hygiene — Failed";
  const body = isAuthError
    ? "Dynamics session has expired. Open Alfred.app, log back into Dynamics, then ask Claude: **\"Run hygiene sweep and post to Teams\"**"
    : isTeamsError
      ? `Teams rejected the hygiene card: ${e.message}`
      : `Hygiene sweep failed: ${e.message}. Please run manually via Claude Desktop.`;

  if (config.teamsWebhook) await postTeamsRaw(config.teamsWebhook, title, body);
  macosNotify("CRM Hygiene Sweep", isAuthError ? "Login required — open Alfred" : "Sweep failed — check Teams");
  process.exit(1);
}
