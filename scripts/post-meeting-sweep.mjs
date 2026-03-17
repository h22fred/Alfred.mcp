#!/usr/bin/env node
/**
 * Standalone post-meeting sweep runner — called by the Friday 2pm cron job.
 *
 * Behaviour:
 *  1. Auto-launches ChromeLink if not already running
 *  2. Scans this week's ended online meetings (Mon → now)
 *  3. If auth fails → posts a Teams reminder + macOS notification
 *  4. Posts a summary Adaptive Card to Teams listing meetings that may need engagements
 *
 * Usage: node scripts/post-meeting-sweep.mjs
 * Cron:  0 14 * * 5  (every Friday at 2:00pm)
 */

import { detectPostMeetingEngagements } from "../dist/tools/postMeetingClient.js";
import { fetchEngagementsByOpportunity } from "../dist/tools/dynamicsClient.js";
import { setTeamsWebhook } from "../dist/tools/teamsClient.js";
import { ensureChromeLink } from "../dist/auth/tokenExtractor.js";
import { readFileSync, existsSync } from "fs";
import { homedir } from "os";
import { join } from "path";
import { execFileSync } from "child_process";

const log = (msg) => console.log(`[post-meeting] ${msg}`);
const err = (msg) => console.error(`[post-meeting] ${msg}`);

// ---------------------------------------------------------------------------
// Load config
// ---------------------------------------------------------------------------
const configPath = join(homedir(), ".sc-engagement-config.json");
if (!existsSync(configPath)) {
  err("No config at ~/.sc-engagement-config.json — run Setup.command first.");
  process.exit(1);
}
const config = JSON.parse(readFileSync(configPath, "utf8"));

if (config.teamsWebhook) setTeamsWebhook(config.teamsWebhook);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function postAdaptiveCardRaw(webhookUrl, card) {
  try {
    await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "message",
        attachments: [{
          contentType: "application/vnd.microsoft.card.adaptive",
          content: card,
        }],
      }),
    });
  } catch { /* webhook failure is non-fatal */ }
}

async function postTeamsSimple(webhookUrl, title, body) {
  await postAdaptiveCardRaw(webhookUrl, {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      { type: "TextBlock", text: title, weight: "Bolder", size: "Large", wrap: true },
      { type: "TextBlock", text: body, wrap: true, spacing: "Medium" },
    ],
  });
}

function macosNotify(title, message) {
  try {
    execFileSync("osascript", ["-e",
      `display notification "${message}" with title "${title}" sound name "Ping"`
    ]);
  } catch { /* non-fatal */ }
}

// ISO date string for a given day offset from now
function isoDate(offsetDays = 0) {
  const d = new Date();
  d.setDate(d.getDate() + offsetDays);
  return d.toISOString().slice(0, 10);
}

// Monday of the current week
function mondayOfWeek() {
  const d = new Date();
  const day = d.getDay(); // 0=Sun, 1=Mon...
  const diff = day === 0 ? -6 : 1 - day; // days back to Monday
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0, 10);
}

function fmtDate(isoStr) {
  if (!isoStr) return "—";
  return new Date(isoStr).toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short" });
}

function fmtTime(isoStr) {
  if (!isoStr) return "";
  return new Date(isoStr).toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" });
}

function truncate(s, max) {
  return s && s.length > max ? s.slice(0, max - 1) + "…" : (s ?? "");
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

log("Starting post-meeting sweep...");

// Step 1 — ensure ChromeLink is running
try {
  log("Checking ChromeLink...");
  await ensureChromeLink(log);
  log("ChromeLink is running");
} catch (e) {
  err(`Could not launch ChromeLink: ${e.message}`);
  const msg = "ChromeLink could not start. Open ChromeLink.app, log in, then ask Claude: **\"Detect post-meeting engagements from this week\"**";
  if (config.teamsWebhook) await postTeamsSimple(config.teamsWebhook, "⚠️ Friday Meeting Review — Action Required", msg);
  macosNotify("Meeting Review", "ChromeLink failed to start — run manually");
  process.exit(1);
}

// Step 2 — scan this week's meetings
const startDate = mondayOfWeek();
const endDate   = isoDate(0); // today

log(`Scanning meetings ${startDate} → ${endDate}...`);

let candidates;
try {
  candidates = await detectPostMeetingEngagements({
    hoursBack: 7 * 24, // full week
  }, log);
} catch (e) {
  err(`Sweep failed: ${e.message}`);

  const isAuthError = e.message.includes("cookie") || e.message.includes("auth") ||
                      e.message.includes("401") || e.message.includes("logged in") ||
                      e.message.includes("ChromeLink") || e.message.includes("Graph token");

  const title = "⚠️ Friday Meeting Review — Login Required";
  const body = isAuthError
    ? "Session expired. Open ChromeLink.app, log into Outlook and Dynamics, then ask Claude: **\"Detect post-meeting engagements from this week\"**"
    : `Meeting sweep failed: ${e.message}. Run manually via Claude Desktop.`;

  if (config.teamsWebhook) await postTeamsSimple(config.teamsWebhook, title, body);
  macosNotify("Meeting Review", isAuthError ? "Login required — open ChromeLink" : "Sweep failed — check Teams");
  process.exit(1);
}

if (candidates.length === 0) {
  log("No ended online meetings found this week — nothing to post.");
  macosNotify("Meeting Review", "No customer meetings found this week");
  process.exit(0);
}

log(`Found ${candidates.length} meeting candidate(s) — checking engagement hygiene...`);

// Step 3 — check hygiene for matched opps
const SC_REQUIRED = ["Discovery", "Demo", "Technical Win"];

const hygieneByOpp = new Map(); // oppId → { missing: string[], status: "red"|"yellow"|"green" }
for (const c of candidates) {
  if (!c.suggestedOpportunityId || hygieneByOpp.has(c.suggestedOpportunityId)) continue;
  try {
    const engagements = await fetchEngagementsByOpportunity(c.suggestedOpportunityId, log);
    const typeNames = engagements.map(e => e.engagementTypeName ?? "").filter(Boolean);
    const missing = SC_REQUIRED.filter(t => !typeNames.includes(t));
    const status = missing.length > 0 ? "red" : "green";
    hygieneByOpp.set(c.suggestedOpportunityId, { missing, status });
  } catch {
    // non-fatal — hygiene check fails gracefully
  }
}

log("Building Teams card...");

// Step 4 — build Adaptive Card
const today = new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
const withTranscript = candidates.filter(c => c.transcriptAvailable).length;
const withMatch      = candidates.filter(c => c.suggestedOpportunityName).length;

const cardBody = [
  {
    type: "TextBlock",
    text: `📅 Friday Meeting Review — ${today}`,
    weight: "Bolder", size: "Large", wrap: true,
  },
  {
    type: "ColumnSet", spacing: "Small",
    columns: [
      { type: "Column", width: "auto", items: [{ type: "TextBlock", text: `**${candidates.length}** meetings`, size: "Small" }] },
      { type: "Column", width: "auto", items: [{ type: "TextBlock", text: `**${withTranscript}** with transcript`, size: "Small" }] },
      { type: "Column", width: "auto", items: [{ type: "TextBlock", text: `**${withMatch}** opp matched`, size: "Small" }] },
    ],
  },
  { type: "Separator" },
  // Column headers
  {
    type: "ColumnSet", spacing: "Small",
    columns: [
      { type: "Column", width: "stretch", items: [{ type: "TextBlock", text: "MEETING", size: "Small", weight: "Bolder", isSubtle: true }] },
      { type: "Column", width: "auto",    items: [{ type: "TextBlock", text: "DAY / TIME", size: "Small", weight: "Bolder", isSubtle: true, horizontalAlignment: "Right" }] },
      { type: "Column", width: "auto",    items: [{ type: "TextBlock", text: "MIN", size: "Small", weight: "Bolder", isSubtle: true, horizontalAlignment: "Right" }] },
      { type: "Column", width: "140px",   items: [{ type: "TextBlock", text: "OPPORTUNITY", size: "Small", weight: "Bolder", isSubtle: true, horizontalAlignment: "Right" }] },
      { type: "Column", width: "120px",   items: [{ type: "TextBlock", text: "MISSING", size: "Small", weight: "Bolder", isSubtle: true, horizontalAlignment: "Right" }] },
    ],
  },
];

for (const c of candidates) {
  const txIcon  = c.transcriptAvailable ? "📝" : "🎙";
  const dayTime = `${fmtDate(c.meetingStart)} ${fmtTime(c.meetingStart)}`;
  const duration = c.durationMinutes ? String(c.durationMinutes) : "—";
  const oppText  = c.suggestedAccountName ? truncate(c.suggestedAccountName, 22) : "—";

  const hygiene = c.suggestedOpportunityId ? hygieneByOpp.get(c.suggestedOpportunityId) : null;
  const missingText = hygiene
    ? (hygiene.missing.length ? hygiene.missing.join(" · ") : "✓ complete")
    : "—";
  const missingColor = hygiene?.status === "red" ? "Attention" : hygiene ? "Good" : "Default";

  cardBody.push({
    type: "ColumnSet", spacing: "Small",
    columns: [
      {
        type: "Column", width: "stretch",
        items: [{ type: "TextBlock", text: `${txIcon}  ${truncate(c.meetingSubject, 38)}`, size: "Small", wrap: false }],
      },
      {
        type: "Column", width: "auto",
        items: [{ type: "TextBlock", text: dayTime, size: "Small", isSubtle: true, horizontalAlignment: "Right" }],
      },
      {
        type: "Column", width: "auto",
        items: [{ type: "TextBlock", text: duration, size: "Small", isSubtle: true, horizontalAlignment: "Right" }],
      },
      {
        type: "Column", width: "140px",
        items: [{ type: "TextBlock", text: oppText, size: "Small",
          color: c.suggestedAccountName ? "Accent" : "Default",
          wrap: false, horizontalAlignment: "Right" }],
      },
      {
        type: "Column", width: "120px",
        items: [{ type: "TextBlock", text: missingText, size: "Small",
          color: missingColor, wrap: false, horizontalAlignment: "Right" }],
      },
    ],
  });
}

cardBody.push(
  { type: "Separator", spacing: "Medium" },
  {
    type: "TextBlock",
    text: "Open Claude Desktop and ask: **\"Detect post-meeting engagements from this week\"** to review and log these.",
    wrap: true, size: "Small", isSubtle: true,
  }
);

const card = {
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  type: "AdaptiveCard",
  version: "1.4",
  body: cardBody,
};

if (config.teamsWebhook) {
  await postAdaptiveCardRaw(config.teamsWebhook, card);
  log("✅ Posted to Teams");
} else {
  log("No Teams webhook configured — skipping post");
}

// macOS notification
macosNotify("Meeting Review", `${candidates.length} meetings this week — check Teams`);
log("Post-meeting sweep complete ✅");
