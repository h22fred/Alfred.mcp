#!/usr/bin/env node
/**
 * Standalone post-meeting sweep runner — called by the Friday 2pm cron job.
 *
 * Behaviour:
 *  1. Auto-launches Alfred if not already running
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
import { ensureAlfred } from "../dist/auth/tokenExtractor.js";
import { readFileSync, existsSync, writeFileSync } from "fs";
import { homedir } from "os";
import { join } from "path";
import { execFileSync } from "child_process";

const log = (msg) => console.log(`[post-meeting] ${msg}`);
const err = (msg) => console.error(`[post-meeting] ${msg}`);

// ---------------------------------------------------------------------------
// Load config
// ---------------------------------------------------------------------------
const configPath = join(homedir(), ".alfred-config.json");
if (!existsSync(configPath)) {
  err("No config at ~/.alfred-config.json — run Setup.command first.");
  process.exit(1);
}
const config = JSON.parse(readFileSync(configPath, "utf8"));

if (config.teamsWebhook) setTeamsWebhook(config.teamsWebhook);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function postAdaptiveCardRaw(webhookUrl, card) {
  try {
    const body = JSON.stringify({
      type: "message",
      attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: card }],
    });
    const sizeKb = body.length / 1024;
    log(`📦 Card payload: ${sizeKb.toFixed(1)}KB`);
    if (sizeKb > 27) {
      err(`Card payload too large (${sizeKb.toFixed(1)}KB > 27KB limit) — Teams will likely reject it`);
    }
    const res = await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body,
    });
    const responseText = await res.text().catch(() => "");
    if (!res.ok) {
      err(`Teams webhook error: ${res.status} ${res.statusText}${responseText ? ` — ${responseText}` : ""}`);
      return false;
    }
    // Teams sometimes returns an error message with status 200 (e.g. "1" = success, anything else = warning)
    if (responseText && responseText !== "1") {
      log(`Teams webhook response body: ${responseText}`);
      if (responseText.toLowerCase().includes("failed") || responseText.toLowerCase().includes("error")) {
        err(`Teams rejected the card: ${responseText}`);
        return false;
      }
    }
    return true;
  } catch (e) {
    err(`Teams webhook request failed: ${e.message}`);
    return false;
  }
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
    // Strip characters that could break out of the AppleScript string literal
    const safeTitle   = String(title).replace(/["\n\r]/g, " ").slice(0, 80);
    const safeMessage = String(message).replace(/["\n\r]/g, " ").slice(0, 200);
    execFileSync("osascript", ["-e",
      `display notification "${safeMessage}" with title "${safeTitle}" sound name "Ping"`
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

// Step 1 — ensure Alfred is running
try {
  log("Checking Alfred...");
  await ensureAlfred(log);
  log("Alfred is running");
} catch (e) {
  err(`Could not launch Alfred: ${e.message}`);
  const msg = "Alfred could not start. Open Alfred.app, log in, then ask Claude: **\"Detect post-meeting engagements from this week\"**";
  if (config.teamsWebhook) await postTeamsSimple(config.teamsWebhook, "⚠️ Friday Meeting Review — Action Required", msg);
  macosNotify("Meeting Review", "Alfred failed to start — run manually");
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

  const isCdpError = e.message.includes("connectOverCDP") || e.message.includes("Timeout") ||
                     e.message.includes("Alfred") || e.message.includes("stale");
  const isAuthError = e.message.includes("cookie") || e.message.includes("auth") ||
                      e.message.includes("401") || e.message.includes("logged in") ||
                      e.message.includes("Graph token");

  const title = "⚠️ Friday Meeting Review — Action Required";
  const body = (isCdpError || isAuthError)
    ? "Alfred session needs a refresh. **Close and reopen Alfred.app**, log into Outlook and Dynamics, then ask Claude: **\"Detect post-meeting engagements from this week\"**"
    : `Meeting sweep failed: ${e.message}. Run manually via Claude Desktop.`;

  if (config.teamsWebhook) await postTeamsSimple(config.teamsWebhook, title, body);
  macosNotify("Meeting Review", isAuthError ? "Login required — open Alfred" : "Sweep failed — check Teams");
  process.exit(1);
}

if (candidates.length === 0) {
  log("No ended online meetings found this week — nothing to post.");
  macosNotify("Meeting Review", "No customer meetings found this week");
  process.exit(0);
}

log(`Found ${candidates.length} meeting candidate(s) — checking engagement hygiene...`);

// Step 3 — check hygiene for matched opps
const SC_REQUIRED = config.engagementTypes?.length
  ? config.engagementTypes
  : ["Discovery", "Demo", "Technical Win"];

const hygieneByOpp = new Map(); // oppId → { missing: string[], status: "red"|"yellow"|"green" }
for (const c of candidates) {
  if (!c.suggestedOpportunityId || hygieneByOpp.has(c.suggestedOpportunityId)) continue;
  try {
    const engagements = await fetchEngagementsByOpportunity(c.suggestedOpportunityId, log);
    const typeNames = engagements.map(e => e.engagementTypeName ?? "").filter(Boolean);
    const missing  = SC_REQUIRED.filter(t => !typeNames.includes(t));
    const existing = SC_REQUIRED.filter(t =>  typeNames.includes(t));
    const status = missing.length > 0 ? "red" : "green";
    hygieneByOpp.set(c.suggestedOpportunityId, { missing, existing, status });
  } catch {
    // non-fatal — hygiene check fails gracefully
  }
}

// Cap at 15 most relevant — matched opps first, then by duration desc
const displayCandidates = [...candidates]
  .sort((a, b) => {
    if (a.suggestedOpportunityId && !b.suggestedOpportunityId) return -1;
    if (!a.suggestedOpportunityId && b.suggestedOpportunityId) return 1;
    return (b.durationMinutes ?? 0) - (a.durationMinutes ?? 0);
  })
  .slice(0, 15);

const trimmed = candidates.length > 15;
log(`Building Teams card (${displayCandidates.length}/${candidates.length} meetings)...`);

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
];

for (const c of displayCandidates) {
  const txIcon  = c.transcriptAvailable ? "📝" : "🎙";
  const dayTime = `${fmtDate(c.meetingStart)} ${fmtTime(c.meetingStart)}`;
  const duration = c.durationMinutes ? ` · ${c.durationMinutes}m` : "";
  const subtitle = `${dayTime}${duration}`;

  const hygiene = c.suggestedOpportunityId ? hygieneByOpp.get(c.suggestedOpportunityId) : null;

  const items = [
    // Meeting name — full, wrapping
    { type: "TextBlock", text: `${txIcon} **${c.meetingSubject}**`, wrap: true },
    // Date / duration
    { type: "TextBlock", text: subtitle, size: "Small", isSubtle: true, spacing: "None" },
  ];

  if (!c.suggestedOpportunityId) {
    items.push({
      type: "TextBlock",
      text: "⚠️ No opportunity matched — open Claude Desktop to find it",
      color: "Warning", wrap: true, size: "Small",
    });
  } else {
    // Opportunity name + why it was matched
    const matchDesc = c.matchReason ? ` _(matched by ${c.matchReason})_` : "";
    const oppDisplay = c.suggestedOpportunityName ?? c.suggestedAccountName ?? "Unknown";
    items.push({
      type: "TextBlock",
      text: `🏢 **${oppDisplay}**${matchDesc}`,
      wrap: true, color: "Accent", size: "Small",
    });

    if (hygiene?.missing.length) {
      // Explain what to log and why
      items.push({
        type: "TextBlock",
        text: `📝 **Log new engagement${hygiene.missing.length > 1 ? "s" : ""}:** ${hygiene.missing.join(", ")} — these types are missing on the opportunity and should be created from this meeting`,
        wrap: true, color: "Attention", size: "Small", spacing: "None",
      });
    } else if (hygiene) {
      items.push({
        type: "TextBlock",
        text: "✅ All required engagement types are already logged",
        wrap: true, color: "Good", size: "Small", spacing: "None",
      });
    }

    if (hygiene?.existing.length) {
      // Explain what to update and what to add
      items.push({
        type: "TextBlock",
        text: `✏️ **Update existing engagement${hygiene.existing.length > 1 ? "s" : ""}:** ${hygiene.existing.join(", ")} — open each record and add meeting notes, outcome, and next steps from this meeting`,
        wrap: true, size: "Small", spacing: "None",
      });
    }

    // Direct link to the opportunity in Dynamics
    if (config.dynamicsUrl && c.suggestedOpportunityId) {
      const oppUrl = `${config.dynamicsUrl}/main.aspx?etn=opportunity&pagetype=entityrecord&id=${c.suggestedOpportunityId}`;
      items.push({
        type: "TextBlock",
        text: `[Open opportunity in Dynamics →](${oppUrl})`,
        wrap: true, size: "Small", spacing: "None",
      });
    }
  }

  cardBody.push({
    type: "Container", spacing: "Small", separator: cardBody.length > 2,
    items,
  });
}

const footerText = trimmed
  ? `Showing ${displayCandidates.length} of ${candidates.length} meetings (top by opp match + duration). Open Claude Desktop and ask: **"Detect post-meeting engagements from this week"** to review all of them.`
  : `Open Claude Desktop and ask: **"Detect post-meeting engagements from this week"** to review and log these.`;

cardBody.push(
  { type: "TextBlock", text: footerText, wrap: true, size: "Small", isSubtle: true, separator: true, spacing: "Medium" }
);

const card = {
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  type: "AdaptiveCard",
  version: "1.4",
  body: cardBody,
};

// Dump card to file for debugging — only when ALFRED_DEBUG=1 is set
if (process.env.ALFRED_DEBUG) {
  const debugPath = join(homedir(), ".alfred-meeting-debug.json");
  writeFileSync(debugPath, JSON.stringify({ type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: card }] }, null, 2));
  try { execFileSync("chmod", ["600", debugPath]); } catch { /* non-fatal */ }
  log(`Card JSON saved to ${debugPath} for inspection`);
}

if (config.teamsWebhook) {
  const ok = await postAdaptiveCardRaw(config.teamsWebhook, card);
  if (ok) log("✅ Posted to Teams");
  else err("❌ Teams card failed to post — check payload size or webhook URL");
} else {
  log("No Teams webhook configured — skipping post");
}

// macOS notification
macosNotify("Meeting Review", `${candidates.length} meetings this week — check Teams`);
log("Post-meeting sweep complete ✅");
