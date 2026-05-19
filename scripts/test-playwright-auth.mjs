#!/usr/bin/env node
/**
 * POC: test Playwright-based cookie extraction for Alfred auth.
 *
 * Run with:   node scripts/test-playwright-auth.mjs
 *
 * What it does:
 *   1. Launches a headed Chromium window (persistent profile at ~/.alfred-pw-test)
 *   2. Opens Dynamics, Outlook and Teams tabs simultaneously
 *   3. Polls until all 3 services have auth cookies (or 3-min timeout)
 *   4. Prints cookie counts + expiry, then closes the browser
 */

import { chromium } from "playwright";
import { readFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const configPath = resolve(__dirname, "../alfred-config.json");

let dynamicsUrl;
try {
  const cfg = JSON.parse(readFileSync(configPath, "utf8"));
  dynamicsUrl = cfg.dynamicsUrl?.replace(/\/?$/, "");
} catch {
  dynamicsUrl = process.env.ALFRED_DYNAMICS_URL ?? "https://servicenow.crm.dynamics.com";
}

const OUTLOOK_URL = "https://outlook.cloud.microsoft";
const TEAMS_URL   = "https://teams.microsoft.com/v2/";

const PROFILE_DIR      = `${process.env.HOME}/.alfred-pw-test`;
const AUTH_COOKIE_NAMES = ["CrmOwinAuthC1", "CrmOwinAuthC2", "CrmOwinAuth"];
const TIMEOUT_MS       = 3 * 60 * 1000;

console.log(`Dynamics : ${dynamicsUrl}`);
console.log(`Outlook  : ${OUTLOOK_URL}`);
console.log(`Teams    : ${TEAMS_URL}`);

const context = await chromium.launchPersistentContext(PROFILE_DIR, {
  headless: false,
  args: [
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-features=mDnsResponder", // suppress macOS "access local network" dialog
  ],
});

// Reuse the blank tab Playwright opens automatically, create 2 more
const existingPages = context.pages();
const dynPage   = existingPages[0] ?? await context.newPage();
const outPage   = await context.newPage();
const teamsPage = await context.newPage();

await Promise.all([
  dynPage.goto(dynamicsUrl),
  outPage.goto(OUTLOOK_URL),
  teamsPage.goto(TEAMS_URL),
]);

console.log("\nAll 3 tabs open — log in if prompted. Waiting up to 3 minutes...\n");

const deadline = Date.now() + TIMEOUT_MS;
let result = null;

while (Date.now() < deadline) {
  const [dynCookies, outCookies, teamsCookies] = await Promise.all([
    context.cookies([dynamicsUrl]),
    context.cookies([OUTLOOK_URL, "https://outlook.office.com", "https://outlook.office365.com"]),
    context.cookies([TEAMS_URL]),
  ]);

  const authCookies  = dynCookies.filter(c => AUTH_COOKIE_NAMES.includes(c.name));
  const outlookOk    = outCookies.length > 0;
  const teamsOk      = teamsCookies.length > 0;

  process.stdout.write(
    `\r  Dynamics: ${authCookies.length ? "✅" : "⏳"}  Outlook: ${outlookOk ? "✅" : "⏳"}  Teams: ${teamsOk ? "✅" : "⏳"}   `
  );

  if (authCookies.length > 0 && outlookOk && teamsOk) {
    result = { authCookies, outCookies, teamsCookies };
    break;
  }

  await new Promise(r => setTimeout(r, 2_000));
}

console.log("\n");

if (result) {
  console.log(`✅ Dynamics — ${result.authCookies.length} auth cookie(s): ${result.authCookies.map(c => c.name).join(", ")}`);
  for (const c of result.authCookies) {
    const exp = c.expires > 0 ? new Date(c.expires * 1000).toISOString() : "session";
    console.log(`   ${c.name}: expires ${exp}`);
  }
  console.log(`✅ Outlook — ${result.outCookies.length} cookie(s)`);
  console.log(`✅ Teams   — ${result.teamsCookies.length} cookie(s)`);
  console.log("\nSession is persistent — next run will skip login entirely.");
} else {
  console.log("❌ TIMEOUT — one or more services didn't authenticate in 3 minutes.");
}

await context.close();
console.log("Browser closed.");
