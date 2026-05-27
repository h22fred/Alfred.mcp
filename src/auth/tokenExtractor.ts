import { chromium, type BrowserContext } from "playwright";
import { existsSync, readdirSync, readFileSync, writeFileSync } from "fs";
import { execFileSync } from "child_process";
import { homedir } from "os";
import { join } from "path";
import { DYNAMICS_HOST } from "../config.js";
import { loadCachedAuth, saveCachedAuth, clearCachedAuthFile } from "./authFileCache.js";

const DYNAMICS_URL = DYNAMICS_HOST;
const OUTLOOK_URLS = ["https://outlook.cloud.microsoft", "https://outlook.cloud.microsoft.com", "https://outlook.office.com", "https://outlook.office365.com"];
const COOKIE_REFRESH_MARGIN_MS = 5 * 60 * 1000;

const PROFILE_DIR = `${process.env.HOME}/.alfred-pw`;

// Cookie names that authenticate Dynamics requests
const AUTH_COOKIE_NAMES = ["CrmOwinAuthC1", "CrmOwinAuthC2", "CrmOwinAuth"];

interface CachedAuth {
  cookieHeader: string; // "Name1=val1; Name2=val2"
  expiresAt: number;
}

let cachedAuth: CachedAuth | null = null;
let cachedOutlookAuth: CachedAuth | null = null;

// Dedup concurrent auth refreshes — only one inflight request at a time
let inflightDynamics: Promise<string> | null = null;
let inflightOutlook: Promise<string> | null = null;

export type ProgressFn = (msg: string) => void;

// Module-level Playwright context
let _context: BrowserContext | null = null;

// Auto-close the browser after a short idle period (session persists via ~/.alfred-pw profile).
// Short window (10s) because file cache covers auth for hours — browser only needed during extraction.
const IDLE_CLOSE_MS = 10_000;
let _idleTimer: ReturnType<typeof setTimeout> | null = null;

export function scheduleIdleClose(ms = IDLE_CLOSE_MS): void {
  if (_idleTimer) clearTimeout(_idleTimer);
  _idleTimer = setTimeout(() => {
    _idleTimer = null;
    if (_context) {
      process.stderr.write("[alfred:info] Auto-closing Alfred browser (idle after auth)\n");
      _context.close().catch(() => {});
    }
  }, ms);
}

export function cancelIdleClose(): void {
  if (_idleTimer) { clearTimeout(_idleTimer); _idleTimer = null; }
}

// Patch Chromium app display name to "Alfred" via Info.plist (macOS only, non-fatal)
function patchChromiumName(): void {
  if (process.platform !== "darwin") return;
  try {
    const pwCache = join(homedir(), "Library", "Caches", "ms-playwright");
    if (!existsSync(pwCache)) return;
    const dirs = readdirSync(pwCache).filter(d => d.startsWith("chromium-"));
    for (const dir of dirs) {
      const plist = join(pwCache, dir, "chrome-mac", "Chromium.app", "Contents", "Info.plist");
      let content: string;
      try { content = readFileSync(plist, "utf8"); } catch { continue; }
      if (content.includes("<string>Alfred</string>")) break; // already patched
      content = content
        .replace(/(<key>CFBundleName<\/key>\s*<string>)[^<]*(<\/string>)/, "$1Alfred$2")
        .replace(/(<key>CFBundleDisplayName<\/key>\s*<string>)[^<]*(<\/string>)/, "$1Alfred$2");
      writeFileSync(plist, content, "utf8");
      // Force Dock to re-read the bundle — without this the old name stays cached
      const appBundle = join(pwCache, dir, "chrome-mac", "Chromium.app");
      try {
        execFileSync("/System/Library/Frameworks/CoreServices.framework/Versions/A/Frameworks/LaunchServices.framework/Versions/A/Support/lsregister", ["-f", appBundle], { stdio: "ignore" });
      } catch { /* non-fatal */ }
      break;
    }
  } catch { /* non-fatal */ }
}

/** Clear in-memory auth state only (file cache survives for cross-restart resilience). */
export function clearMemoryAuthCache(): void {
  cachedAuth = null;
  cachedOutlookAuth = null;
  inflightDynamics = null;
  inflightOutlook = null;
}

/** Full wipe — in-memory + file cache. Only call on confirmed 401/403. */
export function clearAuthCache(): void {
  clearMemoryAuthCache();
  clearCachedAuthFile("dynamics");
  clearCachedAuthFile("outlook");
}

/** Health check — tries to call cookies on context; nulls it out on any error. */
async function isContextAlive(): Promise<boolean> {
  if (!_context) return false;
  try {
    await _context.cookies([]);
    scheduleIdleClose(); // reset idle timer on every use
    return true;
  } catch {
    _context = null;
    cancelIdleClose();
    clearMemoryAuthCache();
    return false;
  }
}

/** Fast async health check — checks whether the Playwright context is alive. */
export async function isAlfredgable(): Promise<boolean> {
  return isContextAlive();
}

/** Launch a persistent Playwright context and register lifecycle handlers. */
async function launchContext(): Promise<BrowserContext> {
  patchChromiumName();
  const ctx = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    args: ["--no-first-run", "--no-default-browser-check", "--disable-features=mDnsResponder"],
  });
  ctx.on("close", () => {
    cancelIdleClose();
    _context = null;
    clearMemoryAuthCache();
  });
  return ctx;
}

/** Open default tabs: reuse page[0] for Dynamics, create 2 more for Outlook + Teams. */
async function openDefaultTabs(ctx: BrowserContext): Promise<void> {
  const pages = ctx.pages();
  const dynPage = pages[0] ?? await ctx.newPage();
  if (!dynPage.url().startsWith(DYNAMICS_URL)) {
    await dynPage.goto(DYNAMICS_URL, { waitUntil: "domcontentloaded", timeout: 30_000 }).catch((e) => {
      process.stderr.write(`[alfred:warn] failed to navigate to Dynamics: ${e instanceof Error ? e.message : String(e)}\n`);
    });
  }
  for (const url of [OUTLOOK_URLS[0]!, "https://teams.microsoft.com/v2/"]) {
    await ctx.newPage().then(p => p.goto(url, { waitUntil: "domcontentloaded", timeout: 30_000 })).catch((e) => {
      process.stderr.write(`[alfred:warn] failed to open tab ${url}: ${e instanceof Error ? e.message : String(e)}\n`);
    });
  }
}

/**
 * Returns the live Playwright BrowserContext.
 * Throws if Alfred is not running.
 */
export async function getAlfredContext(): Promise<BrowserContext> {
  if (_context && await isContextAlive()) return _context;
  throw new Error("Alfred browser is not running. It launches automatically — restart Claude Desktop to trigger it.");
}

/** Returns all open pages from the live context. Throws if not running. */
export async function getAlfredPages() {
  const ctx = await getAlfredContext();
  return ctx.pages();
}

/** Stop the Alfred browser gracefully. */
export async function exitAlfred(progress: ProgressFn = () => {}): Promise<boolean> {
  cancelIdleClose();
  if (!await isAlfredgable()) {
    progress("ℹ️ Alfred is not running");
    return false;
  }
  progress("🛑 Closing Alfred...");
  try {
    await _context!.close();
  } catch { /* already closed */ }

  clearAuthCache();
  progress("✅ Alfred closed — all auth tokens cleared");
  return true;
}

/** Restart Alfred: exit, then relaunch. */
export async function restartAlfred(progress: ProgressFn = () => {}): Promise<void> {
  await exitAlfred(progress);
  await new Promise(r => setTimeout(r, 1_500));
  progress("🚀 Restarting Alfred...");
  _context = await launchContext();
  await openDefaultTabs(_context);
  progress("✅ Alfred restarted — please log into Dynamics, Outlook and Teams if needed");
}

export async function ensureAlfred(progress: ProgressFn = () => {}): Promise<void> {
  if (await isAlfredgable()) {
    await verifySessionHealth(progress);
    return;
  }
  clearMemoryAuthCache();
  progress("🚀 Launching Alfred automatically...");
  _context = await launchContext();
  await openDefaultTabs(_context);
  progress("✅ Alfred ready — please log into Dynamics, Outlook and Teams in the new window");
}

/** Probe Dynamics and Outlook cookies to warn if sessions are expired or missing. */
async function verifySessionHealth(progress: ProgressFn): Promise<void> {
  try {
    const ctx = await getAlfredContext();
    const cookies = await ctx.cookies([DYNAMICS_URL, ...OUTLOOK_URLS]);
    const hasDynamics = cookies.some(c => AUTH_COOKIE_NAMES.includes(c.name));
    const hasOutlook = cookies.some(c => c.domain?.includes("outlook") || c.domain?.includes("office"));

    const warnings: string[] = [];
    if (!hasDynamics) warnings.push("Dynamics (not logged in)");
    if (!hasOutlook)  warnings.push("Outlook (not logged in)");

    if (warnings.length > 0) {
      progress(`⚠️ Alfred is running but missing sessions: ${warnings.join(", ")}. Please log in.`);
    }
  } catch {
    // Non-fatal
  }
}

export async function getAuthCookies(progress: ProgressFn = () => {}): Promise<string> {
  if (cachedAuth && Date.now() < cachedAuth.expiresAt - COOKIE_REFRESH_MARGIN_MS) {
    const minsLeft = Math.round((cachedAuth.expiresAt - Date.now()) / 60000);
    progress(`🔑 Using cached session (~${minsLeft} min remaining)`);
    return cachedAuth.cookieHeader;
  }

  // Check file cache before hitting Playwright
  const fileCached = loadCachedAuth("dynamics");
  if (fileCached && Date.now() < fileCached.expiresAt - COOKIE_REFRESH_MARGIN_MS) {
    cachedAuth = { cookieHeader: fileCached.value, expiresAt: fileCached.expiresAt };
    const minsLeft = Math.round((fileCached.expiresAt - Date.now()) / 60000);
    progress(`🔑 Using cached session (~${minsLeft} min remaining)`);
    return fileCached.value;
  }

  // Dedup: if another caller is already refreshing, wait for that result
  if (inflightDynamics) return await inflightDynamics;

  const promise = (async () => {
    try {
      progress("🔐 Acquiring Dynamics session cookies...");
      if (!await isAlfredgable()) {
        clearMemoryAuthCache();
        progress("🚀 Launching Alfred automatically...");
        _context = await launchContext();
        await openDefaultTabs(_context);
        progress("✅ Alfred ready — please log into Dynamics in the new window");
      }

      const ctx = await getAlfredContext();
      progress("🍪 Extracting Dynamics session cookies via Playwright...");
      const allCookies = await ctx.cookies([DYNAMICS_URL]);
      let authCookies = allCookies.filter(c => AUTH_COOKIE_NAMES.includes(c.name));

      if (authCookies.length === 0) {
        // SSO may still be in-flight, or session expired and user needs to log in.
        // Cancel idle close and poll until cookies appear (up to 2 min).
        cancelIdleClose();
        progress("🔄 Waiting for Dynamics session — if expired, log into Dynamics in the Alfred browser window...");
        const pages = ctx.pages();
        const dynPage = pages.find(p => p.url().startsWith(DYNAMICS_URL)) ?? pages[0] ?? await ctx.newPage();
        await dynPage.goto(DYNAMICS_URL, { waitUntil: "domcontentloaded", timeout: 30_000 }).catch(() => {});

        const POLL_INTERVAL_MS = 3_000;
        const POLL_TIMEOUT_MS = 120_000;
        const pollStart = Date.now();
        while (authCookies.length === 0 && Date.now() - pollStart < POLL_TIMEOUT_MS) {
          await new Promise(r => setTimeout(r, POLL_INTERVAL_MS));
          const polled = await ctx.cookies([DYNAMICS_URL]);
          authCookies = polled.filter(c => AUTH_COOKIE_NAMES.includes(c.name));
          if (authCookies.length === 0) {
            const elapsed = Math.round((Date.now() - pollStart) / 1000);
            process.stderr.write(`[alfred:auth] waiting for Dynamics cookies (${elapsed}s)...\n`);
          }
        }
      }

      if (authCookies.length === 0) {
        // Timed out waiting — browser stays open for user to complete login
        process.stderr.write("[alfred] Playwright auth: timed out waiting for Dynamics cookies\n");
        throw new Error(
          "Your Dynamics session has expired and login timed out.\n" +
          "The Alfred browser window is open — log into Dynamics there, then retry your request."
        );
      }

      progress(`✅ Found ${authCookies.length} auth cookie(s): ${authCookies.map(c => c.name).join(", ")}`);
      const cookieHeader = authCookies.map(c => `${c.name}=${c.value}`).join("; ");
      // Playwright cookie expires is Unix timestamp in seconds; convert to ms
      const expiresAt = Math.min(
        ...authCookies.map(c => (c.expires ?? 0) > 0 ? c.expires! * 1000 : Date.now() + 8 * 60 * 60 * 1000)
      );
      cachedAuth = { cookieHeader, expiresAt };
      saveCachedAuth("dynamics", cookieHeader, expiresAt);
      const expiresIn = Math.round((expiresAt - Date.now()) / 60000);
      progress(`✅ Session cookies acquired — valid for ~${expiresIn} minutes`);
      // Cookies cached to disk — browser no longer needed; close in 3s
      scheduleIdleClose(3_000);
      return cookieHeader;
    } catch (e) {
      inflightDynamics = null;
      throw e;
    }
  })();

  inflightDynamics = promise;

  try {
    return await promise;
  } finally {
    // lgtm[js/missing-await] — intentional Promise reference comparison for dedup guard
    if (inflightDynamics === promise) inflightDynamics = null;
  }
}

export async function getOutlookCookies(progress: ProgressFn = () => {}): Promise<string> {
  if (cachedOutlookAuth && Date.now() < cachedOutlookAuth.expiresAt - COOKIE_REFRESH_MARGIN_MS) {
    const minsLeft = Math.round((cachedOutlookAuth.expiresAt - Date.now()) / 60000);
    progress(`🔑 Using cached Outlook session (~${minsLeft} min remaining)`);
    return cachedOutlookAuth.cookieHeader;
  }

  // Check file cache before hitting Playwright
  const fileCached = loadCachedAuth("outlook");
  if (fileCached && Date.now() < fileCached.expiresAt - COOKIE_REFRESH_MARGIN_MS) {
    cachedOutlookAuth = { cookieHeader: fileCached.value, expiresAt: fileCached.expiresAt };
    const minsLeft = Math.round((fileCached.expiresAt - Date.now()) / 60000);
    progress(`🔑 Using cached Outlook session (~${minsLeft} min remaining)`);
    return fileCached.value;
  }

  // Dedup: if another caller is already refreshing, wait for that result
  if (inflightOutlook) return await inflightOutlook;

  const promise = (async () => {
    try {
      progress("🔐 Extracting Outlook cookies via Playwright...");

      if (!await isAlfredgable()) {
        throw new Error("Alfred browser not available — it should launch automatically. Try restarting Claude Desktop.");
      }

      const ctx = await getAlfredContext();
      const allCookies = await ctx.cookies(OUTLOOK_URLS);
      if (allCookies.length === 0) {
        process.stderr.write("[alfred] Playwright auth: no Outlook cookies found — user not logged in\n");
        throw new Error(
          "Not logged into Outlook in the Alfred window.\n" +
          "Log into Outlook in the Alfred window, then retry."
        );
      }

      const cookieHeader = allCookies.map(c => `${c.name}=${c.value}`).join("; ");
      const expiresAt = Math.min(
        ...allCookies.map(c => (c.expires ?? 0) > 0 ? c.expires! * 1000 : Date.now() + 8 * 60 * 60 * 1000)
      );

      cachedOutlookAuth = { cookieHeader, expiresAt };
      saveCachedAuth("outlook", cookieHeader, expiresAt);
      const minsLeft = Math.round((expiresAt - Date.now()) / 60000);
      progress(`✅ Outlook cookies acquired — valid for ~${minsLeft} minutes`);
      // Cookies cached to disk — browser no longer needed; close in 3s
      scheduleIdleClose(3_000);
      return cookieHeader;
    } catch (e) {
      inflightOutlook = null;
      throw e;
    }
  })();

  inflightOutlook = promise;

  try {
    return await promise;
  } finally {
    // lgtm[js/missing-await] — intentional Promise reference comparison for dedup guard
    if (inflightOutlook === promise) inflightOutlook = null;
  }
}

export function setManualCookies(cookieHeader: string): void {
  const expiresAt = Date.now() + 8 * 60 * 60 * 1000; // assume 8h
  cachedAuth = { cookieHeader, expiresAt };
  saveCachedAuth("dynamics", cookieHeader, expiresAt);
  console.error(`[auth] Manual cookies set`);
}
