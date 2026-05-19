import { chromium, type BrowserContext } from "playwright";
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
    return true;
  } catch {
    _context = null;
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
  const ctx = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    args: ["--no-first-run", "--no-default-browser-check", "--disable-features=mDnsResponder"],
  });
  ctx.on("close", () => {
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
  throw new Error("Alfred browser is not running. Launch Alfred from your Desktop first.");
}

/** Returns all open pages from the live context. Throws if not running. */
export async function getAlfredPages() {
  const ctx = await getAlfredContext();
  return ctx.pages();
}

/** Stop the Alfred browser gracefully. */
export async function exitAlfred(progress: ProgressFn = () => {}): Promise<boolean> {
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
  if (inflightDynamics) return inflightDynamics;

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
      const authCookies = allCookies.filter(c => AUTH_COOKIE_NAMES.includes(c.name));
      if (authCookies.length === 0) {
        process.stderr.write("[alfred] Playwright auth: no Dynamics cookies found — user not logged in\n");
        throw new Error(
          "Alfred is open but you are not logged into Dynamics yet.\n" +
          "Please log into Dynamics in the Alfred window, then retry."
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
  if (inflightOutlook) return inflightOutlook;

  const promise = (async () => {
    try {
      progress("🔐 Extracting Outlook cookies via Playwright...");

      if (!await isAlfredgable()) {
        throw new Error("Alfred browser not available. Launch Alfred from your Desktop first.");
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
    if (inflightOutlook === promise) inflightOutlook = null;
  }
}

export function setManualCookies(cookieHeader: string): void {
  const expiresAt = Date.now() + 8 * 60 * 60 * 1000; // assume 8h
  cachedAuth = { cookieHeader, expiresAt };
  saveCachedAuth("dynamics", cookieHeader, expiresAt);
  console.error(`[auth] Manual cookies set`);
}
