import { execFileSync, execFile } from "child_process";
import { existsSync } from "fs";
import { DYNAMICS_HOST } from "../config.js";
import { loadCachedAuth, saveCachedAuth, clearCachedAuthFile } from "./authFileCache.js";

const DYNAMICS_URL = DYNAMICS_HOST;
const OUTLOOK_URL  = "https://outlook.office.com";
const CDP_PORT = 9222;
const COOKIE_REFRESH_MARGIN_MS = 5 * 60 * 1000;

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

const CHROME_PROFILE_DIR = `${process.env.HOME}/.alfred-profile`;

function isChromeProcessRunning(): boolean {
  try {
    execFileSync("pgrep", ["-f", CHROME_PROFILE_DIR], { timeout: 2_000 });
    return true;
  } catch {
    return false;
  }
}

function isAlfredgable(): boolean {
  try {
    execFileSync("curl", ["-s", "--max-time", "3", `http://localhost:${CDP_PORT}/json/version`], { timeout: 5_000 });
    return true;
  } catch {
    return false;
  }
}

function launchAlfred(): void {
  if (isChromeProcessRunning()) {
    console.error("[auth] Alfred process already running — waiting for port to become ready...");
    return;
  }
  console.error("[auth] Launching Alfred...");
  execFile("/bin/sh", ["-c",
    `mkdir -p "${CHROME_PROFILE_DIR}" && \
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
      --remote-debugging-port=${CDP_PORT} \
      --remote-debugging-address=127.0.0.1 \
      --user-data-dir="${CHROME_PROFILE_DIR}" \
      --no-first-run \
      --no-default-browser-check \
      > /dev/null 2>&1 &`
  ]);
}

async function waitForChrome(timeoutMs = 15_000): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (isAlfredgable()) return;
    await new Promise(r => setTimeout(r, 500));
  }
  throw new Error("Alfred browser did not start in time. Double-click Alfred on your Desktop to launch manually.");
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

/** Stop the Alfred Chrome process gracefully via CDP Browser.close — cross-platform. */
export async function exitAlfred(progress: ProgressFn = () => {}): Promise<boolean> {
  if (!isAlfredgable()) {
    progress("ℹ️ Alfred is not running");
    return false;
  }
  progress("🛑 Closing Alfred (only the Alfred Chrome — your regular Chrome is untouched)...");
  try {
    // Use CDP Browser.close — works on macOS and Windows, targets only the debug-port Chrome
    const verRes = await fetch(`http://localhost:${CDP_PORT}/json/version`);
    const verInfo = await verRes.json() as { webSocketDebuggerUrl?: string };
    const wsUrl = verInfo.webSocketDebuggerUrl;
    if (wsUrl) {
      await new Promise<void>((resolve) => {
        const ws = new WebSocket(wsUrl);
        ws.addEventListener("open", () => {
          ws.send(JSON.stringify({ id: 1, method: "Browser.close" }));
          // Don't wait for response — Chrome closes immediately
          setTimeout(resolve, 500);
        });
        ws.addEventListener("error", () => resolve());
        setTimeout(resolve, 3_000); // safety timeout
      });
    }
  } catch { /* Chrome may have already closed */ }

  // Fallback: if CDP close didn't work, try platform-specific kill
  await new Promise(r => setTimeout(r, 500));
  if (isAlfredgable()) {
    try {
      if (process.platform === "win32") {
        execFileSync("taskkill", ["/F", "/IM", "chrome.exe", "/FI", `WINDOWTITLE eq *alfred*`], { timeout: 5_000 });
      } else {
        execFileSync("pkill", ["-f", CHROME_PROFILE_DIR], { timeout: 5_000 });
      }
    } catch { /* already gone */ }
  }

  clearAuthCache();
  progress("✅ Alfred closed — all auth tokens cleared");
  return true;
}

/** Restart Alfred: exit, then relaunch via Desktop shortcut (preserves Dock icon on macOS). */
export async function restartAlfred(progress: ProgressFn = () => {}): Promise<void> {
  await exitAlfred(progress);
  // Wait for the process to fully terminate
  await new Promise(r => setTimeout(r, 1_500));
  progress("🚀 Restarting Alfred...");

  // Relaunch via Desktop shortcut (preserves icon) before falling back to direct launch
  const home = process.env.HOME ?? process.env.USERPROFILE ?? "";
  const alfredApp = `${home}/Desktop/Alfred.app`;
  const alfredBat = `${home}\\Desktop\\Alfred.bat`;

  if (process.platform === "darwin" && existsSync(alfredApp)) {
    execFile("open", [alfredApp]);
  } else if (process.platform === "win32" && existsSync(alfredBat)) {
    execFile("cmd", ["/c", "start", "", alfredBat]);
  } else {
    launchAlfred();
  }

  await waitForChrome(20_000);
  // Open standard tabs
  for (const url of [DYNAMICS_URL, OUTLOOK_URL, "https://teams.microsoft.com/v2/"]) {
    await fetch(`http://localhost:${CDP_PORT}/json/new?${url}`).catch((e) => { process.stderr.write(`[alfred:warn] failed to open tab ${url}: ${e instanceof Error ? e.message : String(e)}\n`); });
  }
  progress("✅ Alfred restarted — please log into Dynamics, Outlook and Teams if needed");
}

export async function ensureAlfred(progress: ProgressFn = () => {}): Promise<void> {
  if (isAlfredgable()) return;
  // Chrome is not running — clear in-memory state only; file cache survives
  // so persisted tokens can still be used if they haven't expired.
  clearMemoryAuthCache();
  progress("🚀 Launching Alfred automatically...");
  launchAlfred();
  await waitForChrome();
  // Open Dynamics, Outlook and Teams tabs via CDP
  for (const url of [DYNAMICS_URL, OUTLOOK_URL, "https://teams.microsoft.com/v2/"]) {
    await fetch(`http://localhost:${CDP_PORT}/json/new?${url}`).catch((e) => { process.stderr.write(`[alfred:warn] failed to open tab ${url}: ${e instanceof Error ? e.message : String(e)}\n`); });
  }
  progress("✅ Alfred ready — please log into Dynamics, Outlook and Teams in the new window");
}

export async function freshCdpEndpoint(): Promise<string> {
  // Always use the browser-level WebSocket URL (re-fetched every call — never stale)
  const res = await fetch(`http://localhost:${CDP_PORT}/json/version`);
  const info = await res.json() as { webSocketDebuggerUrl?: string };
  if (info.webSocketDebuggerUrl) return info.webSocketDebuggerUrl;
  throw new Error("Could not resolve CDP WebSocket URL from Alfred.");
}

export async function connectWithRetry(retries = 3) {
  const { chromium } = await import("playwright");
  let lastError: Error = new Error("Unknown error");
  for (let i = 0; i < retries; i++) {
    try {
      const wsUrl = await freshCdpEndpoint();
      return await chromium.connectOverCDP(wsUrl, { timeout: 10_000 });
    } catch (e) {
      lastError = e as Error;
      if (i < retries - 1) await new Promise(r => setTimeout(r, 1_000));
    }
  }
  throw new Error(
    "Could not connect to Alfred. Please close and reopen Alfred from your Desktop, then retry.\n" +
    `(${lastError.message})`
  );
}

// ---------------------------------------------------------------------------
// Raw CDP WebSocket helper — avoids Playwright entirely (no Browser.close risk)
// ---------------------------------------------------------------------------

interface RawCookie { name: string; value: string; expires: number; domain: string; }

async function getCookiesViaRawCDP(urls: string[]): Promise<RawCookie[]> {
  // Use any available page target — Network.getCookies works across all domains
  const listRes = await fetch(`http://localhost:${CDP_PORT}/json/list`);
  const targets = await listRes.json() as Array<{ webSocketDebuggerUrl?: string; type?: string }>;
  const target = targets.find(t => t.type === "page" && t.webSocketDebuggerUrl);
  if (!target?.webSocketDebuggerUrl) {
    throw new Error("No browser tabs found. Make sure the Alfred browser is running.");
  }

  return new Promise((resolve, reject) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    const timer = setTimeout(() => {
      try { ws.close(); } catch { /* ignore */ }
      reject(new Error("CDP cookie fetch timed out"));
    }, 8_000);

    ws.addEventListener("open", () => {
      ws.send(JSON.stringify({ id: 1, method: "Network.getCookies", params: { urls } }));
    });

    ws.addEventListener("message", (event: MessageEvent) => {
      clearTimeout(timer);
      try { ws.close(); } catch { /* ignore */ }
      try {
        const msg = JSON.parse(event.data as string) as { id: number; result?: { cookies: RawCookie[] }; error?: { message: string } };
        if (msg.error) reject(new Error(`CDP: ${msg.error.message}`));
        else resolve(msg.result?.cookies ?? []);
      } catch (e) { reject(e); }
    });

    ws.addEventListener("error", () => {
      clearTimeout(timer);
      reject(new Error("CDP WebSocket error — is Alfred running?"));
    });
  });
}

async function getAuthCookiesViaCDP(progress: ProgressFn): Promise<CachedAuth> {
  progress("🍪 Extracting Dynamics session cookies via CDP...");
  const allCookies = await getCookiesViaRawCDP([DYNAMICS_URL]);

  const authCookies = allCookies.filter(c => AUTH_COOKIE_NAMES.includes(c.name));
  if (authCookies.length === 0) {
    process.stderr.write("[alfred] CDP auth: no Dynamics cookies found — user not logged in\n");
    throw new Error(
      "Alfred is open but you are not logged into Dynamics yet.\n" +
      "Please log into Dynamics in the Alfred window, then retry."
    );
  }

  progress(`✅ Found ${authCookies.length} auth cookie(s): ${authCookies.map(c => c.name).join(", ")}`);

  const cookieHeader = authCookies.map(c => `${c.name}=${c.value}`).join("; ");
  // CDP Network.getCookies returns expires as Unix timestamp in seconds; convert to ms
  const expiresAt = Math.min(
    ...authCookies.map(c => c.expires > 0 ? c.expires * 1000 : Date.now() + 8 * 60 * 60 * 1000)
  );
  return { cookieHeader, expiresAt };
}

export async function getAuthCookies(progress: ProgressFn = () => {}): Promise<string> {
  if (cachedAuth && Date.now() < cachedAuth.expiresAt - COOKIE_REFRESH_MARGIN_MS) {
    const minsLeft = Math.round((cachedAuth.expiresAt - Date.now()) / 60000);
    progress(`🔑 Using cached session (~${minsLeft} min remaining)`);
    return cachedAuth.cookieHeader;
  }

  // Check file cache before hitting CDP
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
      await ensureAlfred(progress);

      const auth = await getAuthCookiesViaCDP(progress);
      cachedAuth = auth;
      saveCachedAuth("dynamics", auth.cookieHeader, auth.expiresAt);

      const expiresIn = Math.round((auth.expiresAt - Date.now()) / 60000);
      progress(`✅ Session cookies acquired — valid for ~${expiresIn} minutes`);
      return auth.cookieHeader;
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

  // Check file cache before hitting CDP
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
      progress("🔐 Extracting Outlook cookies via CDP...");

      if (!isAlfredgable()) {
        throw new Error("Chrome debug port not available. Launch Alfred from your Desktop first.");
      }

      const allCookies = await getCookiesViaRawCDP([OUTLOOK_URL]);
      if (allCookies.length === 0) {
        process.stderr.write("[alfred] CDP auth: no Outlook cookies found — user not logged in\n");
        throw new Error(
          "Not logged into Outlook in the Alfred window.\n" +
          "Log into Outlook in the Alfred Chrome window, then retry."
        );
      }

      const cookieHeader = allCookies.map(c => `${c.name}=${c.value}`).join("; ");
      const expiresAt = Math.min(
        ...allCookies.map(c => c.expires > 0 ? c.expires * 1000 : Date.now() + 8 * 60 * 60 * 1000)
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

