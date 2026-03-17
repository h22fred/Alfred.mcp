import { execFileSync, execFile } from "child_process";

const DYNAMICS_URL = "https://servicenow.crm.dynamics.com";
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

export type ProgressFn = (msg: string) => void;

function isChromeProcessRunning(): boolean {
  try {
    execFileSync("pgrep", ["-f", `remote-debugging-port=${CDP_PORT}`], { timeout: 2_000 });
    return true;
  } catch {
    return false;
  }
}

function isChromeLinkgable(): boolean {
  try {
    execFileSync("curl", ["-s", "--max-time", "1", `http://localhost:${CDP_PORT}/json/version`], { timeout: 2_000 });
    return true;
  } catch {
    return false;
  }
}

function launchChromeLink(): void {
  if (isChromeProcessRunning()) {
    console.error("[auth] ChromeLink process already running — waiting for port to become ready...");
    return;
  }
  console.error("[auth] Launching ChromeLink...");
  execFile("/bin/sh", ["-c",
    `"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
      --remote-debugging-port=${CDP_PORT} \
      --no-first-run \
      --no-default-browser-check \
      > /dev/null 2>&1 &`
  ]);
}

async function waitForChrome(timeoutMs = 15_000): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (isChromeLinkgable()) return;
    await new Promise(r => setTimeout(r, 500));
  }
  throw new Error("ChromeLink did not start in time. Try opening ChromeLink.app manually.");
}

export function clearAuthCache(): void {
  cachedAuth = null;
  cachedOutlookAuth = null;
}

export async function ensureChromeLink(progress: ProgressFn = () => {}): Promise<void> {
  if (isChromeLinkgable()) return;
  // Chrome is not running — clear stale caches before launching fresh session
  clearAuthCache();
  progress("🚀 Launching ChromeLink automatically...");
  launchChromeLink();
  await waitForChrome();
  // Open Dynamics, Outlook and Teams tabs via CDP
  for (const url of [DYNAMICS_URL, OUTLOOK_URL, "https://teams.microsoft.com"]) {
    await fetch(`http://localhost:${CDP_PORT}/json/new?${url}`).catch(() => {});
  }
  progress("✅ ChromeLink ready — please log into Dynamics, Outlook and Teams in the new window");
}

export async function freshCdpEndpoint(): Promise<string> {
  // Always use the browser-level WebSocket URL (re-fetched every call — never stale)
  const res = await fetch(`http://localhost:${CDP_PORT}/json/version`);
  const info = await res.json() as { webSocketDebuggerUrl?: string };
  if (info.webSocketDebuggerUrl) return info.webSocketDebuggerUrl;
  throw new Error("Could not resolve CDP WebSocket URL from ChromeLink.");
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
    "Could not connect to ChromeLink. Please close and reopen ChromeLink from your Desktop, then retry.\n" +
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
    throw new Error("No page targets in ChromeLink. Make sure ChromeLink.app is running.");
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
      reject(new Error("CDP WebSocket error — is ChromeLink running?"));
    });
  });
}

async function getAuthCookiesViaCDP(progress: ProgressFn): Promise<CachedAuth> {
  progress("🍪 Extracting Dynamics session cookies via CDP...");
  const allCookies = await getCookiesViaRawCDP([DYNAMICS_URL]);

  const authCookies = allCookies.filter(c => AUTH_COOKIE_NAMES.includes(c.name));
  if (authCookies.length === 0) {
    throw new Error(
      "ChromeLink is open but you are not logged into Dynamics yet.\n" +
      "Please log into servicenow.crm.dynamics.com in the ChromeLink window, then retry."
    );
  }

  progress(`✅ Found ${authCookies.length} auth cookie(s): ${authCookies.map(c => c.name).join(", ")}`);

  const cookieHeader = authCookies.map(c => `${c.name}=${c.value}`).join("; ");
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

  progress("🔐 Acquiring Dynamics session cookies...");
  await ensureChromeLink(progress);

  const auth = await getAuthCookiesViaCDP(progress);
  cachedAuth = auth;

  const expiresIn = Math.round((auth.expiresAt - Date.now()) / 60000);
  progress(`✅ Session cookies acquired — valid for ~${expiresIn} minutes`);
  return auth.cookieHeader;
}

export async function getOutlookCookies(progress: ProgressFn = () => {}): Promise<string> {
  if (cachedOutlookAuth && Date.now() < cachedOutlookAuth.expiresAt - COOKIE_REFRESH_MARGIN_MS) {
    const minsLeft = Math.round((cachedOutlookAuth.expiresAt - Date.now()) / 60000);
    progress(`🔑 Using cached Outlook session (~${minsLeft} min remaining)`);
    return cachedOutlookAuth.cookieHeader;
  }

  progress("🔐 Extracting Outlook cookies via CDP...");

  if (!isChromeLinkgable()) {
    throw new Error("Chrome debug port not available. Open ChromeLink.app first.");
  }

  const allCookies = await getCookiesViaRawCDP([OUTLOOK_URL]);
  if (allCookies.length === 0) {
    throw new Error(
      "Not logged into Outlook in the ChromeLink window.\n" +
      "Log into https://outlook.office.com in the ChromeLink Chrome window, then retry."
    );
  }

  const cookieHeader = allCookies.map(c => `${c.name}=${c.value}`).join("; ");
  const expiresAt = Math.min(
    ...allCookies.map(c => c.expires > 0 ? c.expires * 1000 : Date.now() + 8 * 60 * 60 * 1000)
  );

  cachedOutlookAuth = { cookieHeader, expiresAt };
  const minsLeft = Math.round((expiresAt - Date.now()) / 60000);
  progress(`✅ Outlook cookies acquired — valid for ~${minsLeft} minutes`);
  return cookieHeader;
}

export function setManualCookies(cookieHeader: string): void {
  cachedAuth = {
    cookieHeader,
    expiresAt: Date.now() + 8 * 60 * 60 * 1000, // assume 8h
  };
  console.error(`[auth] Manual cookies set`);
}

export async function closeBrowser(): Promise<void> {
  // Nothing to close — we use the user's running Chrome
}

