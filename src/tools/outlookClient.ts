import { execFileSync } from "child_process";
import { dirname, join } from "path";
import { fileURLToPath } from "url";
import { connectWithRetry } from "../auth/tokenExtractor.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";
import { loadCachedAuth, saveCachedAuth, clearCachedAuthFile } from "../auth/authFileCache.js";
import { stripHtml, urlHostMatches } from "../shared.js";
import { sanitizeODataSearch } from "./dynamicsClient.js";

const CDP_PORT = 9222;

/** All known Outlook web domains — new ones can be added here. */
const OUTLOOK_DOMAINS = ["outlook.cloud.microsoft", "outlook.cloud.microsoft.com", "outlook.office.com", "outlook.office365.com", "outlook.live.com"] as const;

/** Match any known Outlook domain (legacy + new). */
function isOutlookUrl(url: string): boolean {
  return OUTLOOK_DOMAINS.some(d => urlHostMatches(url, d));
}

/** Quick check if an Alfred update is available — returns hint string or empty. */
function checkForAlfredUpdate(): string {
  try {
    const __fn = fileURLToPath(import.meta.url);
    const installDir = join(dirname(__fn), "..", "..");
    const localSha = execFileSync("git", ["-C", installDir, "rev-parse", "--short", "HEAD"], { encoding: "utf8", timeout: 5_000 }).trim();
    // Use already-fetched remote ref (git fetch runs at startup) — don't block on network here
    const remoteSha = execFileSync("git", ["-C", installDir, "rev-parse", "--short", "origin/main"], { encoding: "utf8", timeout: 3_000 }).trim();
    if (localSha !== remoteSha) {
      const behind = execFileSync("git", ["-C", installDir, "rev-list", "--count", `${localSha}..origin/main`], { encoding: "utf8", timeout: 3_000 }).trim();
      return `\n\n💡 An Alfred update is available (${behind} commit(s) behind). Ask Claude to run "update_alfred" — this may fix the issue.`;
    }
  } catch { /* git not available or not a git repo */ }
  return "";
}

/**
 * Detected Outlook origin from the browser tab (set during token acquisition).
 * E.g. "https://outlook.cloud.microsoft.com" or "https://outlook.office.com".
 */
let detectedOutlookOrigin: string | null = null;
/** Decode the audience (aud) claim from a JWT without verifying signature. */
function decodeJwtAudience(jwt: string): string {
  try {
    const payload = JSON.parse(Buffer.from(jwt.split(".")[1]!, "base64url").toString());
    return (typeof payload.aud === "string" ? payload.aud : "").toLowerCase();
  } catch { return ""; }
}

const TOKEN_CACHE_FALLBACK_MS = 45 * 60 * 1000; // 45 min — fallback when MSAL doesn't report expiry
const TOKEN_REFRESH_MARGIN_MS = 5 * 60 * 1000;  // refresh 5 min before expiry

interface TokenCache {
  token: string;
  expiresAt: number;
  aud?: string;  // JWT audience — determines which API endpoint to use
}
let tokenCache: TokenCache | null = null;
let outlookRestTokenCache: TokenCache | null = null;

export function clearGraphTokenCache(): void {
  tokenCache = null;
  clearCachedAuthFile("graphToken");
}

export function clearOutlookRestTokenCache(): void {
  outlookRestTokenCache = null;
  clearCachedAuthFile("outlookRestToken");
}

/** @internal — test-only helper to pre-seed the Graph token cache */
export function _seedGraphTokenCache(token: string, ttlMs = TOKEN_CACHE_FALLBACK_MS): void {
  tokenCache = { token, expiresAt: Date.now() + ttlMs };
}

/** @internal — test-only helper to pre-seed the Outlook REST token cache */
export function _seedOutlookRestTokenCache(token: string, ttlMs = TOKEN_CACHE_FALLBACK_MS): void {
  outlookRestTokenCache = { token, expiresAt: Date.now() + ttlMs };
}

// ---------------------------------------------------------------------------
// Acquire a Graph Bearer token via raw CDP WebSocket — no Playwright needed.
// Enables Network tracking on the Outlook tab, triggers a lightweight OWA
// service call, and captures the outgoing Authorization header.
// ---------------------------------------------------------------------------

// JS snippet injected into a browser tab to extract a Bearer token from MSAL cache.
// Scans sessionStorage + localStorage for ANY AccessToken entry (MSAL v2 format).
// Also checks for MSAL v2 "accesstoken" key pattern used by outlook.cloud.microsoft.com.
// Returns { token, expiresAt, debug } for diagnostics.
const MSAL_EXTRACT_JS = `(function() {
  const diag = { sessionKeys: 0, localKeys: 0, accessTokens: 0, matched: [], skippedExpired: 0, origin: location.origin };
  const isJwt = (s) => typeof s === 'string' && s.length > 100 && s.split('.').length === 3;
  const tryStorage = (s, sName) => {
    try {
      const keys = Object.keys(s);
      if (sName === 'session') diag.sessionKeys = keys.length;
      else diag.localKeys = keys.length;
      let bestScoped = null;
      let bestAny = null;
      for (const key of keys) {
        try {
          const raw = s.getItem(key);
          if (!raw || raw.length < 50) continue;
          const val = JSON.parse(raw);
          if (!val || typeof val !== 'object') continue;
          // MSAL v2 format: credentialType=AccessToken, secret=<jwt>
          const secret = val.secret || val.access_token || val.accessToken;
          const credType = (val.credentialType || val.credential_type || '').toLowerCase();
          if (credType.includes('accesstoken') && secret && isJwt(secret)) {
            diag.accessTokens++;
            const target = (val.target || val.scopes || '').toLowerCase();
            const exp = Number(val.expiresOn || val.expires_on || val.extended_expires_on || 0);
            diag.matched.push({ key: key.slice(0, 60), target: target.slice(0, 80), exp, sName });
            if (exp && exp * 1000 < Date.now()) { diag.skippedExpired++; continue; }
            const result = { token: secret, expiresAt: exp ? exp * 1000 : 0 };
            if (target.includes('mail') || target.includes('calendar') || target.includes('outlook')) {
              if (!bestScoped) bestScoped = result;
            } else {
              if (!bestAny) bestAny = result;
            }
          }
          // Also check for plain JWT stored directly (some new Outlook versions)
          else if (key.toLowerCase().includes('accesstoken') && isJwt(raw)) {
            diag.accessTokens++;
            diag.matched.push({ key: key.slice(0, 60), target: 'raw-jwt-key', exp: 0, sName });
            if (!bestAny) bestAny = { token: raw, expiresAt: 0 };
          }
        } catch {}
      }
      return bestScoped || bestAny;
    } catch {}
    return null;
  };
  const token = tryStorage(sessionStorage, 'session') || tryStorage(localStorage, 'local');
  if (token) {
    token.debug = diag;
    return token;
  }
  return { token: null, expiresAt: 0, debug: diag };
})()`;

async function acquireGraphTokenRawCDP(progress: ProgressFn): Promise<string> {
  if (tokenCache && Date.now() < tokenCache.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    const mins = Math.round((tokenCache.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Graph token (~${mins} min remaining)`);
    return tokenCache.token;
  }

  // Check file cache before hitting CDP — apply refresh margin
  const fileCached = loadCachedAuth("graphToken");
  if (fileCached && Date.now() < fileCached.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    tokenCache = { token: fileCached.value, expiresAt: fileCached.expiresAt };
    const mins = Math.round((fileCached.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Graph token (~${mins} min remaining)`);
    return fileCached.value;
  }

  progress("🔑 Acquiring Graph Bearer token via CDP...");

  const listRes = await fetch(`http://localhost:${CDP_PORT}/json/list`);
  const targets = await listRes.json() as Array<{ webSocketDebuggerUrl?: string; type?: string; url?: string }>;

  // Prefer Teams tab — it actually talks to graph.microsoft.com.
  // Outlook uses the Outlook REST API (outlook.office.com), so its MSAL cache
  // rarely has Graph-audience tokens. Try Outlook only as fallback.
  const teamsTarget   = targets.find(t => t.type === "page" && t.url && urlHostMatches(t.url, "teams.microsoft.com") && t.webSocketDebuggerUrl);
  const outlookTarget = targets.find(t => t.type === "page" && t.url && isOutlookUrl(t.url) && t.webSocketDebuggerUrl);
  const anyTarget     = targets.find(t => t.type === "page" && t.webSocketDebuggerUrl);
  const target        = teamsTarget ?? outlookTarget ?? anyTarget;

  if (!target?.webSocketDebuggerUrl) {
    throw new Error("No browser tabs found. Launch Alfred from your Desktop and log into Teams or Outlook.");
  }

  // Log which tab we're using so the user can see what's happening
  const tabLabel = teamsTarget ? "Teams" : outlookTarget ? "Outlook" : "other";
  const tabUrl = target.url ?? "unknown";
  process.stderr.write(`[alfred] CDP: using ${tabLabel} tab (${tabUrl.slice(0, 80)})\n`);
  if (!teamsTarget) {
    process.stderr.write(`[alfred] CDP: no Teams tab found — trying ${outlookTarget ? "Outlook" : "other"} tab for MSAL cache\n`);
  }

  // Detect if tabs are stuck on login pages
  for (const t of [teamsTarget, outlookTarget]) {
    if (t?.url && (t.url.includes("login.microsoftonline.com") || t.url.includes("login.live.com") || t.url.includes("/oauth2/"))) {
      throw new Error(
        "A tab in Alfred is on the Microsoft login page — your session has expired.\n" +
        "Please log back into Teams/Outlook in the Alfred window, then retry."
      );
    }
  }

  // Step 1: try reading token directly from MSAL storage — fast, no network interception
  interface MsalDiag { sessionKeys: number; localKeys: number; accessTokens: number; matched: Array<{ key: string; target: string; exp: number; sName: string }>; skippedExpired: number; origin: string }
  const msalResult = await new Promise<{ token: string | null; expiresAt: number; debug?: MsalDiag } | null>((resolve) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    const timer = setTimeout(() => {
      process.stderr.write("[alfred:warn] MSAL extraction timed out (5s) — tab may be unresponsive\n");
      try { ws.close(); } catch {}
      resolve(null);
    }, 5_000);

    ws.addEventListener("open", () => {
      ws.send(JSON.stringify({ id: 1, method: "Runtime.evaluate", params: { expression: MSAL_EXTRACT_JS, returnByValue: true } }));
    });
    ws.addEventListener("message", (event: MessageEvent) => {
      clearTimeout(timer);
      try { ws.close(); } catch {}
      try {
        const msg = JSON.parse(event.data as string) as { result?: { result?: { value?: { token: string | null; expiresAt: number; debug?: MsalDiag } | null } } };
        resolve(msg.result?.result?.value ?? null);
      } catch (e) {
        process.stderr.write(`[alfred:warn] MSAL extraction parse error: ${e instanceof Error ? e.message : String(e)}\n`);
        resolve(null);
      }
    });
    ws.addEventListener("error", () => {
      clearTimeout(timer);
      process.stderr.write(`[alfred:warn] MSAL extraction WebSocket error on tab\n`);
      try { ws.close(); } catch {}
      resolve(null);
    });
  });

  // Log diagnostics regardless of outcome — both stderr and user-facing progress
  if (msalResult?.debug) {
    const d = msalResult.debug;
    const diagMsg = `MSAL scan on ${d.origin}: ${d.sessionKeys} session keys, ${d.localKeys} local keys, ${d.accessTokens} AccessTokens, ${d.skippedExpired} expired`;
    process.stderr.write(`[alfred:diag] ${diagMsg}\n`);
    progress(`🔍 ${diagMsg}`);
    for (const m of d.matched.slice(0, 5)) {
      process.stderr.write(`[alfred:diag]   key="${m.key}" target="${m.target}" exp=${m.exp} storage=${m.sName}\n`);
    }
  } else if (!msalResult) {
    progress("⚠️ MSAL extraction returned no result (timeout or WebSocket error)");
  }

  if (msalResult?.token) {
    // Use real MSAL expiry when available, otherwise fall back to 45 min
    const expiresAt = msalResult.expiresAt > 0 ? msalResult.expiresAt : Date.now() + TOKEN_CACHE_FALLBACK_MS;
    tokenCache = { token: msalResult.token, expiresAt };
    saveCachedAuth("graphToken", msalResult.token, expiresAt);
    const mins = Math.round((expiresAt - Date.now()) / 60_000);
    progress(`✅ Graph token acquired from MSAL cache (~${mins} min valid)`);
    return msalResult.token;
  }

  // Step 2: fallback — enable Network tracking and trigger an OWA API fetch
  progress("🔑 MSAL cache miss — capturing token via network interception...");

  return new Promise((resolve, reject) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    let capturedToken: string | null = null;

    const timer = setTimeout(() => {
      try { ws.close(); } catch {}
      process.stderr.write("[alfred] CDP: Graph token capture timed out from Outlook tab\n");
      reject(new Error(
        "Could not capture Graph token from Outlook.\n" +
        "Make sure you are logged into Outlook in the Alfred window." +
        checkForAlfredUpdate()
      ));
    }, 10_000);

    const done = (token: string) => {
      clearTimeout(timer);
      try { ws.close(); } catch {}
      const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
      tokenCache = { token, expiresAt };
      saveCachedAuth("graphToken", token, expiresAt);
      progress("✅ Graph token acquired");
      resolve(token);
    };

    let msgId = 0;
    const send = (method: string, params?: Record<string, unknown>) =>
      ws.send(JSON.stringify({ id: ++msgId, method, params }));

    ws.addEventListener("open", () => send("Network.enable"));

    ws.addEventListener("message", (event: MessageEvent) => {
      try {
        const msg = JSON.parse(event.data as string) as { id?: number; method?: string; params?: Record<string, unknown> };
        if (msg.id === 1) {
          // Network enabled — trigger API calls against all known Outlook API domains
          const now = new Date().toISOString();
          const tomorrow = new Date(Date.now()+86400000).toISOString();
          send("Runtime.evaluate", {
            expression: [
              `fetch('https://outlook.microsoft.com/api/v2.0/me/messages?$top=1', { credentials: 'include' }).catch(()=>{})`,
              `fetch('https://outlook.cloud.microsoft.com/api/v2.0/me/messages?$top=1', { credentials: 'include' }).catch(()=>{})`,
              `fetch('https://outlook.office.com/api/v2.0/me/messages?$top=1', { credentials: 'include' }).catch(()=>{})`,
              `fetch('/api/v2.0/me/calendarview?$top=1&$select=Id&startDateTime=${now}&endDateTime=${tomorrow}', { credentials: 'include' }).catch(()=>{})`,
            ].join(';'),
            awaitPromise: false,
          });
          setTimeout(() => { if (!capturedToken) send("Page.reload"); }, 3_000);
        }
        if (msg.method === "Network.requestWillBeSent") {
          const headers = ((msg.params?.request as Record<string, unknown>)?.headers ?? {}) as Record<string, string>;
          const auth = headers["Authorization"] ?? headers["authorization"] ?? "";
          if (!capturedToken && auth.startsWith("Bearer ")) {
            capturedToken = auth.slice(7);
            done(capturedToken);
          }
        }
      } catch (e) { process.stderr.write(`[alfred:warn] CDP message parse error: ${e instanceof Error ? e.message : String(e)}\n`); }
    });

    ws.addEventListener("error", () => {
      clearTimeout(timer);
      reject(new Error("CDP WebSocket error — is Alfred running?"));
    });
  });
}

// acquireGraphToken — tries raw CDP first (fast), falls back to Playwright page load
async function acquireGraphToken(progress: ProgressFn): Promise<string> {
  try {
    return await acquireGraphTokenRawCDP(progress);
  } catch (e) {
    process.stderr.write(`[alfred:warn] raw CDP token acquisition failed, falling back to Playwright: ${e instanceof Error ? e.message : String(e)}\n`);
    // Raw CDP failed — use Playwright to capture a Graph API token.
    // Prefer Teams tab (it talks to graph.microsoft.com), fall back to Outlook.
    progress("📡 Falling back to Playwright token capture via Teams...");
    const browser = await connectWithRetry();
    try {
      const ctx = browser.contexts()[0];
      if (!ctx) throw new Error("No browser context found in Alfred");

      const existingPages = ctx.pages();
      // Prefer Teams — it actually uses Graph API. Outlook uses its own REST API.
      let page = existingPages.find(p => urlHostMatches(p.url(), "teams.microsoft.com"))
              ?? existingPages.find(p => isOutlookUrl(p.url()));

      let isNewPage = false;
      if (!page) {
        page = existingPages.find(p => p.url().startsWith("http")) ?? await ctx.newPage();
        isNewPage = !existingPages.includes(page);
      }

      let capturedToken: string | null = null;
      // Only capture Bearer tokens destined for Graph API
      await page.route("**/graph.microsoft.com/**", async (route) => {
        const auth = route.request().headers()["authorization"] ?? "";
        if (!capturedToken && auth.startsWith("Bearer ")) capturedToken = auth.slice(7);
        await route.continue();
      });

      // Navigate/reload to trigger Graph API calls
      if (urlHostMatches(page.url(), "teams.microsoft.com")) {
        await page.reload({ waitUntil: "domcontentloaded", timeout: 20_000 }).catch((e) => { process.stderr.write(`[alfred:warn] Teams page.reload failed: ${e instanceof Error ? e.message : String(e)}\n`); });
      } else if (!urlHostMatches(page.url(), "teams.microsoft.com")) {
        await page.goto("https://teams.microsoft.com/v2/", { waitUntil: "domcontentloaded", timeout: 20_000 }).catch((e) => { process.stderr.write(`[alfred:warn] Teams page.goto failed: ${e instanceof Error ? e.message : String(e)}\n`); });
      }

      const deadline = Date.now() + 10_000;
      while (!capturedToken && Date.now() < deadline) await page.waitForTimeout(500);

      // Clean up route interceptor
      await page.unroute("**/graph.microsoft.com/**").catch((e) => { process.stderr.write(`[alfred:warn] unroute failed: ${e instanceof Error ? e.message : String(e)}\n`); });
      // Only close pages we created
      if (isNewPage) await page.close();

      if (!capturedToken) throw new Error("Could not capture Graph token from Outlook.\nMake sure you are logged into Outlook in the Alfred window." + checkForAlfredUpdate());
      const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
      tokenCache = { token: capturedToken, expiresAt };
      saveCachedAuth("graphToken", capturedToken, expiresAt);
      progress("✅ Graph token acquired via Playwright");
      return capturedToken;
    } finally {
      // Do NOT call browser.close() — it kills the user's actual Alfred Chrome process.
      // The CDP connection wrapper is GC'd; Alfred Chrome keeps running.
    }
  }
}

// ---------------------------------------------------------------------------
// Outlook REST token — for email/folder operations against outlook.office.com
// Separate from Graph token (which is for calendar via graph.microsoft.com)
// ---------------------------------------------------------------------------

const OUTLOOK_REST_MSAL_JS = `(function() {
  const decodeAud = (jwt) => {
    try {
      const payload = JSON.parse(atob(jwt.split('.')[1].replace(/-/g,'+').replace(/_/g,'/')));
      return (payload.aud || '').toLowerCase();
    } catch { return ''; }
  };
  const isMailAud = (aud) =>
    aud.includes('outlook.office') || aud.includes('outlook.cloud.microsoft') ||
    aud.includes('outlook.microsoft') || aud.includes('substrate.office') ||
    aud.includes('graph.microsoft.com');
  const isJwt = (s) => typeof s === 'string' && s.length > 100 && s.split('.').length === 3;
  const tryStorage = (s) => {
    try {
      for (const key of Object.keys(s)) {
        try {
          const raw = s.getItem(key);
          if (!raw || raw.length < 50) continue;
          const val = JSON.parse(raw);
          if (!val || typeof val !== 'object') continue;
          const secret = val.secret || val.access_token || val.accessToken;
          const credType = (val.credentialType || val.credential_type || '').toLowerCase();
          if (credType.includes('accesstoken') && secret && isJwt(secret)) {
            const aud = decodeAud(secret);
            if (!isMailAud(aud)) continue;
            const exp = Number(val.expiresOn || val.expires_on || val.extended_expires_on || 0);
            if (exp && exp * 1000 < Date.now()) continue;
            return { token: secret, expiresAt: exp ? exp * 1000 : 0, aud };
          }
        } catch {}
      }
    } catch {}
    return null;
  };
  return tryStorage(sessionStorage) || tryStorage(localStorage);
})()`;

// ---------------------------------------------------------------------------
// Pre-flight connection check — called before every Outlook tool operation
// ---------------------------------------------------------------------------

/**
 * Verify that the Alfred browser is running and an Outlook tab is active.
 * Throws a clear, actionable error if the connection is not healthy.
 */
async function verifyOutlookConnection(progress: ProgressFn): Promise<void> {
  // Skip check if we have a valid cached token (connection was recently verified)
  if (outlookRestTokenCache && Date.now() < outlookRestTokenCache.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    return;
  }

  const updateHint = checkForAlfredUpdate();

  // Step 1: Can we reach the Alfred browser (CDP)?
  let targets: Array<{ type?: string; url?: string; webSocketDebuggerUrl?: string }>;
  try {
    const res = await fetch(`http://localhost:${CDP_PORT}/json/list`, { signal: AbortSignal.timeout(3_000) });
    targets = await res.json() as typeof targets;
  } catch {
    throw new Error(
      "Cannot connect to Alfred — the Alfred browser window is not running or CDP is not reachable.\n" +
      "Please launch Alfred from your Desktop and make sure it's running." +
      updateHint
    );
  }

  // Step 2: Is there an Outlook tab open?
  // Log all targets for diagnostics — helps debug when matching fails
  const targetSummary = targets.map(t => `${t.type ?? "?"}|${t.url?.slice(0, 80) ?? "no-url"}`).join("; ");
  process.stderr.write(`[alfred:cdp] Targets (${targets.length}): ${targetSummary}\n`);

  // Try page first, then any target type (service_worker, iframe, etc.)
  const outlookTab = targets.find(t => t.type === "page" && t.url && isOutlookUrl(t.url))
    ?? targets.find(t => t.url && isOutlookUrl(t.url));
  if (!outlookTab) {
    throw new Error(
      "Alfred is running but no Outlook tab was found.\n" +
      `CDP returned ${targets.length} target(s): ${targetSummary}\n` +
      "Please open Outlook (outlook.cloud.microsoft.com) in the Alfred window and log in." +
      updateHint
    );
  }

  // Step 3: Is the Outlook tab stuck on a login page?
  const tabUrl = outlookTab.url ?? "";
  if (tabUrl.includes("login.microsoftonline.com") || tabUrl.includes("login.live.com") || tabUrl.includes("/oauth2/")) {
    throw new Error(
      "Outlook session has expired — the Alfred browser is on the Microsoft login page.\n" +
      "Please log back into Outlook in the Alfred window (make sure the inbox loads fully), then retry." +
      updateHint
    );
  }

  progress("✅ Alfred connection verified — Outlook tab is active");
}

async function acquireOutlookRestToken(progress: ProgressFn): Promise<string> {
  // Check in-memory cache (fast path — no CDP call needed)
  if (outlookRestTokenCache && Date.now() < outlookRestTokenCache.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    const mins = Math.round((outlookRestTokenCache.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Outlook REST token (~${mins} min remaining)`);
    return outlookRestTokenCache.token;
  }

  // Check file cache
  const fileCached = loadCachedAuth("outlookRestToken");
  if (fileCached && Date.now() < fileCached.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    const aud = decodeJwtAudience(fileCached.value);
    outlookRestTokenCache = { token: fileCached.value, expiresAt: fileCached.expiresAt, aud };
    const mins = Math.round((fileCached.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Outlook REST token (~${mins} min remaining)`);
    return fileCached.value;
  }

  progress("🔑 Acquiring Outlook REST token via CDP...");

  const listRes = await fetch(`http://localhost:${CDP_PORT}/json/list`);
  const targets = await listRes.json() as Array<{ webSocketDebuggerUrl?: string; type?: string; url?: string }>;

  // Detect which Outlook domain is open — sets detectedOutlookOrigin for API URL resolution
  if (!detectedOutlookOrigin) {
    const outlookTabUrl = (targets.find(t => t.type === "page" && t.url && isOutlookUrl(t.url))
      ?? targets.find(t => t.url && isOutlookUrl(t.url)))?.url;
    if (outlookTabUrl) {
      try {
        detectedOutlookOrigin = new URL(outlookTabUrl).origin;
        process.stderr.write(`[alfred] Detected Outlook origin: ${detectedOutlookOrigin}\n`);
      } catch {}
    }
  }

  // Outlook REST tokens live in the Outlook tab's MSAL cache
  // Try page first, then any target type (service_worker, iframe, etc.)
  const outlookTarget = targets.find(t => t.type === "page" && t.url && isOutlookUrl(t.url) && t.webSocketDebuggerUrl)
    ?? targets.find(t => t.url && isOutlookUrl(t.url) && t.webSocketDebuggerUrl);
  const anyTarget = targets.find(t => t.type === "page" && t.webSocketDebuggerUrl)
    ?? targets.find(t => t.webSocketDebuggerUrl);
  const target = outlookTarget ?? anyTarget;

  if (!target?.webSocketDebuggerUrl) {
    throw new Error("No browser tabs found. Launch Alfred from your Desktop and log into Outlook.");
  }

  if (!outlookTarget) {
    process.stderr.write("[alfred] CDP: no Outlook tab found — trying other tab for Outlook REST MSAL token\n");
  }

  // Detect login-page redirect
  if (outlookTarget?.url) {
    const tabUrl = outlookTarget.url;
    if (tabUrl.includes("login.microsoftonline.com") || tabUrl.includes("login.live.com") || tabUrl.includes("/oauth2/")) {
      throw new Error(
        "Outlook tab is on the Microsoft login page — your session has expired.\n" +
        "Please log back into Outlook in the Alfred window (make sure the inbox is fully loaded), then retry."
      );
    }
  }

  // Extract Outlook REST token from MSAL cache
  const msalResult = await new Promise<{ token: string; expiresAt: number; aud?: string } | null>((resolve) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    const timer = setTimeout(() => {
      process.stderr.write("[alfred:warn] Outlook REST MSAL extraction timed out (5s)\n");
      try { ws.close(); } catch {}
      resolve(null);
    }, 5_000);

    ws.addEventListener("open", () => {
      ws.send(JSON.stringify({ id: 1, method: "Runtime.evaluate", params: { expression: OUTLOOK_REST_MSAL_JS, returnByValue: true } }));
    });
    ws.addEventListener("message", (event: MessageEvent) => {
      clearTimeout(timer);
      try { ws.close(); } catch {}
      try {
        const msg = JSON.parse(event.data as string) as { result?: { result?: { value?: { token: string; expiresAt: number; aud?: string } | null } } };
        resolve(msg.result?.result?.value ?? null);
      } catch (e) {
        process.stderr.write(`[alfred:warn] Outlook REST MSAL parse error: ${e instanceof Error ? e.message : String(e)}\n`);
        resolve(null);
      }
    });
    ws.addEventListener("error", () => {
      clearTimeout(timer);
      process.stderr.write("[alfred:warn] Outlook REST MSAL extraction WebSocket error\n");
      try { ws.close(); } catch {};
      resolve(null);
    });
  });

  if (msalResult) {
    const expiresAt = msalResult.expiresAt > 0 ? msalResult.expiresAt : Date.now() + TOKEN_CACHE_FALLBACK_MS;
    outlookRestTokenCache = { token: msalResult.token, expiresAt, aud: msalResult.aud };
    saveCachedAuth("outlookRestToken", msalResult.token, expiresAt);
    const mins = Math.round((expiresAt - Date.now()) / 60_000);
    progress(`✅ Outlook REST token acquired from MSAL cache (~${mins} min valid, audience: ${msalResult.aud ?? "unknown"})`);
    return msalResult.token;
  }

  // MSAL cache may be encrypted (outlook.cloud.microsoft.com uses AES-GCM).
  // Fallback: capture Bearer token from the Outlook tab's own API requests.
  progress("🔑 MSAL cache unreadable — capturing token via network interception...");

  const capturedResult = await new Promise<{ token: string; requestUrl: string } | null>((resolve) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    let found = false;

    const timer = setTimeout(() => {
      process.stderr.write("[alfred:warn] Outlook REST network interception timed out (10s)\n");
      try { ws.close(); } catch {}
      resolve(null);
    }, 10_000);

    let msgId = 0;
    const send = (method: string, params?: Record<string, unknown>) =>
      ws.send(JSON.stringify({ id: ++msgId, method, params }));

    ws.addEventListener("open", () => send("Network.enable"));

    ws.addEventListener("message", (event: MessageEvent) => {
      try {
        const msg = JSON.parse(event.data as string) as { id?: number; method?: string; params?: Record<string, unknown> };
        if (msg.id === 1) {
          // Network enabled — reload the page (ignoring cache) to trigger fresh authenticated requests.
          // This forces OWA to re-acquire tokens and call startupdata.ashx with a full-scope Bearer token.
          send("Page.reload", { ignoreCache: true });
          // Backup: if reload doesn't generate captures in 3s, try explicit fetches
          setTimeout(() => {
            if (!found) {
              const now = new Date().toISOString();
              const tomorrow = new Date(Date.now()+86400000).toISOString();
              send("Runtime.evaluate", {
                expression: [
                  `fetch('/owa/service.svc?action=GetFolder', {method:'POST', credentials:'include', headers:{'Content-Type':'application/json'}}).catch(()=>{})`,
                  `fetch('/api/v2.0/me/calendarview?$top=1&$select=Id&startDateTime=${now}&endDateTime=${tomorrow}', { credentials: 'include' }).catch(()=>{})`,
                  `fetch('https://graph.microsoft.com/v1.0/me/messages?$top=1', { credentials: 'include' }).catch(()=>{})`,
                ].join(';'),
                awaitPromise: false,
              });
            }
          }, 3_000);
        }
        if (msg.method === "Network.requestWillBeSent") {
          const reqUrl = ((msg.params?.request as Record<string, unknown>)?.url ?? "") as string;
          const headers = ((msg.params?.request as Record<string, unknown>)?.headers ?? {}) as Record<string, string>;
          const auth = headers["Authorization"] ?? headers["authorization"] ?? "";
          if (!found && auth.startsWith("Bearer ")) {
            const candidateToken = auth.slice(7);
            // Decode JWT — only accept tokens with Outlook audience AND mail/calendar scopes.
            // The notification channel token has audience outlook.office.com but only Owa.Notifications.All scope → 403 on REST API.
            const candidateAud = decodeJwtAudience(candidateToken);
            if (candidateAud.includes("outlook.office") || candidateAud.includes("outlook.cloud.microsoft") || candidateAud.includes("outlook.microsoft")) {
              // Verify the token has mail/calendar scopes (not just notification scopes)
              let scp = "";
              try { scp = JSON.parse(Buffer.from(candidateToken.split(".")[1]!, "base64url").toString()).scp ?? ""; } catch {}
              if (scp.includes("Mail") || scp.includes("Calendar")) {
                found = true;
                clearTimeout(timer);
                try { ws.close(); } catch {}
                resolve({ token: candidateToken, requestUrl: reqUrl });
              } else {
                process.stderr.write(`[alfred:cdp] Skipping Outlook token with limited scopes "${scp.slice(0, 40)}" from ${reqUrl.slice(0, 60)}\n`);
              }
            } else {
              process.stderr.write(`[alfred:cdp] Skipping token with audience "${candidateAud}" (not Outlook) from ${reqUrl.slice(0, 60)}\n`);
            }
          }
        }
      } catch (e) { process.stderr.write(`[alfred:warn] CDP network interception error: ${e instanceof Error ? e.message : String(e)}\n`); }
    });

    ws.addEventListener("error", () => {
      clearTimeout(timer);
      process.stderr.write("[alfred:warn] Outlook REST network interception WebSocket error\n");
      try { ws.close(); } catch {}
      resolve(null);
    });
  });

  if (capturedResult) {
    const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
    // Derive audience from the JWT itself (more reliable than request URL)
    const capturedAud = decodeJwtAudience(capturedResult.token) || (() => { try { return new URL(capturedResult.requestUrl).origin.toLowerCase(); } catch { return ""; } })();
    outlookRestTokenCache = { token: capturedResult.token, expiresAt, aud: capturedAud };
    saveCachedAuth("outlookRestToken", capturedResult.token, expiresAt);
    progress(`✅ Outlook REST token captured via network interception (${capturedAud || "unknown origin"})`);
    return capturedResult.token;
  }

  const updateHint = checkForAlfredUpdate();
  throw new Error(
    "Could not capture Outlook token.\n" +
    "Make sure Outlook is fully loaded (inbox visible) in the Alfred window, then retry." +
    updateHint
  );
}

// Resolve the correct mail API base URL based on token audience.
// New Outlook uses Graph API (graph.microsoft.com), old uses REST v2.0 (outlook.office.com).
function getOutlookApiBase(): string {
  const aud = outlookRestTokenCache?.aud ?? "";
  // Graph API token — new Outlook uses this
  if (aud.includes("graph.microsoft.com")) return "https://graph.microsoft.com/v1.0/me";
  // Outlook REST API variants
  if (aud.includes("outlook.office365")) return "https://outlook.office365.com/api/v2.0/me";
  if (aud.includes("outlook.cloud.microsoft")) return "https://outlook.cloud.microsoft.com/api/v2.0/me";
  if (aud.includes("outlook.microsoft")) return "https://outlook.microsoft.com/api/v2.0/me";
  if (aud.includes("outlook.office.com")) return "https://outlook.office.com/api/v2.0/me";
  // If we detected the browser origin, use that
  if (detectedOutlookOrigin) return `${detectedOutlookOrigin}/api/v2.0/me`;
  // Fallback — try Graph API first (new Outlook default)
  return "https://graph.microsoft.com/v1.0/me";
}

// ---------------------------------------------------------------------------
// Outlook REST v2 fetch using Outlook REST Bearer token
// ---------------------------------------------------------------------------

async function outlookApiFetch(path: string, token: string, progress?: ProgressFn, _retryCount = 0): Promise<Record<string, unknown>> {
  const res = await fetch(`${getOutlookApiBase()}${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
    signal: AbortSignal.timeout(30_000),
  });

  if (res.status === 429 && _retryCount < 3) {
    const retryAfter = parseInt(res.headers.get("Retry-After") ?? "", 10);
    const delayMs = retryAfter > 0 ? retryAfter * 1000 : 1000 * Math.pow(2, _retryCount);
    progress?.(`⏳ Outlook API throttled (429) — retrying in ${(delayMs / 1000).toFixed(0)}s...`);
    await new Promise(r => setTimeout(r, delayMs));
    return outlookApiFetch(path, token, progress, _retryCount + 1);
  }

  if (res.status === 401) {
    outlookRestTokenCache = null;
    clearCachedAuthFile("outlookRestToken");
    progress?.("🔄 Outlook REST token expired — re-acquiring...");
    const freshToken = await acquireOutlookRestToken(progress ?? (() => {}));
    const retry = await fetch(`${getOutlookApiBase()}${path}`, {
      headers: { Authorization: `Bearer ${freshToken}`, Accept: "application/json" },
      signal: AbortSignal.timeout(30_000),
    });
    if (!retry.ok) {
      // Clear caches again so next call starts fresh
      outlookRestTokenCache = null;
      clearCachedAuthFile("outlookRestToken");
      if (retry.status === 401) {
        const updateHint = checkForAlfredUpdate();
        throw new Error(
          "Outlook session has expired — re-acquired token was also rejected (401).\n" +
          "Please go to the Alfred window and log back into Outlook (make sure the inbox loads fully), then retry." +
          updateHint
        );
      }
      const body = await retry.text().catch(() => "");
      throw new Error(`Outlook API ${retry.status} ${retry.statusText}${body ? `: ${body.slice(0, 200)}` : ""}`);
    }
    return retry.json() as Promise<Record<string, unknown>>;
  }

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Outlook API ${res.status} ${res.statusText}${body ? `: ${body.slice(0, 200)}` : ""}`);
  }

  return res.json() as Promise<Record<string, unknown>>;
}


// ---------------------------------------------------------------------------
// Calendar events
// ---------------------------------------------------------------------------

export interface CalendarEvent {
  id: string;
  subject: string;
  start: string;
  end: string;
  location?: string;
  organizer?: string;
  organizerEmail?: string;
  attendees?: { name: string; email: string }[];
  isOnlineMeeting?: boolean;
  bodyPreview?: string;
  webLink?: string;
}

export async function getCalendarEvents(
  startDate: string,
  endDate: string,
  search?: string,
  progress: ProgressFn = () => {},
  top: number = 100,
): Promise<CalendarEvent[]> {
  await verifyOutlookConnection(progress);
  progress(`📅 Fetching calendar events ${startDate} → ${endDate}${search ? ` (filter: "${search}")` : ""}...`);

  // NOTE: /calendarview does NOT support $search — we filter client-side below.
  // Use same proven token path as emails — routes to Graph API or Outlook REST depending on token audience.
  // Use camelCase field names — works with both Graph API (canonical) and Outlook REST (case-insensitive).
  const params = new URLSearchParams({
    startDateTime: `${startDate}T00:00:00Z`,
    endDateTime:   `${endDate}T23:59:59Z`,
    $select: "subject,start,end,location,organizer,attendees,isOnlineMeeting,webLink",
    $top: String(top),
    $orderby: "start/dateTime",
  });

  if (search) {
    try {
      const safe = sanitizeODataSearch(search);
      params.set("$filter", `contains(subject,'${safe}')`);
    } catch (e) {
      process.stderr.write(`[alfred:warn] OData search sanitize failed, using client-side filter: ${e instanceof Error ? e.message : String(e)}\n`);
    }
  }

  const token = await acquireOutlookRestToken(progress);

  let data: { value?: Record<string, unknown>[] };
  try {
    data = await outlookApiFetch(`/calendarview?${params}`, token, progress) as { value?: Record<string, unknown>[] };
  } catch (e) {
    // If $filter is not supported (400), retry without it
    if (search && params.has("$filter") && e instanceof Error && (e.message.includes("400") || e.message.includes("UnsupportedQuery"))) {
      progress("⚠️ Server-side filter not supported on calendarView — falling back to client-side filter...");
      params.delete("$filter");
      data = await outlookApiFetch(`/calendarview?${params}`, token, progress) as { value?: Record<string, unknown>[] };
    } else {
      throw e;
    }
  }

  // Handle both PascalCase (Outlook REST) and camelCase (Graph) response fields
  let events = (data.value ?? []).map(e => {
    const orgRaw = (e.Organizer ?? e.organizer) as { EmailAddress?: { Name: string; Address: string }; emailAddress?: { name: string; address: string } } | undefined;
    const org = orgRaw?.EmailAddress
      ? { name: orgRaw.EmailAddress.Name, address: orgRaw.EmailAddress.Address }
      : orgRaw?.emailAddress
      ? { name: orgRaw.emailAddress.name, address: orgRaw.emailAddress.address }
      : undefined;
    const rawAttendees = ((e.Attendees ?? e.attendees) as Array<{ EmailAddress?: { Name: string; Address: string }; emailAddress?: { name: string; address: string } }>) ?? [];
    const attendees = rawAttendees.map(a => ({
      name:  a.EmailAddress?.Name  ?? a.emailAddress?.name  ?? "",
      email: a.EmailAddress?.Address ?? a.emailAddress?.address ?? "",
    }));
    const startField = (e.Start ?? e.start) as { DateTime?: string; dateTime?: string } | undefined;
    const endField   = (e.End ?? e.end) as { DateTime?: string; dateTime?: string } | undefined;
    const locationField = (e.Location ?? e.location) as { DisplayName?: string; displayName?: string } | undefined;
    return {
      id:              "",
      subject:         ((e.Subject ?? e.subject) as string) || "",
      start:           startField?.DateTime ?? startField?.dateTime ?? "",
      end:             endField?.DateTime ?? endField?.dateTime ?? "",
      location:        locationField?.DisplayName ?? locationField?.displayName ?? undefined,
      organizer:       org?.name || undefined,
      organizerEmail:  org?.address || undefined,
      attendees,
      isOnlineMeeting: (e.IsOnlineMeeting ?? e.isOnlineMeeting) as boolean,
      bodyPreview:     undefined,
      webLink:         ((e.WebLink ?? e.webLink) as string) || undefined,
    };
  });

  // Client-side search filter (always applied when search is set, in case server-side filter was not used)
  if (search) {
    const needle = search.toLowerCase();
    const beforeCount = events.length;
    events = events.filter(e =>
      e.subject?.toLowerCase().includes(needle) ||
      e.organizer?.toLowerCase().includes(needle) ||
      e.attendees.some(a => a.name.toLowerCase().includes(needle) || a.email.toLowerCase().includes(needle))
    );
    if (events.length < beforeCount) {
      progress(`🔍 Filtered ${beforeCount} → ${events.length} events matching "${search}"`);
    }
  }

  // Warn if we hit the $top limit — results may be truncated
  const rawCount = (data.value ?? []).length;
  if (rawCount >= top) {
    progress(`⚠️ Returned ${rawCount} events (hit limit) — narrow the date range or add a search filter for complete results`);
  }

  progress(`✅ Found ${events.length} calendar event(s)`);
  return events;
}

// ---------------------------------------------------------------------------
// Email / messages
// ---------------------------------------------------------------------------

export interface EmailMessage {
  id: string;
  subject: string;
  from: string;
  fromAddress: string;
  receivedDateTime: string;
  bodyPreview: string;
  body?: string;       // full body (plain text, stripped from HTML) — only when full_body requested
  isRead: boolean;
  hasAttachments: boolean;
}

export async function getEmails(opts: {
  search?: string;
  folder?: string;
  top?: number;
  unreadOnly?: boolean;
  fullBody?: boolean;
}, progress: ProgressFn = () => {}): Promise<EmailMessage[]> {
  const { search, folder: rawFolder, top = 25, unreadOnly, fullBody } = opts;
  await verifyOutlookConnection(progress);
  progress("📧 Fetching emails...");

  // When searching with no folder specified, search ALL mail (not just inbox).
  // When browsing (no search), default to inbox.
  const hasExplicitFolder = rawFolder !== undefined && rawFolder !== "";
  const folder = hasExplicitFolder
    ? await resolveMailFolder(rawFolder, progress)
    : null;

  // Use camelCase — works with both Graph API and Outlook REST (case-insensitive)
  const selectFields = fullBody
    ? "id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments,body"
    : "id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments";

  // Build folder path prefix: empty string = all mail, "/mailfolders/{id}" = specific folder
  // Outlook REST folder IDs are API-generated (base64-like) — do NOT encodeURIComponent
  // as it turns '=' into '%3D' which causes ErrorInvalidIdMalformed.
  const folderPrefix = folder ? `/mailfolders/${folder}` : "";

  let path: string;
  if (search) {
    const p = new URLSearchParams({
      $search: `"${search.replace(/"/g, "")}"`,
      $select: selectFields,
      $top: String(top),
    });
    // No folder = search across ALL folders (inbox, sent, custom client folders, etc.)
    path = `${folderPrefix}/messages?${p}`;
  } else {
    // Browsing without search — default to inbox if no folder specified
    const browsePrefix = folderPrefix || "/mailfolders/inbox";
    const filters: string[] = [];
    if (unreadOnly) filters.push("isRead eq false");
    const p = new URLSearchParams({
      $select: selectFields,
      $top: String(top),
      $orderby: "receivedDateTime desc",
      ...(filters.length ? { $filter: filters.join(" and ") } : {}),
    });
    path = `${browsePrefix}/messages?${p}`;
  }

  const token = await acquireOutlookRestToken(progress);
  const data = await outlookApiFetch(path, token, progress);
  // Handle both PascalCase (Outlook REST) and camelCase (Graph API) response fields
  const messages = (data.value as Record<string, unknown>[] ?? []).map(m => {
    // From field: REST = { EmailAddress: { Name, Address } }, Graph = { emailAddress: { name, address } }
    const fromRaw = (m.From ?? m.from) as { EmailAddress?: { Name: string; Address: string }; emailAddress?: { name: string; address: string } } | undefined;
    const fromEA = fromRaw?.EmailAddress
      ? { name: fromRaw.EmailAddress.Name, address: fromRaw.EmailAddress.Address }
      : fromRaw?.emailAddress
      ? { name: fromRaw.emailAddress.name, address: fromRaw.emailAddress.address }
      : undefined;
    // Body: REST = { Content, ContentType }, Graph = { content, contentType }
    const bodyRaw = (m.Body ?? m.body) as { Content?: string; ContentType?: string; content?: string; contentType?: string } | undefined;
    const bodyHtml = bodyRaw?.Content ?? bodyRaw?.content;
    const bodyType = (bodyRaw?.ContentType ?? bodyRaw?.contentType ?? "").toLowerCase();
    const bodyText = bodyHtml
      ? (bodyType === "html" ? stripHtml(bodyHtml) : bodyHtml)
      : undefined;
    return {
      id:               ((m.Id ?? m.id) as string) || "",
      subject:          ((m.Subject ?? m.subject) as string) || "",
      from:             fromEA?.name || "",
      fromAddress:      fromEA?.address || "",
      receivedDateTime: ((m.ReceivedDateTime ?? m.receivedDateTime) as string) || "",
      bodyPreview:      ((m.BodyPreview ?? m.bodyPreview) as string) || "",
      ...(bodyText !== undefined ? { body: bodyText } : {}),
      isRead:           ((m.IsRead ?? m.isRead) as boolean) ?? false,
      hasAttachments:   ((m.HasAttachments ?? m.hasAttachments) as boolean) ?? false,
    };
  });

  progress(`✅ Found ${messages.length} message(s)`);
  return messages;
}

// ---------------------------------------------------------------------------
// Mail folders — list all folders so users can browse custom client folders
// ---------------------------------------------------------------------------

export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId: string;
  childFolderCount: number;
  totalItemCount: number;
  unreadItemCount: number;
}

export async function listMailFolders(progress: ProgressFn = () => {}): Promise<MailFolder[]> {
  await verifyOutlookConnection(progress);
  progress("📁 Fetching mail folders...");
  const token = await acquireOutlookRestToken(progress);
  const SELECT = "$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount&$top=100";
  const data = await outlookApiFetch(`/mailfolders?${SELECT}`, token, progress);

  const mapFolder = (f: Record<string, unknown>): MailFolder => ({
    id:               (f.Id ?? f.id) as string,
    displayName:      (f.DisplayName ?? f.displayName) as string,
    parentFolderId:   (f.ParentFolderId ?? f.parentFolderId) as string,
    childFolderCount: (f.ChildFolderCount ?? f.childFolderCount ?? 0) as number,
    totalItemCount:   (f.TotalItemCount ?? f.totalItemCount ?? 0) as number,
    unreadItemCount:  (f.UnreadItemCount ?? f.unreadItemCount ?? 0) as number,
  });

  const folders = (data.value as Record<string, unknown>[] ?? []).map(mapFolder);

  // Recursively fetch child folders (Clients → PMI - Philip Morris → …)
  const MAX_DEPTH = 4;
  const fetchChildren = async (parents: MailFolder[], depth: number) => {
    if (depth >= MAX_DEPTH) return;
    const withChildren = parents.filter(f => f.childFolderCount > 0);
    for (const parent of withChildren) {
      try {
        const childData = await outlookApiFetch(
          `/mailfolders/${parent.id}/childfolders?${SELECT}`,
          token, progress
        );
        const children = (childData.value as Record<string, unknown>[] ?? []).map(mapFolder);
        folders.push(...children);
        // Recurse into children that also have subfolders
        await fetchChildren(children, depth + 1);
      } catch (e) {
        process.stderr.write(`[alfred:warn] child folder fetch failed for "${parent.displayName}" (${parent.id}): ${e instanceof Error ? e.message : String(e)}\n`);
      }
    }
  };
  await fetchChildren(folders, 0);

  progress(`✅ Found ${folders.length} mail folder(s)`);
  return folders;
}

/** Normalize folder name for matching: collapse whitespace, normalize dashes (em/en dash → hyphen), lowercase */
function normFolder(s: string): string {
  return s.toLowerCase().replace(/[\u2013\u2014]/g, "-").replace(/\s+/g, " ").trim();
}

/** Resolve a folder name to its Outlook REST folder ID. Tries well-known names first, then searches user folders. */
export async function resolveMailFolder(folder: string, progress: ProgressFn = () => {}): Promise<string> {
  const wellKnown = ["inbox", "sentitems", "drafts", "deleteditems", "junkemail", "archive"];
  if (wellKnown.includes(folder.toLowerCase())) return folder.toLowerCase();

  progress(`📁 Looking up folder "${folder}"...`);
  const folders = await listMailFolders(progress);
  const needle = normFolder(folder);

  // 1. Exact match (normalized)
  const exact = folders.find(f => normFolder(f.displayName) === needle);
  if (exact) {
    progress(`📁 Exact match: "${exact.displayName}" (${exact.totalItemCount} items)`);
    return exact.id;
  }

  // 2. Partial: folder name contained in display name (e.g. "PMI" matches "PMI - Philip Morris")
  const partial = folders.find(f => normFolder(f.displayName).includes(needle));
  if (partial) {
    progress(`📁 Partial match: "${partial.displayName}" (${partial.totalItemCount} items)`);
    return partial.id;
  }

  // 3. Reverse partial: display name contained in folder name (e.g. "PMI - Philip Morris International"
  //    matches a folder named "PMI - Philip Morris")
  const reverse = folders.find(f => needle.includes(normFolder(f.displayName)));
  if (reverse) {
    progress(`📁 Reverse match: "${reverse.displayName}" (${reverse.totalItemCount} items)`);
    return reverse.id;
  }

  // 4. Token match: split on " - " and match first segment (e.g. "PMI - Philip Morris" → "PMI")
  const firstToken = needle.split(" - ")[0]?.trim();
  if (firstToken && firstToken !== needle) {
    const tokenMatch = folders.find(f => normFolder(f.displayName).includes(firstToken));
    if (tokenMatch) {
      progress(`📁 Token match on "${firstToken}": "${tokenMatch.displayName}" (${tokenMatch.totalItemCount} items)`);
      return tokenMatch.id;
    }
  }

  progress(`⚠️ No folder match for "${folder}" — available: ${folders.map(f => f.displayName).slice(0, 10).join(", ")}`);
  throw new Error(
    `Mail folder "${folder}" not found.\n` +
    `Available folders: ${folders.map(f => f.displayName).join(", ")}\n` +
    `Use list_mail_folders to see all available folders.`
  );
}

