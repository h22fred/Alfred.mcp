import { connectWithRetry } from "../auth/tokenExtractor.js";
import type { ProgressFn } from "../auth/tokenExtractor.js";
import { loadCachedAuth, saveCachedAuth, clearCachedAuthFile } from "../auth/authFileCache.js";
import { stripHtml, urlHostMatches } from "../shared.js";
import { sanitizeODataSearch } from "./dynamicsClient.js";

const CDP_PORT = 9222;
const OUTLOOK_ORIGIN = "https://outlook.office.com";

/** Match both legacy (outlook.office.com) and new (outlook.cloud.microsoft.com) Outlook domains */
function isOutlookUrl(url: string): boolean {
  return urlHostMatches(url, "outlook.office.com") || urlHostMatches(url, "outlook.cloud.microsoft.com");
}
const TOKEN_CACHE_FALLBACK_MS = 45 * 60 * 1000; // 45 min — fallback when MSAL doesn't report expiry
const TOKEN_REFRESH_MARGIN_MS = 5 * 60 * 1000;  // refresh 5 min before expiry

interface TokenCache {
  token: string;
  expiresAt: number;
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
    throw new Error("No page targets found in Alfred. Open Alfred.app and log into Teams or Outlook.");
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
        "Make sure you are logged into Outlook in the Alfred window."
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

      if (!capturedToken) throw new Error("Could not capture Graph token from Outlook.\nMake sure you are logged into Outlook in the Alfred window.");
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
            if (!aud.includes('outlook.office')) continue;
            const exp = Number(val.expiresOn || val.expires_on || val.extended_expires_on || 0);
            if (exp && exp * 1000 < Date.now()) continue;
            return { token: secret, expiresAt: exp ? exp * 1000 : 0 };
          }
        } catch {}
      }
    } catch {}
    return null;
  };
  return tryStorage(sessionStorage) || tryStorage(localStorage);
})()`;

async function acquireOutlookRestToken(progress: ProgressFn): Promise<string> {
  // Check in-memory cache
  if (outlookRestTokenCache && Date.now() < outlookRestTokenCache.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    const mins = Math.round((outlookRestTokenCache.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Outlook REST token (~${mins} min remaining)`);
    return outlookRestTokenCache.token;
  }

  // Check file cache
  const fileCached = loadCachedAuth("outlookRestToken");
  if (fileCached && Date.now() < fileCached.expiresAt - TOKEN_REFRESH_MARGIN_MS) {
    outlookRestTokenCache = { token: fileCached.value, expiresAt: fileCached.expiresAt };
    const mins = Math.round((fileCached.expiresAt - Date.now()) / 60_000);
    progress(`🔑 Using cached Outlook REST token (~${mins} min remaining)`);
    return fileCached.value;
  }

  progress("🔑 Acquiring Outlook REST token via CDP...");

  const listRes = await fetch(`http://localhost:${CDP_PORT}/json/list`);
  const targets = await listRes.json() as Array<{ webSocketDebuggerUrl?: string; type?: string; url?: string }>;

  // Outlook REST tokens live in the Outlook tab's MSAL cache
  const outlookTarget = targets.find(t => t.type === "page" && t.url && isOutlookUrl(t.url) && t.webSocketDebuggerUrl);
  const anyTarget = targets.find(t => t.type === "page" && t.webSocketDebuggerUrl);
  const target = outlookTarget ?? anyTarget;

  if (!target?.webSocketDebuggerUrl) {
    throw new Error("No page targets found in Alfred. Open Alfred.app and log into Outlook.");
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
  const msalResult = await new Promise<{ token: string; expiresAt: number } | null>((resolve) => {
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
        const msg = JSON.parse(event.data as string) as { result?: { result?: { value?: { token: string; expiresAt: number } | null } } };
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
    outlookRestTokenCache = { token: msalResult.token, expiresAt };
    saveCachedAuth("outlookRestToken", msalResult.token, expiresAt);
    const mins = Math.round((expiresAt - Date.now()) / 60_000);
    progress(`✅ Outlook REST token acquired from MSAL cache (~${mins} min valid)`);
    return msalResult.token;
  }

  // MSAL cache may be encrypted (outlook.cloud.microsoft.com uses AES-GCM).
  // Fallback: capture Bearer token from the Outlook tab's own API requests.
  progress("🔑 MSAL cache unreadable — capturing token via network interception...");

  const capturedToken = await new Promise<string | null>((resolve) => {
    const ws = new WebSocket(target.webSocketDebuggerUrl!);
    let found: string | null = null;

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
          // Network enabled — trigger API calls against all known Outlook API domains.
          // The page is on outlook.cloud.microsoft.com but API goes to outlook.microsoft.com.
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
          // Backup: if no token captured in 3s, reload page
          setTimeout(() => { if (!found) send("Page.reload"); }, 3_000);
        }
        if (msg.method === "Network.requestWillBeSent") {
          const headers = ((msg.params?.request as Record<string, unknown>)?.headers ?? {}) as Record<string, string>;
          const auth = headers["Authorization"] ?? headers["authorization"] ?? "";
          if (!found && auth.startsWith("Bearer ")) {
            found = auth.slice(7);
            clearTimeout(timer);
            try { ws.close(); } catch {}
            resolve(found);
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

  if (capturedToken) {
    const expiresAt = Date.now() + TOKEN_CACHE_FALLBACK_MS;
    outlookRestTokenCache = { token: capturedToken, expiresAt };
    saveCachedAuth("outlookRestToken", capturedToken, expiresAt);
    progress("✅ Outlook REST token captured via network interception");
    return capturedToken;
  }

  throw new Error(
    "Could not capture Outlook token.\n" +
    "Make sure Outlook is fully loaded (inbox visible) in the Alfred window, then retry."
  );
}

// outlook.microsoft.com is browser-only (CORS-gated); outlook.office.com works for server-side Bearer calls
const OUTLOOK_API = "https://outlook.office.com/api/v2.0/me";

// ---------------------------------------------------------------------------
// Outlook REST v2 fetch using Outlook REST Bearer token
// ---------------------------------------------------------------------------

async function outlookApiFetch(path: string, token: string, progress?: ProgressFn, _retryCount = 0): Promise<Record<string, unknown>> {
  const res = await fetch(`${OUTLOOK_API}${path}`, {
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
    const retry = await fetch(`${OUTLOOK_API}${path}`, {
      headers: { Authorization: `Bearer ${freshToken}`, Accept: "application/json" },
      signal: AbortSignal.timeout(30_000),
    });
    if (!retry.ok) {
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
  progress(`📅 Fetching calendar events ${startDate} → ${endDate}${search ? ` (filter: "${search}")` : ""}...`);

  // NOTE: /calendarview does NOT support $search — we filter client-side below.
  // Use Outlook REST API — same proven token path as emails.
  // Outlook REST v2.0 uses PascalCase field names.
  const params = new URLSearchParams({
    startDateTime: `${startDate}T00:00:00Z`,
    endDateTime:   `${endDate}T23:59:59Z`,
    $select: "Subject,Start,End,Location,Organizer,Attendees,IsOnlineMeeting,WebLink",
    $top: String(top),
    $orderby: "Start/DateTime",
  });

  if (search) {
    try {
      const safe = sanitizeODataSearch(search);
      params.set("$filter", `contains(Subject,'${safe}')`);
    } catch (e) {
      process.stderr.write(`[alfred:warn] OData search sanitize failed, using client-side filter: ${e instanceof Error ? e.message : String(e)}\n`);
    }
  }

  const token = await acquireOutlookRestToken(progress);

  let res = await fetch(`${OUTLOOK_API}/calendarview?${params}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
    signal: AbortSignal.timeout(30_000),
  });

  // If $filter is not supported, retry without it
  if (!res.ok && search && params.has("$filter")) {
    progress("⚠️ Server-side filter not supported on calendarView — falling back to client-side filter...");
    params.delete("$filter");
    res = await fetch(`${OUTLOOK_API}/calendarview?${params}`, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
      signal: AbortSignal.timeout(30_000),
    });
  }

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Calendar API ${res.status} ${res.statusText}${body ? `: ${body.slice(0, 200)}` : ""}`);
  }
  const data = await res.json() as { value?: Record<string, unknown>[] };

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
  progress("📧 Fetching emails...");

  // When searching with no folder specified, search ALL mail (not just inbox).
  // When browsing (no search), default to inbox.
  const hasExplicitFolder = rawFolder !== undefined && rawFolder !== "";
  const folder = hasExplicitFolder
    ? await resolveMailFolder(rawFolder, progress)
    : null;

  const selectFields = fullBody
    ? "Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments,Body"
    : "Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments";

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
    if (unreadOnly) filters.push("IsRead eq false");
    const p = new URLSearchParams({
      $select: selectFields,
      $top: String(top),
      $orderby: "ReceivedDateTime desc",
      ...(filters.length ? { $filter: filters.join(" and ") } : {}),
    });
    path = `${browsePrefix}/messages?${p}`;
  }

  const token = await acquireOutlookRestToken(progress);
  const data = await outlookApiFetch(path, token, progress);
  const messages = (data.value as Record<string, unknown>[] ?? []).map(m => {
    const fromEA = (m.From as { EmailAddress: { Name: string; Address: string } })?.EmailAddress;
    const bodyContent = (m.Body as { Content?: string; ContentType?: string } | undefined);
    const bodyText = bodyContent?.Content
      ? (bodyContent.ContentType === "html" || bodyContent.ContentType === "HTML"
          ? stripHtml(bodyContent.Content)
          : bodyContent.Content)
      : undefined;
    return {
      id:               m.Id as string,
      subject:          m.Subject as string,
      from:             fromEA?.Name || "",
      fromAddress:      fromEA?.Address || "",
      receivedDateTime: m.ReceivedDateTime as string,
      bodyPreview:      m.BodyPreview as string,
      ...(bodyText !== undefined ? { body: bodyText } : {}),
      isRead:           m.IsRead as boolean,
      hasAttachments:   m.HasAttachments as boolean,
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
  progress("📁 Fetching mail folders...");
  const token = await acquireOutlookRestToken(progress);
  const SELECT = "$select=Id,DisplayName,ParentFolderId,ChildFolderCount,TotalItemCount,UnreadItemCount&$top=100";
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

/** Resolve a folder name to its Outlook REST folder ID. Tries well-known names first, then searches user folders. */
export async function resolveMailFolder(folder: string, progress: ProgressFn = () => {}): Promise<string> {
  const wellKnown = ["inbox", "sentitems", "drafts", "deleteditems", "junkemail", "archive"];
  if (wellKnown.includes(folder.toLowerCase())) return folder.toLowerCase();

  progress(`📁 Looking up folder "${folder}"...`);
  // Search by display name across all folders including subfolders (case-insensitive)
  const folders = await listMailFolders(progress);
  const match = folders.find(f => f.displayName.toLowerCase() === folder.toLowerCase());
  if (match) {
    progress(`📁 Exact match: "${match.displayName}" (${match.totalItemCount} items)`);
    return match.id;
  }

  // Partial match fallback (e.g. "PMI" matches "PMI - Philip Morris")
  const partial = folders.find(f => f.displayName.toLowerCase().includes(folder.toLowerCase()));
  if (partial) {
    progress(`📁 Partial match: "${partial.displayName}" (${partial.totalItemCount} items)`);
    return partial.id;
  }

  // Return as-is — let Graph API error if invalid
  return folder;
}

