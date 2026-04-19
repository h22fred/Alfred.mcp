import { readFileSync, writeFileSync, existsSync, unlinkSync } from "fs";
import { homedir } from "os";
import { join } from "path";

const CACHE_FILE = join(homedir(), ".alfred-auth-cache.json");

interface CacheEntry {
  value: string;
  expiresAt: number;
}

type CacheKey = "dynamics" | "outlook" | "graphToken" | "teamsGraphToken" | "teamsSkypeToken" | "outlookRestToken";

interface CacheFile {
  dynamics?: CacheEntry;
  outlook?: CacheEntry;
  graphToken?: CacheEntry;
  teamsGraphToken?: CacheEntry;
  teamsSkypeToken?: CacheEntry;
  outlookRestToken?: CacheEntry;
}

function readCache(): CacheFile {
  try {
    if (existsSync(CACHE_FILE)) {
      return JSON.parse(readFileSync(CACHE_FILE, "utf8")) as CacheFile;
    }
  } catch (e) {
    process.stderr.write(`[alfred:warn] auth cache read failed: ${e instanceof Error ? e.message : String(e)}\n`);
  }
  return {};
}

function writeCache(cache: CacheFile): void {
  try {
    writeFileSync(CACHE_FILE, JSON.stringify(cache, null, 2), { mode: 0o600 });
  } catch (e) {
    process.stderr.write(`[alfred:warn] auth cache write failed: ${e instanceof Error ? e.message : String(e)}\n`);
  }
}

export function loadCachedAuth(key: CacheKey): CacheEntry | null {
  const cache = readCache();
  const entry = cache[key];
  if (entry && Date.now() < entry.expiresAt) return entry;
  return null;
}

export function saveCachedAuth(key: CacheKey, value: string, expiresAt: number): void {
  const cache = readCache();
  cache[key] = { value, expiresAt };
  writeCache(cache);
}

export function clearCachedAuthFile(key?: CacheKey): void {
  if (!key) {
    try { unlinkSync(CACHE_FILE); } catch { /* file may not exist */ }
    return;
  }
  const cache = readCache();
  delete cache[key];
  writeCache(cache);
}
