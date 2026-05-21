import { readFileSync, writeFileSync, chmodSync, existsSync, unlinkSync } from "fs";
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

let _memCache: CacheFile | null = null;

function readCache(): CacheFile {
  if (_memCache !== null) return _memCache;
  try {
    if (existsSync(CACHE_FILE)) {
      _memCache = JSON.parse(readFileSync(CACHE_FILE, "utf8")) as CacheFile;
      return _memCache;
    }
  } catch (e) {
    process.stderr.write(`[alfred:warn] auth cache read failed: ${e instanceof Error ? e.message : String(e)}\n`);
  }
  _memCache = {};
  return _memCache;
}

function writeCache(cache: CacheFile): void {
  _memCache = cache;
  try {
    writeFileSync(CACHE_FILE, JSON.stringify(cache, null, 2), { mode: 0o600 });
    // chmodSync ensures perms are enforced even if the file already existed
    chmodSync(CACHE_FILE, 0o600);
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
    _memCache = null;
    try { unlinkSync(CACHE_FILE); } catch { /* file may not exist */ }
    return;
  }
  const cache = readCache();
  delete cache[key];
  writeCache(cache);
}
