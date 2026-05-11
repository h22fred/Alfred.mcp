#!/usr/bin/env node
/**
 * Analyzes sc/index.ts vs sales/index.ts — categorizes tools as:
 *   IDENTICAL     — safe to move to common-tools.ts as-is
 *   COSMETIC      — differ only in comments/whitespace/description wording (merge with best version)
 *   REAL DIFF     — genuinely different params, logic, or role-specific behavior (needs manual review)
 *   SC-ONLY       — stays in sc/index.ts
 *   SALES-ONLY    — stays in sales/index.ts
 *
 * Run: node scripts/analyze-tool-overlap.mjs [--verbose]
 */

import { readFileSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, "..");
const verbose = process.argv.includes("--verbose");

const scSrc   = readFileSync(join(root, "src/sc/index.ts"),    "utf8");
const salesSrc = readFileSync(join(root, "src/sales/index.ts"), "utf8");

// ── Extract tool blocks ────────────────────────────────────────────────────

function extractToolBlocks(src) {
  const tools = new Map();
  const re = /server\.tool\(\s*\n?\s*"([^"]+)"/g;
  let m;
  while ((m = re.exec(src)) !== null) {
    const name = m[1];
    let depth = 0, i = m.index, started = false;
    while (i < src.length) {
      if (src[i] === "(") { depth++; started = true; }
      if (src[i] === ")") depth--;
      if (started && depth === 0) { i++; break; }
      i++;
    }
    tools.set(name, src.slice(m.index, i).trim());
  }
  return tools;
}

// ── Normalisation passes ───────────────────────────────────────────────────

/** Structural normalisation: strips comments, collapses whitespace, unifies known variable aliases. */
function normaliseStructural(s) {
  return s
    .replace(/\r\n/g, "\n")
    // strip single-line // comments (but not URLs)
    .replace(/(?<!https?:)\/\/[^\n]*/g, "")
    // strip block comments
    .replace(/\/\*[\s\S]*?\*\//g, "")
    // unify known variable aliases
    .replace(/\bDYNAMICS_BASE_URL\b/g, "DYNAMICS_HOST")
    // collapse runs of whitespace/blank lines
    .replace(/[ \t]+/g, " ")
    .replace(/\n{2,}/g, "\n")
    // strip blank lines (lines with only whitespace after comment removal)
    .replace(/^\s*$/gm, "")
    .replace(/\n{2,}/g, "\n")
    .trim();
}

/** Surface normalisation: additionally strips description string differences (prose only). */
function normaliseSurface(s) {
  return normaliseStructural(s)
    // collapse template literal strings — keep only the variable references
    .replace(/`[^`]*`/g, "`…`")
    // collapse quoted description strings
    .replace(/"[^"]{30,}"/g, '"…"')
    .replace(/'[^']{30,}'/g, "'…'");
}

// ── Diff helpers ───────────────────────────────────────────────────────────

function lineDiff(a, b, ctx = 2) {
  const aL = a.split("\n"), bL = b.split("\n");
  const max = Math.max(aL.length, bL.length);
  const changed = new Set();
  for (let i = 0; i < max; i++) {
    if (aL[i] !== bL[i]) {
      for (let c = Math.max(0, i - ctx); c <= Math.min(max - 1, i + ctx); c++) changed.add(c);
    }
  }
  const out = []; let last = -2;
  for (const i of [...changed].sort((a, b) => a - b)) {
    if (i > last + 1) out.push("  ...");
    const a = aL[i] ?? "", b = bL[i] ?? "";
    if (a === b) out.push(`     ${a}`);
    else { if (a) out.push(`  SC - ${a.trim()}`); if (b) out.push(`SALES + ${b.trim()}`); }
    last = i;
  }
  return out.join("\n");
}

// ── Classify each tool ─────────────────────────────────────────────────────

const scTools    = extractToolBlocks(scSrc);
const salesTools = extractToolBlocks(salesSrc);
const scNames    = new Set(scTools.keys());
const salesNames = new Set(salesTools.keys());
const allNames   = new Set([...scNames, ...salesNames]);

const identical = [], cosmetic = [], realDiff = [], scOnly = [], salesOnly = [];

for (const name of [...allNames].sort()) {
  const inSc = scNames.has(name), inSales = salesNames.has(name);
  if (!inSc)    { salesOnly.push(name); continue; }
  if (!inSales) { scOnly.push(name);    continue; }

  const sc = scTools.get(name), sal = salesTools.get(name);

  if (normaliseStructural(sc) === normaliseStructural(sal)) {
    identical.push(name);
  } else if (normaliseSurface(sc) === normaliseSurface(sal)) {
    cosmetic.push({ name, sc, sales: sal });
  } else {
    realDiff.push({ name, sc, sales: sal });
  }
}

// ── Output ─────────────────────────────────────────────────────────────────

const W = 60;
const hr = "─".repeat(W);
const dhr = "═".repeat(W);

console.log(`\n╔${"═".repeat(W - 2)}╗`);
console.log(`║   Alfred SC vs Sales — Tool Overlap Report (v2)${" ".repeat(W - 50)}║`);
console.log(`╚${"═".repeat(W - 2)}╝\n`);

console.log(`✅ IDENTICAL — move to common-tools.ts as-is (${identical.length})`);
for (const n of identical) console.log(`   • ${n}`);

console.log(`\n🟡 COSMETIC DIFF — description/comment wording only (${cosmetic.length})`);
console.log(`   → Merge using SC version (tends to be more verbose/correct)`);
for (const { name } of cosmetic) console.log(`   • ${name}`);

console.log(`\n🔴 REAL DIFF — schema, params, or logic differs (${realDiff.length})`);
console.log(`   → Each needs manual review before merge`);
for (const { name } of realDiff) console.log(`   • ${name}`);

console.log(`\n🔵 SC-only — stays in sc/index.ts (${scOnly.length})`);
for (const n of scOnly) console.log(`   • ${n}`);

console.log(`\n🟠 SALES-only — stays in sales/index.ts (${salesOnly.length})`);
for (const n of salesOnly) console.log(`   • ${n}`);

const canMove = identical.length + cosmetic.length;
const lineEst = canMove * 20;
console.log(`\n${hr}`);
console.log(`SUMMARY`);
console.log(`  ${identical.length} identical + ${cosmetic.length} cosmetic = ${canMove} tools safe to move to common`);
console.log(`  ${realDiff.length} real diffs need review | ${scOnly.length} SC-only | ${salesOnly.length} Sales-only`);
console.log(`  Estimated lines eliminated: ~${lineEst}–${lineEst * 2}`);

if (realDiff.length > 0 && verbose) {
  console.log(`\n${dhr}`);
  console.log("REAL DIFF DETAILS (structural diff, comments stripped):\n");
  for (const { name, sc, sales } of realDiff) {
    console.log(`┌─ ${name} ${"─".repeat(Math.max(0, W - name.length - 4))}┐`);
    console.log(lineDiff(normaliseStructural(sc), normaliseStructural(sales)));
    console.log(`└${hr}┘\n`);
  }
} else if (realDiff.length > 0) {
  console.log(`\nRun with --verbose to see structural diffs for the ${realDiff.length} real-diff tools.`);
}

if (cosmetic.length > 0 && verbose) {
  console.log(`\n${dhr}`);
  console.log("COSMETIC DIFF DETAILS (what differs in wording):\n");
  for (const { name, sc, sales } of cosmetic) {
    console.log(`┌─ ${name} ${"─".repeat(Math.max(0, W - name.length - 4))}┐`);
    console.log(lineDiff(normaliseStructural(sc), normaliseStructural(sales)));
    console.log(`└${hr}┘\n`);
  }
}
