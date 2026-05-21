#!/usr/bin/env node
/**
 * One-off script: dump Dynamics field metadata for opportunity + sn_engagement.
 * Run with: node scripts/dump-schema.mjs
 * Requires Alfred Chrome window to be running (uses existing cookie auth).
 */

// Reuse Alfred's auth + fetch
const { getAuthCookies } = await import("../dist/auth/tokenExtractor.js");
const { readFileSync } = await import("fs");
const { homedir } = await import("os");

const cfgPath = `${homedir()}/.alfred-config.json`;
const cfg = JSON.parse(readFileSync(cfgPath, "utf-8"));
const BASE = `${cfg.dynamicsUrl}/api/data/v9.2`;

async function fetchJson(path) {
  const cookie = await getAuthCookies(() => {});
  const res = await fetch(`${BASE}${path}`, {
    headers: {
      Cookie: cookie,
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    },
    signal: AbortSignal.timeout(30_000),
  });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}: ${await res.text().catch(() => "")}`);
  return res.json();
}

async function dumpEntity(logicalName, label) {
  console.log(`\n${"=".repeat(70)}`);
  console.log(`  ${label} — entity: ${logicalName}`);
  console.log(`${"=".repeat(70)}\n`);

  const data = await fetchJson(
    `/EntityDefinitions(LogicalName='${logicalName}')/Attributes` +
    `?$select=LogicalName,DisplayName,AttributeType,RequiredLevel,IsValidForCreate,IsValidForUpdate,IsValidForRead,Description` +
    `&$filter=IsValidForRead eq true` +
    `&$orderby=LogicalName`
  );

  const attrs = data.value ?? [];

  // Filter to fields we care about + all sn_ custom fields
  const interesting = attrs.filter(a => {
    const n = a.LogicalName;
    return n.startsWith("sn_") || n.startsWith("msdyn_") ||
           ["name", "description", "totalamount", "estimatedclosedate", "stepname",
            "closeprobability", "statuscode", "statecode", "opportunityid",
            "_ownerid_value", "_accountid_value", "createdon", "modifiedon"].includes(n);
  });

  console.log(`Total fields: ${attrs.length} | SN/relevant fields shown: ${interesting.length}\n`);
  console.log("Field".padEnd(45), "Type".padEnd(18), "Required".padEnd(12), "Create".padEnd(8), "Update".padEnd(8), "Display Name");
  console.log("-".repeat(130));

  for (const a of interesting) {
    const displayName = a.DisplayName?.UserLocalizedLabel?.Label ?? "";
    const required = a.RequiredLevel?.Value ?? "None";
    console.log(
      a.LogicalName.padEnd(45),
      (a.AttributeType ?? "").padEnd(18),
      required.padEnd(12),
      (a.IsValidForCreate ? "Yes" : "No").padEnd(8),
      (a.IsValidForUpdate ? "Yes" : "No").padEnd(8),
      displayName
    );
  }
}

try {
  await dumpEntity("opportunity", "OPPORTUNITY");
  await dumpEntity("sn_engagement", "ENGAGEMENT");
  console.log("\nDone.");
} catch (e) {
  console.error("Error:", e.message);
  process.exit(1);
}
