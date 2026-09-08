#!/usr/bin/env node
/**
 * Rewrite manifest URLs for an HTTPS host and verify endpoints respond.
 *
 * Usage:
 *   node scripts/set-manifest-url.mjs https://abc123.ngrok-free.app
 */

import { normalizeBaseUrl, writeManifestForBase } from "./manifest-utils.mjs";

async function checkUrl(url) {
  try {
    const response = await fetch(url, { redirect: "follow" });
    return { url, ok: response.ok, status: response.status };
  } catch (error) {
    return { url, ok: false, status: String(error) };
  }
}

const nextBase = normalizeBaseUrl(process.argv[2] || "");
const result = writeManifestForBase(nextBase);

console.log(`Wrote public/manifest.xml`);
console.log(`  Base URL : ${nextBase}`);
console.log(`  AppDomain: ${result.host}`);
console.log(`\nFor Railway, set PUBLIC_BASE_URL=${nextBase} and use https://${result.host}/manifest.xml`);

const checks = [
  `${nextBase}/icons/icon-64.png`,
  `${nextBase}/icons/icon-128.png`,
  `${nextBase}/taskpane.html`,
  `${nextBase}/commands.html`,
  `${nextBase}/launchevent.js`,
  `${nextBase}/manifest.xml`,
];

console.log("\nPreflight checks:");
let failed = 0;
for (const url of checks) {
  const check = await checkUrl(url);
  const label = check.ok ? "OK" : "FAIL";
  console.log(`  [${label}] ${url}${check.ok ? "" : ` (${check.status})`}`);
  if (!check.ok) failed += 1;
}

if (failed > 0) {
  console.error("\nOne or more URLs are not reachable.");
  process.exit(1);
}
