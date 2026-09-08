#!/usr/bin/env node
/**
 * Rewrite manifest URLs for an HTTPS tunnel and verify endpoints respond.
 *
 * Usage:
 *   node scripts/set-manifest-url.mjs https://abc123.ngrok-free.app
 */

import {
  normalizeBaseUrl,
  writeManifestForBase,
} from "./manifest-utils.mjs";

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

if (!result.changed) {
  console.log(`Manifest already points at ${nextBase}`);
} else {
  console.log(`Updated public/manifest.xml`);
  console.log(`  Base URL : ${nextBase}`);
  console.log(`  AppDomain: ${result.host}`);
}

const checks = [
  `${nextBase}/icons/icon-64.png`,
  `${nextBase}/icons/icon-128.png`,
  `${nextBase}/taskpane.html`,
  `${nextBase}/commands.html`,
  `${nextBase}/launchevent.js`,
];

console.log("\nPreflight checks:");
let failed = 0;
for (const url of checks) {
  const result = await checkUrl(url);
  const label = result.ok ? "OK" : "FAIL";
  console.log(`  [${label}] ${url}${result.ok ? "" : ` (${result.status})`}`);
  if (!result.ok) failed += 1;
}

if (failed > 0) {
  console.error(
    "\nOne or more URLs are not reachable. Outlook will fail to install until the host is live.",
  );
  process.exit(1);
}

console.log("\nUpload public/manifest.xml in Outlook Web → My add-ins → Add from file.");
