#!/usr/bin/env node
/**
 * Rewrite manifest URLs for an HTTPS tunnel and verify endpoints respond.
 *
 * Usage:
 *   node scripts/set-manifest-url.mjs https://abc123.ngrok-free.app
 */

import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

const DEFAULT_BASE = "https://localhost:43123";
const manifestPath = resolve("public/manifest.xml");

function normalizeBaseUrl(input) {
  const trimmed = input.trim().replace(/\/+$/, "");
  if (!/^https:\/\/.+/i.test(trimmed)) {
    throw new Error("Base URL must start with https:// (Outlook requires HTTPS).");
  }
  return trimmed;
}

function hostFromUrl(url) {
  return new URL(url).hostname;
}

function replaceManifestUrls(xml, nextBase) {
  const host = hostFromUrl(nextBase);
  let updated = xml.replace(/https:\/\/[^"<]+/g, (match) => {
    const path = match.replace(/^https:\/\/[^/]+/, "");
    return nextBase + (path || "");
  });
  updated = updated.replace(
    /<AppDomain>[^<]+<\/AppDomain>/,
    `<AppDomain>${host}</AppDomain>`,
  );
  return updated;
}

async function checkUrl(url) {
  try {
    const response = await fetch(url, { redirect: "follow" });
    return { url, ok: response.ok, status: response.status };
  } catch (error) {
    return { url, ok: false, status: String(error) };
  }
}

const nextBase = normalizeBaseUrl(process.argv[2] || "");
const current = readFileSync(manifestPath, "utf8");
const updated = replaceManifestUrls(current, nextBase);

if (updated === current) {
  console.log(`Manifest already points at ${nextBase}`);
} else {
  writeFileSync(manifestPath, updated, "utf8");
  console.log(`Updated public/manifest.xml`);
  console.log(`  Base URL : ${nextBase}`);
  console.log(`  AppDomain: ${hostFromUrl(nextBase)}`);
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
    "\nOne or more URLs are not reachable from this machine. Outlook will also fail to install until ngrok and npm start are running.",
  );
  process.exit(1);
}

console.log("\nUpload public/manifest.xml in Outlook Web → My add-ins → Add from file.");
