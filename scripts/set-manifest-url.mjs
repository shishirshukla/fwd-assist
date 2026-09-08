#!/usr/bin/env node
/**
 * Rewrite every https://localhost:43123 URL in public/manifest.xml
 * to your public HTTPS base URL (e.g. an ngrok tunnel).
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

const nextBase = normalizeBaseUrl(process.argv[2] || "");
const current = readFileSync(manifestPath, "utf8");

let updated = current;
if (current.includes(nextBase)) {
  console.log(`Manifest already uses ${nextBase}`);
  process.exit(0);
}

if (current.includes(DEFAULT_BASE)) {
  updated = current.replaceAll(DEFAULT_BASE, nextBase);
} else {
  const match = current.match(/https:\/\/[^"<]+/);
  if (!match) {
    throw new Error("Could not find an existing https:// base URL in manifest.xml");
  }
  updated = current.replaceAll(match[0], nextBase);
}

writeFileSync(manifestPath, updated, "utf8");
console.log(`Updated public/manifest.xml → ${nextBase}`);
