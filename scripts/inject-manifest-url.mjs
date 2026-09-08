#!/usr/bin/env node
/**
 * Inject Railway / production URLs into public/manifest.xml at build time.
 * Skips when no PUBLIC_BASE_URL or RAILWAY_PUBLIC_DOMAIN is set (local dev).
 */

import { DEFAULT_BASE, resolvePublicBaseUrl, writeManifestForBase } from "./manifest-utils.mjs";

const base = resolvePublicBaseUrl();

if (!base) {
  console.log(
    `Manifest inject skipped (no PUBLIC_BASE_URL or RAILWAY_PUBLIC_DOMAIN). Using ${DEFAULT_BASE} in source manifest.`,
  );
  process.exit(0);
}

const result = writeManifestForBase(base);
console.log(
  result.changed
    ? `Injected manifest URLs for ${result.nextBase} (AppDomain: ${result.host})`
    : `Manifest already uses ${result.nextBase}`,
);
