#!/usr/bin/env node

import { assertManifestConfigured, getPublicBaseUrl, resolvePublicBaseUrl } from "./manifest-utils.mjs";

try {
  assertManifestConfigured();
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
}

const configured = resolvePublicBaseUrl();
console.log(
  configured
    ? `Manifest will be served at ${configured}/manifest.xml`
    : `Manifest will be served at ${getPublicBaseUrl()}/manifest.xml (local default)`,
);
