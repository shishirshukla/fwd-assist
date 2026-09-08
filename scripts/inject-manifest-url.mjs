#!/usr/bin/env node
/**
 * Inject Railway / production URLs into a generated manifest file.
 * The live app serves /manifest.xml dynamically; this script is for offline download.
 */

import { assertManifestConfigured, getPublicBaseUrl, writeManifestForBase } from "./manifest-utils.mjs";

assertManifestConfigured();

const result = writeManifestForBase(getPublicBaseUrl());
console.log(`Wrote public/manifest.xml for ${result.nextBase} (AppDomain: ${result.host})`);
