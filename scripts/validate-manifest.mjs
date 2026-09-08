#!/usr/bin/env node
/**
 * Validate public/manifest.xml with Microsoft's office-addin-manifest tool.
 */

import { spawnSync } from "node:child_process";
import { resolve } from "node:path";

const manifest = resolve("public/manifest.xml");
const result = spawnSync(
  "npx",
  ["--yes", "office-addin-manifest", "validate", manifest],
  { stdio: "inherit", shell: true },
);

process.exit(result.status ?? 1);
