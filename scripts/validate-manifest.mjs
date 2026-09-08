#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { buildManifestXml } from "./manifest-utils.mjs";

const dir = mkdtempSync(join(tmpdir(), "manifest-validate-"));
const manifest = join(dir, "manifest.xml");
writeFileSync(manifest, buildManifestXml(), "utf8");

const result = spawnSync(
  "npx",
  ["--yes", "office-addin-manifest", "validate", manifest],
  { stdio: "inherit", shell: true },
);

process.exit(result.status ?? 1);
