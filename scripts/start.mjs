#!/usr/bin/env node
/**
 * Cross-platform production start (Windows cmd/PowerShell cannot expand ${PORT:-43123}).
 */

import { existsSync } from "node:fs";
import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import {
  assertManifestConfigured,
  getPublicBaseUrl,
  resolvePublicBaseUrl,
} from "./manifest-utils.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const nextBin = join(root, "node_modules", "next", "dist", "bin", "next");
const port = process.env.PORT || "43123";

try {
  assertManifestConfigured();
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
}

if (!existsSync(join(root, "node_modules"))) {
  console.error("Dependencies are missing. Run:  npm install");
  process.exit(1);
}

if (!existsSync(join(root, ".next"))) {
  console.error("No production build found. Run:  npm run build");
  console.error("Then run:                      npm start");
  console.error("For local UI work, use:         npm run dev");
  process.exit(1);
}

if (!existsSync(nextBin)) {
  console.error("Next.js is not installed. Run:  npm install");
  process.exit(1);
}

const configured = resolvePublicBaseUrl();
console.log(
  configured
    ? `Manifest: ${configured}/manifest.xml`
    : `Manifest: ${getPublicBaseUrl()}/manifest.xml (local default)`,
);
console.log(`Starting on http://0.0.0.0:${port}`);

const require = createRequire(import.meta.url);
const nextCli = require.resolve("next/dist/bin/next");

const child = spawn(
  process.execPath,
  [nextCli, "start", "--hostname", "0.0.0.0", "-p", String(port)],
  { stdio: "inherit", cwd: root, env: process.env },
);

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exit(code ?? 1);
});
