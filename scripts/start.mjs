#!/usr/bin/env node
/**
 * Start the app on Windows, macOS, and Linux.
 * Uses next start when a production build exists; otherwise next dev.
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
const port = process.env.PORT || "43123";
const hasBuild = existsSync(join(root, ".next", "BUILD_ID"));

try {
  assertManifestConfigured();
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
}

if (!existsSync(join(root, "node_modules", "next"))) {
  console.error("Dependencies are missing. In this folder run:\n  npm install");
  process.exit(1);
}

const configured = resolvePublicBaseUrl();
console.log(
  configured
    ? `Manifest: ${configured}/manifest.xml`
    : `Manifest: ${getPublicBaseUrl()}/manifest.xml (local default)`,
);

const require = createRequire(import.meta.url);
const nextCli = require.resolve("next/dist/bin/next");
const mode = hasBuild ? "start" : "dev";

if (!hasBuild) {
  console.log("No production build (.next) yet — starting development server.");
  console.log("For a production server later:  npm run build   then   npm start");
}

console.log(`Open http://localhost:${port}`);

const child = spawn(
  process.execPath,
  [nextCli, mode, "--hostname", "0.0.0.0", "--port", String(port)],
  { stdio: "inherit", cwd: root, env: process.env },
);

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exit(code ?? 1);
});
