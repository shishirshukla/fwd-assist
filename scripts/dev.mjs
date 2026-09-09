#!/usr/bin/env node

import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const port = process.env.PORT || "43123";
const require = createRequire(import.meta.url);
const nextCli = require.resolve("next/dist/bin/next");

process.env.WATCHPACK_POLLING = process.env.WATCHPACK_POLLING || "true";
process.env.CHOKIDAR_USEPOLLING = process.env.CHOKIDAR_USEPOLLING || "true";

console.log(`Starting Next.js on http://127.0.0.1:${port}`);
console.log("Wait for a Ready / Local line. First compile can take a minute.");
console.log("If this hangs on /mnt/c, copy the repo to ~/src/fwd-assist and run from there.");

const child = spawn(
  process.execPath,
  [nextCli, "dev", "--hostname", "127.0.0.1", "--port", String(port)],
  { stdio: "inherit", cwd: root, env: process.env },
);

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exit(code ?? 1);
});
