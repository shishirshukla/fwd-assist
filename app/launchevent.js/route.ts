import { readFileSync } from "node:fs";
import { join } from "node:path";

import { getPublicBaseUrl } from "@/lib/manifest";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

export function GET() {
  const source = readFileSync(
    /* turbopackIgnore: true */ join(process.cwd(), "outlook", "launchevent.js"),
    "utf8",
  );
  // A single self-contained script is required by classic Outlook on Windows.
  const config = `var EMAIL_TO_LMS_BASE_URL = ${JSON.stringify(getPublicBaseUrl())};\n`;
  return new Response(config + source, {
    headers: {
      "Content-Type": "application/javascript; charset=utf-8",
      "Cache-Control": "no-store, max-age=0",
    },
  });
}
