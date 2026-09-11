import { appendLog } from "@/lib/app-log";
import {
  dummyLetterSubmitFields,
  submitLetter,
  type LetterSubmitFields,
} from "@/lib/letter-submit";

export const dynamic = "force-dynamic";
export const maxDuration = 60;

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

export function OPTIONS() {
  return new Response(null, { status: 204, headers: cors });
}

async function runDummy(overrides: Partial<LetterSubmitFields> = {}) {
  const payload = dummyLetterSubmitFields(overrides);
  const started = Date.now();
  const result = await submitLetter(payload);
  const elapsedMs = Date.now() - started;

  appendLog({
    level: result.ok ? "info" : "error",
    source: "letter-dummy",
    message: result.ok
      ? `Dummy letter POST succeeded in ${elapsedMs}ms (${result.status})`
      : `Dummy letter POST failed in ${elapsedMs}ms: ${result.error || "unknown error"}`,
    details: {
      elapsedMs,
      url: result.url,
      status: result.status,
      ok: result.ok,
      error: result.error,
      payload,
      responseBody: result.responseText || "",
    },
  });

  return {
    ok: result.ok,
    elapsedMs,
    url: result.url,
    payload,
    status: result.status,
    error: result.error,
    responseText: result.responseText,
  };
}

export async function GET() {
  const report = await runDummy();
  return Response.json(report, {
    status: report.ok ? 200 : 502,
    headers: cors,
  });
}

export async function POST(request: Request) {
  let overrides: Partial<LetterSubmitFields> = {};
  const contentType = request.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    try {
      const body = (await request.json()) as Partial<LetterSubmitFields>;
      if (body && typeof body === "object") {
        overrides = body;
      }
    } catch {
      return Response.json(
        { error: "Request body must be JSON if Content-Type is application/json." },
        { status: 400, headers: cors },
      );
    }
  }

  const report = await runDummy(overrides);
  return Response.json(report, {
    status: report.ok ? 200 : 502,
    headers: cors,
  });
}
