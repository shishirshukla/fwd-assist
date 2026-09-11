import { readLogFileText, readStoredLogs, type LogLevel } from "@/lib/app-log";

export const dynamic = "force-dynamic";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

export function OPTIONS() {
  return new Response(null, { status: 204, headers: cors });
}

export async function GET(request: Request) {
  const url = new URL(request.url);
  const format = (url.searchParams.get("format") || "json").toLowerCase();
  const level = url.searchParams.get("level") as LogLevel | null;
  const source = url.searchParams.get("source");
  const limitRaw = Number(url.searchParams.get("limit") || "200");
  const limit = Number.isFinite(limitRaw)
    ? Math.min(Math.max(limitRaw, 1), 1000)
    : 200;

  if (format === "text" || format === "txt") {
    return new Response(readLogFileText() || "No logs yet.\n", {
      status: 200,
      headers: {
        ...cors,
        "Content-Type": "text/plain; charset=utf-8",
      },
    });
  }

  let logs = readStoredLogs().slice().reverse();
  if (level === "info" || level === "warn" || level === "error") {
    logs = logs.filter((entry) => entry.level === level);
  }
  if (source) {
    logs = logs.filter((entry) => entry.source === source);
  }

  return Response.json(
    {
      count: logs.length,
      logs: logs.slice(0, limit),
    },
    { headers: cors },
  );
}
