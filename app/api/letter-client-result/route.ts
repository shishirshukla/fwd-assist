import { appendLog } from "@/lib/app-log";
import { updateCaptureRemotePush } from "@/lib/forward-capture";

export const dynamic = "force-dynamic";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

export function OPTIONS() {
  return new Response(null, { status: 204, headers: cors });
}

export async function POST(request: Request) {
  let body: {
    captureId?: string;
    ok?: boolean;
    status?: number | null;
    error?: string | null;
    url?: string;
    letterId?: string;
    headerSet?: boolean;
    responseText?: string;
    opaque?: boolean;
  };
  try {
    body = (await request.json()) as typeof body;
  } catch {
    return Response.json({ error: "Request body must be JSON." }, { status: 400, headers: cors });
  }

  const captureId = String(body.captureId || "");
  const letterId = body.letterId ? String(body.letterId) : "";
  const remotePush = {
    attempted: true,
    ok: Boolean(body.ok),
    status: typeof body.status === "number" ? body.status : null,
    error: body.error ? String(body.error) : null,
    url: body.url,
    source: "browser",
    letterId: letterId || undefined,
    headerSet: Boolean(body.headerSet),
    responseText: body.opaque
      ? "browser-opaque: request sent; response not readable (CORS)."
      : String(body.responseText || ""),
  };

  const updated = captureId ? updateCaptureRemotePush(captureId, remotePush) : null;

  appendLog({
    level: remotePush.ok ? "info" : "error",
    source: "letter-submit",
    message: remotePush.ok
      ? `Browser letter API result for ${captureId || "(no id)"}: letterId=${letterId || "(none)"} header=${remotePush.headerSet ? "X-LETTERID-CGB set" : "not set"} ${remotePush.responseText || "(empty body)"}`
      : `Browser letter API error for ${captureId || "(no id)"}: ${remotePush.error || "unknown error"}`,
    captureId: captureId || undefined,
    details: {
      url: remotePush.url,
      status: remotePush.status,
      ok: remotePush.ok,
      letterId,
      headerSet: remotePush.headerSet,
      opaque: Boolean(body.opaque),
      error: remotePush.error,
      responseBody: remotePush.responseText,
    },
  });

  return Response.json({ ok: true, updated: Boolean(updated) }, { status: 200, headers: cors });
}
