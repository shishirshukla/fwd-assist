import {
  appendCapture,
  newCaptureId,
  normalizeCapture,
  pushCaptureIfConfigured,
  readStoredCaptures,
  type ForwardCaptureInput,
} from "@/lib/forward-capture";

export const dynamic = "force-dynamic";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

export function OPTIONS() {
  return new Response(null, { status: 204, headers: cors });
}

export async function GET() {
  const captures = readStoredCaptures();
  return Response.json(
    {
      count: captures.length,
      captures: captures.slice().reverse(),
      pushUrlConfigured: Boolean(process.env.CAPTURE_PUSH_URL?.trim()),
    },
    { headers: cors },
  );
}

export async function POST(request: Request) {
  let body: ForwardCaptureInput;
  try {
    body = (await request.json()) as ForwardCaptureInput;
  } catch {
    return Response.json(
      { error: "Request body must be JSON." },
      { status: 400, headers: cors },
    );
  }

  const fields = normalizeCapture(body);
  if (
    !fields.subject &&
    !fields.messageBody &&
    fields.toEmailAddresses.length === 0 &&
    !fields.senderEmailId
  ) {
    return Response.json(
      { error: "Provide at least subject, body, sender, or recipients." },
      { status: 400, headers: cors },
    );
  }

  const draft = {
    id: newCaptureId(),
    capturedAt: new Date().toISOString(),
    ...fields,
    remotePush: {
      attempted: false,
      ok: false,
      status: null,
      error: null,
    },
  };

  draft.remotePush = await pushCaptureIfConfigured(draft);
  appendCapture(draft);

  return Response.json(
    { ok: true, id: draft.id, remotePush: draft.remotePush },
    { status: 201, headers: cors },
  );
}
