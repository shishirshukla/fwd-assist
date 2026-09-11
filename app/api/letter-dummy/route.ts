import { appendLog } from "@/lib/app-log";
import {
  dummyLetterSubmitFields,
  letterSubmitBaseUrl,
  submitLetter,
  type LetterSubmitFields,
} from "@/lib/letter-submit";
import {
  createLetterTracer,
  lookupDns,
  probeTcp,
  probeTls,
} from "@/lib/letter-trace";

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

function readCertHint(trace: { event: string; details?: Record<string, unknown> }[]) {
  const noVerify = [...trace].reverse().find((entry) => entry.event === "probe.tls.noverify.ok");
  const verifyFail = [...trace].reverse().find((entry) => entry.event === "probe.tls.verify.fail");
  const verifyOk = [...trace].reverse().find((entry) => entry.event === "probe.tls.verify.ok");
  const timedOut = trace.some((entry) => /timeout/i.test(entry.event));
  const sawCert = Boolean(
    noVerify?.details?.peerCertificate || verifyOk?.details?.peerCertificate,
  );

  if (sawCert && verifyFail) {
    return "TLS handshake completed and a peer certificate was received. Node rejected the certificate (not a TCP timeout).";
  }
  if (sawCert && verifyOk) {
    return "TLS handshake completed, peer certificate was received, and certificate verification succeeded.";
  }
  if (timedOut && !sawCert) {
    return "Timed out before a peer certificate was received. That is a handshake/TCP stall, not a certificate validation error.";
  }
  return "See tcpLogs for DNS, TCP, TLS, and HTTP socket events.";
}

async function runDummy(overrides: Partial<LetterSubmitFields> = {}) {
  const payload = dummyLetterSubmitFields(overrides);
  const tracer = createLetterTracer();
  const target = new URL(letterSubmitBaseUrl());
  const host = target.hostname;
  const port = Number(target.port || (target.protocol === "https:" ? 443 : 80));
  const probeMs = 8000;

  tracer.push("dummy.start", { url: target.toString(), host, port });

  try {
    const addresses = await lookupDns(host);
    tracer.push("dns.ok", { addresses });
  } catch (error) {
    tracer.push("dns.fail", {
      message: error instanceof Error ? error.message : String(error),
    });
  }

  try {
    await probeTcp(host, port, probeMs, tracer);
  } catch (error) {
    tracer.push("probe.tcp.caught", {
      message: error instanceof Error ? error.message : String(error),
    });
  }

  try {
    await probeTls(host, port, probeMs, false, tracer);
  } catch (error) {
    tracer.push("probe.tls.noverify.caught", {
      message: error instanceof Error ? error.message : String(error),
    });
  }

  try {
    await probeTls(host, port, probeMs, true, tracer);
  } catch (error) {
    tracer.push("probe.tls.verify.caught", {
      message: error instanceof Error ? error.message : String(error),
    });
  }

  const started = Date.now();
  const result = await submitLetter(payload, { tracer });
  const elapsedMs = Date.now() - started;
  const certificateHint = readCertHint(tracer.events);

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
      certificateHint,
      responseBody: result.responseText || "",
      tcpLogs: tracer.events,
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
    certificateHint,
    tcpLogs: tracer.events,
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
