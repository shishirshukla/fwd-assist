import dns from "node:dns";
import net from "node:net";
import tls from "node:tls";
import type { TLSSocket } from "node:tls";
import type { ClientRequest } from "node:http";
import type { PeerCertificate } from "node:tls";

import { appendLog } from "@/lib/app-log";

export type LetterTraceEvent = {
  ms: number;
  event: string;
  details?: Record<string, unknown>;
};

export type LetterTracer = {
  events: LetterTraceEvent[];
  startedAt: number;
  push: (event: string, details?: Record<string, unknown>) => void;
};

export function createLetterTracer(): LetterTracer {
  const startedAt = Date.now();
  const events: LetterTraceEvent[] = [];
  return {
    events,
    startedAt,
    push(event, details) {
      const entry: LetterTraceEvent = {
        ms: Date.now() - startedAt,
        event,
        ...(details && Object.keys(details).length > 0 ? { details } : {}),
      };
      events.push(entry);
      const parts = [
        `+${entry.ms}ms ${event}`,
        details ? JSON.stringify(details) : "",
      ].filter(Boolean);
      appendLog({
        level: "info",
        source: "letter-dummy-tcp",
        message: parts.join(" "),
        details: details || {},
      });
    },
  };
}

function summarizeCert(cert: PeerCertificate | undefined, depth = 0): Record<string, unknown> | null {
  if (!cert || typeof cert !== "object" || !("subject" in cert) || depth > 6) {
    return null;
  }
  const issuer =
    "issuerCertificate" in cert &&
    cert.issuerCertificate &&
    cert.issuerCertificate !== cert
      ? summarizeCert(cert.issuerCertificate as PeerCertificate, depth + 1)
      : null;
  return {
    subject: cert.subject,
    issuer: cert.issuer,
    validFrom: cert.valid_from,
    validTo: cert.valid_to,
    fingerprint256: cert.fingerprint256,
    serialNumber: cert.serialNumber,
    subjectAltName: cert.subjectaltname,
    ...(issuer ? { issuerCertificate: issuer } : {}),
  };
}

export function summarizeTlsSocket(socket: TLSSocket): Record<string, unknown> {
  let peer: Record<string, unknown> | null = null;
  try {
    peer = summarizeCert(socket.getPeerCertificate(true));
  } catch (error) {
    peer = { error: error instanceof Error ? error.message : String(error) };
  }
  return {
    authorized: socket.authorized,
    authorizationError:
      socket.authorizationError instanceof Error
        ? socket.authorizationError.message
        : socket.authorizationError
          ? String(socket.authorizationError)
          : null,
    protocol: socket.getProtocol(),
    cipher: socket.getCipher(),
    alpn: socket.alpnProtocol,
    servername: socket.servername,
    remoteAddress: socket.remoteAddress,
    remotePort: socket.remotePort,
    localAddress: socket.localAddress,
    localPort: socket.localPort,
    peerCertificate: peer,
  };
}

export function attachSocketTrace(socket: net.Socket, tracer: LetterTracer, label: string) {
  const bump = (event: string, details?: Record<string, unknown>) => {
    tracer.push(`${label}.${event}`, details);
  };

  bump("assigned", {
    remoteAddress: socket.remoteAddress,
    remotePort: socket.remotePort,
    localAddress: socket.localAddress,
    localPort: socket.localPort,
    connecting: socket.connecting,
    encrypted: "encrypted" in socket ? Boolean((socket as TLSSocket).encrypted) : false,
  });

  socket.on("lookup", (error, address, family, host) => {
    bump("lookup", {
      error: error ? error.message : null,
      address,
      family,
      host,
    });
  });
  socket.on("connect", () => {
    bump("connect", {
      remoteAddress: socket.remoteAddress,
      remotePort: socket.remotePort,
      localAddress: socket.localAddress,
      localPort: socket.localPort,
    });
  });
  socket.on("ready", () => bump("ready"));
  socket.on("secureConnect", () => {
    if ("getProtocol" in socket) {
      bump("secureConnect", summarizeTlsSocket(socket as TLSSocket));
    } else {
      bump("secureConnect");
    }
  });
  socket.on("timeout", () => bump("timeout"));
  socket.on("error", (error) => {
    const err = error as NodeJS.ErrnoException;
    bump("error", {
      name: err.name,
      message: err.message,
      code: err.code,
      errno: err.errno,
      syscall: err.syscall,
    });
  });
  socket.on("end", () => bump("end"));
  socket.on("close", (hadError) => bump("close", { hadError }));
  socket.on("data", (chunk) => {
    const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    bump("data", {
      bytes: buf.length,
      headHex: buf.subarray(0, 32).toString("hex"),
    });
  });
}

export function attachRequestTrace(req: ClientRequest, tracer: LetterTracer) {
  tracer.push("http.request.created", {
    method: req.method,
    path: req.path,
    host: req.host,
  });
  req.on("socket", (socket) => {
    attachSocketTrace(socket, tracer, "tcp");
  });
  req.on("continue", () => tracer.push("http.continue"));
  req.on("information", (info) => {
    tracer.push("http.information", { statusCode: info.statusCode, headers: info.headers });
  });
  req.on("response", (res) => {
    tracer.push("http.response", {
      statusCode: res.statusCode,
      headers: res.headers,
    });
  });
  req.on("timeout", () => tracer.push("http.timeout"));
  req.on("abort", () => tracer.push("http.abort"));
  req.on("close", () => tracer.push("http.close"));
  req.on("finish", () => tracer.push("http.finish"));
  req.on("error", (error) => {
    const err = error as NodeJS.ErrnoException;
    tracer.push("http.error", {
      name: err.name,
      message: err.message,
      code: err.code,
    });
  });
}

export function lookupDns(host: string): Promise<dns.LookupAddress[]> {
  return new Promise((resolve, reject) => {
    dns.lookup(host, { all: true }, (error, addresses) => {
      if (error) {
        reject(error);
        return;
      }
      resolve(addresses);
    });
  });
}

export function probeTcp(
  host: string,
  port: number,
  timeoutMs: number,
  tracer: LetterTracer,
): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    tracer.push("probe.tcp.start", { host, port, timeoutMs });
    let settled = false;
    const socket = net.connect({ host, port });
    attachSocketTrace(socket, tracer, "probe.tcp");
    const timer = setTimeout(() => {
      socket.destroy(new Error(`TCP probe timed out after ${timeoutMs}ms`));
    }, timeoutMs);
    socket.on("connect", () => {
      if (settled) {
        return;
      }
      settled = true;
      const info = {
        remoteAddress: socket.remoteAddress,
        remotePort: socket.remotePort,
        localAddress: socket.localAddress,
        localPort: socket.localPort,
      };
      clearTimeout(timer);
      socket.end();
      tracer.push("probe.tcp.ok", info);
      resolve(info);
    });
    socket.on("error", (error) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      tracer.push("probe.tcp.fail", { message: error.message });
      reject(error);
    });
  });
}

export function probeTls(
  host: string,
  port: number,
  timeoutMs: number,
  rejectUnauthorized: boolean,
  tracer: LetterTracer,
): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    const label = rejectUnauthorized ? "probe.tls.verify" : "probe.tls.noverify";
    tracer.push(`${label}.start`, { host, port, timeoutMs, rejectUnauthorized });
    let settled = false;
    const socket = tls.connect({
      host,
      port,
      servername: host,
      rejectUnauthorized,
      minVersion: "TLSv1.2",
    });
    attachSocketTrace(socket, tracer, label);
    const timer = setTimeout(() => {
      socket.destroy(new Error(`TLS probe timed out after ${timeoutMs}ms`));
    }, timeoutMs);
    const finish = (error?: Error) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      const info = summarizeTlsSocket(socket);
      try {
        socket.end();
      } catch {
        socket.destroy();
      }
      if (error) {
        tracer.push(`${label}.fail`, {
          message: error.message,
          ...info,
        });
        reject(error);
        return;
      }
      tracer.push(`${label}.ok`, info);
      resolve(info);
    };
    socket.on("secureConnect", () => finish());
    socket.on("error", (error) => finish(error));
    socket.on("timeout", () => finish(new Error(`TLS probe timed out after ${timeoutMs}ms`)));
  });
}
