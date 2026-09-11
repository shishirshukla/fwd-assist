import dns from "node:dns/promises";
import net from "node:net";
import tls from "node:tls";

import { letterSubmitBaseUrl } from "@/lib/letter-submit";

export const dynamic = "force-dynamic";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

export function OPTIONS() {
  return new Response(null, { status: 204, headers: cors });
}

function tcpConnect(host: string, port: number, timeoutMs: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const socket = net.connect({ host, port, family: 4 });
    const timer = setTimeout(() => {
      socket.destroy();
      reject(new Error(`TCP timeout to ${host}:${port} after ${timeoutMs}ms`));
    }, timeoutMs);
    socket.on("connect", () => {
      const address = socket.remoteAddress || "";
      clearTimeout(timer);
      socket.end();
      resolve(address);
    });
    socket.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
  });
}

function tlsConnect(host: string, port: number, timeoutMs: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const socket = tls.connect({
      host,
      port,
      servername: host,
      timeout: timeoutMs,
    });
    socket.on("secureConnect", () => {
      const cert = socket.getPeerCertificate();
      const subject =
        cert && "subject" in cert ? JSON.stringify(cert.subject) : "unknown cert";
      socket.end();
      resolve(subject);
    });
    socket.on("error", reject);
    socket.on("timeout", () => {
      socket.destroy();
      reject(new Error(`TLS timeout to ${host}:${port}`));
    });
  });
}

export async function GET() {
  const target = new URL(letterSubmitBaseUrl());
  const host = target.hostname;
  const port = Number(target.port || 443);
  const report: Record<string, unknown> = {
    url: target.toString(),
    host,
    port,
  };

  try {
    const addresses = await dns.lookup(host, { all: true, family: 4 });
    report.dns = addresses;
  } catch (error) {
    report.dnsError = error instanceof Error ? error.message : String(error);
    return Response.json(report, { status: 503, headers: cors });
  }

  try {
    report.tcpAddress = await tcpConnect(host, port, 8000);
  } catch (error) {
    report.tcpError = error instanceof Error ? error.message : String(error);
    report.reachable = false;
    report.hint =
      "This app host cannot open TCP to the letter API. Railway and other public clouds are often blocked. Run Forward Guard on a machine that can open https://eloan.cgbankmobile.in in a browser (office/WSL), or ask the bank to allow this server's outbound IP to 43.225.0.54:443.";
    return Response.json(report, { status: 503, headers: cors });
  }

  try {
    report.tlsCertificate = await tlsConnect(host, port, 8000);
  } catch (error) {
    report.tlsError = error instanceof Error ? error.message : String(error);
    report.reachable = false;
    report.hint =
      "TCP connected but TLS never completed. Public clouds (including Railway) usually cannot finish a handshake to this bank host. Run npm start on WSL/office network that can load the letter API in a browser, expose that host with ngrok for Outlook, or ask the bank to allow this server IP.";
    return Response.json(report, { status: 503, headers: cors });
  }

  report.reachable = !report.tlsError;
  return Response.json(report, { headers: cors });
}
