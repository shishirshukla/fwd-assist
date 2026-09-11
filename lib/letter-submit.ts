import { existsSync, readFileSync } from "node:fs";
import http from "node:http";
import https from "node:https";
import { resolve } from "node:path";
import type { RequestOptions } from "node:https";

import { appendLog } from "@/lib/app-log";

export const DEFAULT_LETTER_SUBMIT_URL =
  "https://eloan.cgbankmobile.in/pensioner_api/auth/api/submit-letter";

export type LetterLookupFile = {
  departmentByToEmail?: Record<string, string>;
  entryByFromEmail?: Record<string, string>;
  defaults?: {
    department?: string;
    entryBy?: string;
  };
};

export type LetterSubmitFields = {
  receivingDate: string;
  senderOffice: string;
  sendName: string;
  letterNo: string;
  letterDate: string;
  letterDesc: string;
  department: string;
  priority: string;
  EntryBy: string;
};

function na(value: string | undefined | null): string {
  const trimmed = (value || "").trim();
  return trimmed ? trimmed : "NA";
}

export function toYyyyMmDd(value: string | undefined, fallback = new Date()): string {
  const trimmed = (value || "").trim();
  const iso = trimmed.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return `${iso[1]}-${iso[2]}-${iso[3]}`;
  }

  const parsed = Date.parse(trimmed);
  if (!Number.isNaN(parsed)) {
    const date = new Date(parsed);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  const year = fallback.getFullYear();
  const month = String(fallback.getMonth() + 1).padStart(2, "0");
  const day = String(fallback.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function normalizeKey(email: string): string {
  return email.trim().toLowerCase();
}

function lookupValue(
  table: Record<string, string> | undefined,
  emails: string[],
  fallback: string,
): string {
  const map: Record<string, string> = {};
  for (const [key, value] of Object.entries(table || {})) {
    map[normalizeKey(key)] = value;
  }

  for (const email of emails) {
    const full = normalizeKey(email);
    if (!full) {
      continue;
    }
    if (map[full]) {
      return map[full];
    }
    const local = full.split("@")[0];
    if (local && map[local]) {
      return map[local];
    }
  }

  return fallback;
}

export function loadLetterLookup(): LetterLookupFile {
  const file = resolve(process.env.LETTER_LOOKUP_PATH || "data/letter-lookup.json");
  if (!existsSync(file)) {
    return {};
  }
  try {
    return JSON.parse(readFileSync(file, "utf8")) as LetterLookupFile;
  } catch {
    return {};
  }
}

export function letterSubmitBaseUrl(): string {
  return (
    process.env.CAPTURE_PUSH_URL?.trim() ||
    process.env.LETTER_SUBMIT_URL?.trim() ||
    DEFAULT_LETTER_SUBMIT_URL
  );
}

export function letterSubmitEnabled(): boolean {
  return process.env.LETTER_SUBMIT_DISABLED !== "1";
}

/** Railway/cloud Node usually cannot reach the bank host. Default is Outlook-client submit. */
export function letterSubmitFromServer(): boolean {
  return process.env.LETTER_SUBMIT_FROM_SERVER === "1";
}

export function letterSubmitHost(): string {
  try {
    return new URL(letterSubmitBaseUrl()).hostname;
  } catch {
    return "eloan.cgbankmobile.in";
  }
}

export function mapLetterSubmitFields(input: {
  originalEmailDate?: string;
  senderEmailId?: string;
  senderName?: string;
  subject?: string;
  priority?: string;
  toEmailAddresses?: string[];
  originalToEmailAddresses?: string[];
  forwardedByEmail?: string;
}): LetterSubmitFields {
  const lookup = loadLetterLookup();
  const departmentDefault = lookup.defaults?.department || "NA";
  const entryByDefault = lookup.defaults?.entryBy || "NA";
  const emailDate = toYyyyMmDd(input.originalEmailDate);

  return {
    receivingDate: emailDate,
    senderOffice: na(input.senderEmailId),
    sendName: na(input.senderName),
    letterNo: "NA",
    letterDate: emailDate,
    letterDesc: na(input.subject),
    department: lookupValue(
      lookup.departmentByToEmail,
      [...(input.toEmailAddresses || []), ...(input.originalToEmailAddresses || [])],
      departmentDefault,
    ),
    priority: na(input.priority),
    EntryBy: lookupValue(
      lookup.entryByFromEmail,
      [input.forwardedByEmail || "", input.senderEmailId || ""].filter(Boolean),
      entryByDefault,
    ),
  };
}

export function letterSubmitPayload(fields: LetterSubmitFields): LetterSubmitFields {
  return {
    receivingDate: fields.receivingDate,
    senderOffice: fields.senderOffice,
    sendName: fields.sendName,
    letterNo: fields.letterNo,
    letterDate: fields.letterDate,
    letterDesc: fields.letterDesc,
    department: fields.department,
    priority: fields.priority,
    EntryBy: fields.EntryBy,
  };
}

export function buildLetterSubmitUrl(baseUrl: string): string {
  const url = new URL(baseUrl);
  url.search = "";
  return url.toString();
}

function flattenError(error: unknown): { message: string; details: Record<string, unknown> } {
  const chain: Record<string, unknown>[] = [];
  let current: unknown = error;
  for (let i = 0; i < 6 && current; i += 1) {
    if (current instanceof Error) {
      const nodeErr = current as NodeJS.ErrnoException & { cause?: unknown };
      chain.push({
        name: current.name,
        message: current.message,
        code: nodeErr.code,
        errno: nodeErr.errno,
        syscall: nodeErr.syscall,
        address: (nodeErr as NodeJS.ErrnoException & { address?: string }).address,
        port: (nodeErr as NodeJS.ErrnoException & { port?: number }).port,
      });
      current = nodeErr.cause;
    } else {
      chain.push({ value: String(current) });
      break;
    }
  }
  const message = chain
    .map((part) =>
      [part.code, part.message || part.value].filter(Boolean).join(": "),
    )
    .filter(Boolean)
    .join(" <- ");
  return { message: message || "unknown error", details: { chain } };
}

function isTlsTrustError(message: string): boolean {
  return /UNABLE_TO_VERIFY|CERT_|self[- ]signed|unable to verify the first certificate|ERR_TLS/i.test(
    message,
  );
}

function postJson(
  urlString: string,
  body: string,
  insecureTls: boolean,
): Promise<{ status: number; text: string }> {
  return new Promise((resolvePromise, reject) => {
    const url = new URL(urlString);
    const client = url.protocol === "http:" ? http : https;
    const timeoutMs = Number(process.env.LETTER_SUBMIT_TIMEOUT_MS || 30000);
    const options: RequestOptions = {
      protocol: url.protocol,
      hostname: url.hostname,
      port: url.port || (url.protocol === "https:" ? 443 : 80),
      path: `${url.pathname}${url.search}`,
      method: "POST",
      family: 4,
      timeout: timeoutMs,
      rejectUnauthorized: !insecureTls,
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json, text/plain, */*",
        "User-Agent": "ForwardGuard/1.0",
        "Content-Length": Buffer.byteLength(body),
      },
    };

    const req = client.request(options, (res) => {
      const chunks: Buffer[] = [];
      res.on("data", (chunk) => {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      });
      res.on("end", () => {
        resolvePromise({
          status: res.statusCode || 0,
          text: Buffer.concat(chunks).toString("utf8").slice(0, 20000),
        });
      });
    });

    req.on("socket", (socket) => {
      socket.setTimeout(timeoutMs);
    });
    req.on("timeout", () => {
      req.destroy(
        new Error(
          `ETIMEDOUT: connection timed out after ${timeoutMs}ms to ${url.hostname}:${options.port}. This Node host cannot reach the letter API. Run Forward Guard on the bank network (WSL + ngrok), or whitelist this server's outbound IP on the bank firewall.`,
        ),
      );
    });
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}

export async function submitLetter(
  fields: LetterSubmitFields,
): Promise<{
  attempted: boolean;
  ok: boolean;
  status: number | null;
  error: string | null;
  url: string;
  responseText: string;
}> {
  const url = buildLetterSubmitUrl(letterSubmitBaseUrl());
  const method = "POST";
  const payload = letterSubmitPayload(fields);
  const body = JSON.stringify(payload);

  appendLog({
    level: "info",
    source: "letter-submit",
    message: `Calling letter API (POST JSON) ${url}`,
    details: { method, url, payload, body },
  });

  async function attempt(insecureTls: boolean) {
    return postJson(url, body, insecureTls);
  }

  try {
    let insecureTls = process.env.LETTER_SUBMIT_TLS_INSECURE === "1";
    let response: { status: number; text: string };
    try {
      response = await attempt(insecureTls);
    } catch (firstError) {
      const first = flattenError(firstError);
      if (!insecureTls && isTlsTrustError(first.message)) {
        appendLog({
          level: "warn",
          source: "letter-submit",
          message: `TLS verification failed, retrying without certificate check: ${first.message}`,
          details: { url, error: first.message, errorDetails: first.details },
        });
        insecureTls = true;
        response = await attempt(true);
      } else {
        throw firstError;
      }
    }

    const responseText = response.text;
    const result = {
      attempted: true,
      ok: response.status >= 200 && response.status < 300,
      status: response.status,
      error:
        response.status >= 200 && response.status < 300
          ? null
          : `Letter API returned ${response.status}`,
      url,
      responseText,
    };
    appendLog({
      level: result.ok ? "info" : "error",
      source: "letter-submit",
      message: result.ok
        ? `Letter API returned ${response.status}: ${responseText || "(empty body)"}`
        : `Letter API error ${response.status}: ${responseText || result.error}`,
      details: {
        url,
        method,
        status: response.status,
        ok: result.ok,
        insecureTls,
        responseBody: responseText || "(empty body)",
        error: result.error,
      },
    });
    return result;
  } catch (error) {
    const serialized = flattenError(error);
    let message = `Letter API request failed: ${serialized.message}`;
    if (/ETIMEDOUT|timed out|ECONNRESET|ENETUNREACH|EHOSTUNREACH|fetch failed/i.test(message)) {
      message += ` Check GET /api/letter-health on this same host.`;
    }
    appendLog({
      level: "error",
      source: "letter-submit",
      message,
      details: {
        url,
        method,
        error: message,
        errorDetails: serialized.details,
        responseBody: "",
      },
    });
    return {
      attempted: true,
      ok: false,
      status: null,
      error: message,
      url,
      responseText: "",
    };
  }
}
