import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";

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

function serializeError(error: unknown): Record<string, unknown> {
  if (error instanceof Error) {
    const cause =
      error.cause instanceof Error
        ? { name: error.cause.name, message: error.cause.message }
        : error.cause
          ? { value: String(error.cause) }
          : undefined;
    return {
      name: error.name,
      message: error.message,
      stack: error.stack || "",
      cause,
    };
  }
  return { message: String(error) };
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

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json, text/plain, */*",
        "Content-Type": "application/json",
      },
      body,
    });
    const responseText = (await response.text()).slice(0, 20000);
    const result = {
      attempted: true,
      ok: response.ok,
      status: response.status,
      error: response.ok ? null : `Letter API returned ${response.status}`,
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
        ok: response.ok,
        responseBody: responseText || "(empty body)",
        error: result.error,
      },
    });
    return result;
  } catch (error) {
    const serialized = serializeError(error);
    const message = `Letter API request failed: ${String(serialized.message || error)}`;
    appendLog({
      level: "error",
      source: "letter-submit",
      message,
      details: {
        url,
        method,
        error: message,
        errorDetails: serialized,
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
