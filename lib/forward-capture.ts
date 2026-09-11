import { mkdirSync, readFileSync, appendFileSync, existsSync } from "node:fs";
import { dirname, resolve } from "node:path";

import {
  letterSubmitEnabled,
  mapLetterSubmitFields,
  submitLetter,
  type LetterSubmitFields,
} from "@/lib/letter-submit";

export type ForwardCaptureInput = {
  originalEmailDate?: string;
  senderEmailId?: string;
  senderName?: string;
  subject?: string;
  messageBody?: string;
  toEmailAddresses?: string[] | string;
  ccEmailAddresses?: string[] | string;
  originalToEmailAddresses?: string[] | string;
  forwardedByEmail?: string;
  classification?: {
    priority?: string;
    endDate?: string;
    category?: string;
  } | null;
};

export type StoredForwardCapture = {
  id: string;
  capturedAt: string;
  originalEmailDate: string;
  senderEmailId: string;
  senderName: string;
  subject: string;
  messageBody: string;
  toEmailAddresses: string[];
  ccEmailAddresses: string[];
  originalToEmailAddresses: string[];
  forwardedByEmail: string;
  classification: {
    priority: string;
    endDate: string;
    category: string;
  } | null;
  letterSubmit: LetterSubmitFields;
  remotePush: {
    attempted: boolean;
    ok: boolean;
    status: number | null;
    error: string | null;
    url?: string;
    responseText?: string;
  };
};

export function captureFilePath(): string {
  return resolve(process.env.CAPTURE_FILE_PATH || "data/forward-captures.txt");
}

export function parseMailbox(line: string): { name: string; email: string } {
  const trimmed = (line || "").trim();
  const angled = trimmed.match(/^(.*?)\s*<([^>]+)>/);
  if (angled) {
    return {
      name: angled[1].replace(/^["']|["']$/g, "").trim(),
      email: angled[2].trim(),
    };
  }
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed)) {
    return { name: "", email: trimmed };
  }
  return { name: trimmed, email: "" };
}

export function parseOriginalForwardHeaders(body: string): {
  originalEmailDate: string;
  senderName: string;
  senderEmailId: string;
  originalToEmailAddresses: string[];
} {
  const text = (body || "").replace(/\r\n/g, "\n");
  const fromLine = matchHeader(text, "From");
  const sentLine = matchHeader(text, "Sent") || matchHeader(text, "Date");
  const from = parseMailbox(fromLine);
  return {
    originalEmailDate: sentLine,
    senderName: from.name,
    senderEmailId: from.email,
    originalToEmailAddresses: normalizeEmailList(matchHeader(text, "To")),
  };
}

function matchHeader(text: string, name: string): string {
  const re = new RegExp(`^${name}:\\s*(.+)$`, "im");
  const match = text.match(re);
  return match ? match[1].trim() : "";
}

export function normalizeEmailList(value: string[] | string | undefined): string[] {
  const raw = Array.isArray(value) ? value.join(",") : value || "";
  return raw
    .split(/[;,]+/)
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const parsed = parseMailbox(part);
      return parsed.email || part;
    });
}

export function normalizeCapture(
  input: ForwardCaptureInput,
): Omit<StoredForwardCapture, "id" | "capturedAt" | "remotePush"> {
  const parsed = parseOriginalForwardHeaders(input.messageBody || "");
  const classification = input.classification
    ? {
        priority: String(input.classification.priority || ""),
        endDate: String(input.classification.endDate || ""),
        category: String(input.classification.category || ""),
      }
    : null;

  const base = {
    originalEmailDate: String(input.originalEmailDate || parsed.originalEmailDate || ""),
    senderEmailId: String(input.senderEmailId || parsed.senderEmailId || ""),
    senderName: String(input.senderName || parsed.senderName || ""),
    subject: String(input.subject || ""),
    messageBody: String(input.messageBody || ""),
    toEmailAddresses: normalizeEmailList(input.toEmailAddresses),
    ccEmailAddresses: normalizeEmailList(input.ccEmailAddresses),
    originalToEmailAddresses:
      normalizeEmailList(input.originalToEmailAddresses).length > 0
        ? normalizeEmailList(input.originalToEmailAddresses)
        : parsed.originalToEmailAddresses,
    forwardedByEmail: String(input.forwardedByEmail || ""),
    classification:
      classification &&
      (classification.priority || classification.endDate || classification.category)
        ? classification
        : null,
  };

  return {
    ...base,
    letterSubmit: mapLetterSubmitFields({
      originalEmailDate: base.originalEmailDate,
      senderEmailId: base.senderEmailId,
      senderName: base.senderName,
      subject: base.subject,
      priority: base.classification?.priority,
      toEmailAddresses: base.toEmailAddresses,
      originalToEmailAddresses: base.originalToEmailAddresses,
      forwardedByEmail: base.forwardedByEmail,
    }),
  };
}

export function formatCaptureText(record: StoredForwardCapture): string {
  return [
    `### CAPTURE ${record.id} ${record.capturedAt}`,
    `Original Email Date: ${record.originalEmailDate}`,
    `Sender Email ID: ${record.senderEmailId}`,
    `Sender Name: ${record.senderName}`,
    `Subject: ${record.subject}`,
    `TO Email addresses: ${record.toEmailAddresses.join(", ")}`,
    `CC Email addresses: ${record.ccEmailAddresses.join(", ")}`,
    `Forwarded by: ${record.forwardedByEmail || ""}`,
    `Letter department: ${record.letterSubmit?.department || ""}`,
    `Letter EntryBy: ${record.letterSubmit?.EntryBy || ""}`,
    `Letter dates: ${record.letterSubmit?.receivingDate || ""}`,
    "Message Body:",
    record.messageBody,
    `### JSON ${JSON.stringify(record)}`,
    "### END",
    "",
  ].join("\n");
}

export function parseCaptureFile(contents: string): StoredForwardCapture[] {
  const records: StoredForwardCapture[] = [];
  const chunks = contents.split("### JSON ");
  for (let i = 1; i < chunks.length; i += 1) {
    const jsonLine = chunks[i].split("\n")[0];
    try {
      records.push(JSON.parse(jsonLine) as StoredForwardCapture);
    } catch {
      // skip a malformed block
    }
  }
  return records;
}

export function readStoredCaptures(): StoredForwardCapture[] {
  const file = captureFilePath();
  if (!existsSync(file)) {
    return [];
  }
  return parseCaptureFile(readFileSync(file, "utf8"));
}

export function appendCapture(record: StoredForwardCapture): void {
  const file = captureFilePath();
  mkdirSync(dirname(file), { recursive: true });
  appendFileSync(file, formatCaptureText(record), "utf8");
}

export async function pushCaptureIfConfigured(
  record: StoredForwardCapture,
): Promise<StoredForwardCapture["remotePush"]> {
  if (!letterSubmitEnabled()) {
    return { attempted: false, ok: false, status: null, error: null };
  }

  const result = await submitLetter(record.letterSubmit);
  return {
    attempted: result.attempted,
    ok: result.ok,
    status: result.status,
    error: result.error,
    url: result.url,
    responseText: result.responseText,
  };
}

export function newCaptureId(): string {
  return `cap_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}
