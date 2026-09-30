import { mkdirSync, readFileSync, appendFileSync, existsSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

import {
  letterSubmitBaseUrl,
  letterSubmitEnabled,
  letterSubmitFromServer,
  mapLetterSubmitFields,
  mapLetterSubmitsForToAddresses,
  subjectHasLmsKeyword,
  submitLetter,
  buildLetterSubmitUrl,
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
  letterSubmits?: { toEmail: string; fields: LetterSubmitFields }[];
  letterSubmitShouldRun?: boolean;
  remotePush: {
    attempted: boolean;
    ok: boolean;
    status: number | null;
    error: string | null;
    letterId?: string;
    headerSet?: boolean;
    url?: string;
    responseText?: string;
    source?: string;
  };
};

export function captureFilePath(): string {
  return process.env.CAPTURE_FILE_PATH || join(process.cwd(), "data", "forward-captures.txt");
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

  const letterSubmits = mapLetterSubmitsForToAddresses({
    originalEmailDate: base.originalEmailDate,
    senderEmailId: base.senderEmailId,
    senderName: base.senderName,
    subject: base.subject,
    priority: base.classification?.priority,
    toEmailAddresses: base.toEmailAddresses,
    forwardedByEmail: base.forwardedByEmail,
  });

  return {
    ...base,
    letterSubmit: letterSubmits[0]?.fields ||
      mapLetterSubmitFields({
        originalEmailDate: base.originalEmailDate,
        senderEmailId: base.senderEmailId,
        senderName: base.senderName,
        subject: base.subject,
        priority: base.classification?.priority,
        toEmailAddresses: base.toEmailAddresses,
        originalToEmailAddresses: base.originalToEmailAddresses,
        forwardedByEmail: base.forwardedByEmail,
      }),
    letterSubmits,
    letterSubmitShouldRun: subjectHasLmsKeyword(base.subject),
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
    `Letter LMS: ${record.letterSubmitShouldRun ? "yes" : "no"}`,
    `Letter To submits: ${(record.letterSubmits || [])
      .map((row) => `${row.toEmail || "(none)"}->${row.fields.department}`)
      .join("; ")}`,
    `Letter ID: ${record.remotePush?.letterId || ""}`,
    `X-LETTERID-CGB: ${record.remotePush?.headerSet ? "set" : ""}`,
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
  if (!existsSync(/* turbopackIgnore: true */ file)) {
    return [];
  }
  return parseCaptureFile(readFileSync(/* turbopackIgnore: true */ file, "utf8"));
}

export function appendCapture(record: StoredForwardCapture): void {
  const file = captureFilePath();
  mkdirSync(/* turbopackIgnore: true */ dirname(file), { recursive: true });
  appendFileSync(/* turbopackIgnore: true */ file, formatCaptureText(record), "utf8");
}

export async function pushCaptureIfConfigured(
  record: StoredForwardCapture,
): Promise<StoredForwardCapture["remotePush"]> {
  if (!letterSubmitEnabled()) {
    return { attempted: false, ok: false, status: null, error: null };
  }

  if (!subjectHasLmsKeyword(record.subject)) {
    return {
      attempted: false,
      ok: false,
      status: null,
      error: null,
      source: letterSubmitFromServer() ? "server" : "browser",
      url: buildLetterSubmitUrl(letterSubmitBaseUrl()),
      responseText: "skipped: subject does not contain [LMS]",
    };
  }

  if (!letterSubmitFromServer()) {
    return {
      attempted: false,
      ok: false,
      status: null,
      error: null,
      url: buildLetterSubmitUrl(letterSubmitBaseUrl()),
      source: "browser",
      responseText:
        "skipped-server: WAF blocks the hosting provider. Outlook/browser POSTs submit-letter.",
    };
  }

  const payloads = (record.letterSubmits || []).map((row) => row.fields);
  const toSubmit = payloads.length > 0 ? payloads : [record.letterSubmit];
  const results = [];
  for (const fields of toSubmit) {
    results.push(await submitLetter(fields));
  }
  const letterIds = results
    .map((result) => {
      const text = result.responseText || "";
      const match = text.match(/"letterId"\s*:\s*"?([^",}\s]+)/i);
      return match ? match[1] : "";
    })
    .filter(Boolean);
  const failed = results.find((result) => !result.ok);
  const last = results[results.length - 1];
  return {
    attempted: true,
    ok: !failed,
    status: last?.status ?? null,
    error: failed?.error || last?.error || null,
    url: last?.url,
    source: "server",
    letterId: letterIds.join(","),
    responseText: results.map((result) => result.responseText || "").join("\n---\n"),
  };
}

export function updateCaptureRemotePush(
  id: string,
  remotePush: StoredForwardCapture["remotePush"],
): StoredForwardCapture | null {
  const records = readStoredCaptures();
  const index = records.findIndex((record) => record.id === id);
  if (index < 0) {
    return null;
  }
  records[index] = { ...records[index], remotePush };
  const file = captureFilePath();
  mkdirSync(/* turbopackIgnore: true */ dirname(file), { recursive: true });
  writeFileSync(
    /* turbopackIgnore: true */ file,
    records.map((record) => formatCaptureText(record)).join(""),
    "utf8",
  );
  return records[index];
}

export function newCaptureId(): string {
  return `cap_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}
