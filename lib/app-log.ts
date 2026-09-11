import { appendFileSync, existsSync, mkdirSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

export type LogLevel = "info" | "warn" | "error";

export type AppLogEntry = {
  id: string;
  timestamp: string;
  level: LogLevel;
  source: string;
  message: string;
  captureId?: string;
  details?: Record<string, unknown>;
};

export function logFilePath(): string {
  return resolve(process.env.APP_LOG_PATH || "data/app-logs.txt");
}

export function newLogId(): string {
  return `log_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

export function formatLogText(entry: AppLogEntry): string {
  return [
    `### LOG ${entry.id} ${entry.timestamp} ${entry.level} ${entry.source}`,
    entry.message,
    `### JSON ${JSON.stringify(entry)}`,
    "### END",
    "",
  ].join("\n");
}

export function parseLogFile(contents: string): AppLogEntry[] {
  const entries: AppLogEntry[] = [];
  const chunks = contents.split("### JSON ");
  for (let i = 1; i < chunks.length; i += 1) {
    const jsonLine = chunks[i].split("\n")[0];
    try {
      entries.push(JSON.parse(jsonLine) as AppLogEntry);
    } catch {
      // skip a malformed block
    }
  }
  return entries;
}

export function readStoredLogs(): AppLogEntry[] {
  const file = logFilePath();
  if (!existsSync(file)) {
    return [];
  }
  try {
    return parseLogFile(readFileSync(file, "utf8"));
  } catch {
    return [];
  }
}

export function readLogFileText(): string {
  const file = logFilePath();
  if (!existsSync(file)) {
    return "";
  }
  try {
    return readFileSync(file, "utf8");
  } catch {
    return "";
  }
}

export function appendLog(input: {
  level?: LogLevel;
  source: string;
  message: string;
  captureId?: string;
  details?: Record<string, unknown>;
}): AppLogEntry | null {
  const entry: AppLogEntry = {
    id: newLogId(),
    timestamp: new Date().toISOString(),
    level: input.level || "info",
    source: input.source,
    message: input.message,
    captureId: input.captureId,
    details: input.details,
  };

  try {
    const file = logFilePath();
    mkdirSync(dirname(file), { recursive: true });
    appendFileSync(file, formatLogText(entry), "utf8");
    return entry;
  } catch {
    return null;
  }
}
