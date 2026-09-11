"use client";

import { useEffect, useState } from "react";
import Link from "next/link";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import type { AppLogEntry } from "@/lib/app-log";

type LogsResponse = {
  count: number;
  logs: AppLogEntry[];
};

export function LogsList() {
  const [data, setData] = useState<LogsResponse | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);

  async function load() {
    setLoading(true);
    setError("");
    try {
      const response = await fetch("/api/logs?limit=200");
      if (!response.ok) {
        throw new Error("Could not load logs.");
      }
      setData((await response.json()) as LogsResponse);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not load logs.");
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    void load();
  }, []);

  return (
    <div className="mx-auto flex w-full max-w-3xl flex-col gap-4 p-4 pb-10 sm:p-6">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <h1 className="text-lg font-semibold">Application logs</h1>
          <p className="mt-1 text-sm text-muted-foreground">
            Stored in <code className="text-xs">data/app-logs.txt</code>. JSON:{" "}
            <code className="text-xs">GET /api/logs</code>. Plain text:{" "}
            <code className="text-xs">GET /api/logs?format=text</code>.
          </p>
        </div>
        <div className="flex gap-2">
          <Button variant="outline" size="sm" onClick={() => void load()}>
            Refresh
          </Button>
          <Link
            href="/captures"
            className="inline-flex h-7 items-center rounded-lg border border-border px-2.5 text-[0.8rem] font-medium"
          >
            Captures
          </Link>
        </div>
      </div>

      {data ? (
        <p className="text-sm text-muted-foreground">{data.count} entries</p>
      ) : null}
      {loading ? (
        <p className="text-sm text-muted-foreground">Loading logs…</p>
      ) : null}
      {error ? <p className="text-sm text-destructive">{error}</p> : null}

      {!loading && data && data.logs.length === 0 ? (
        <Card>
          <CardHeader>
            <CardTitle>No logs yet</CardTitle>
            <CardDescription>
              Classify a forward. Capture and letter-API steps are written here.
            </CardDescription>
          </CardHeader>
        </Card>
      ) : null}

      {data?.logs.map((entry) => (
        <Card key={entry.id}>
          <CardHeader>
            <CardTitle className="flex flex-wrap items-center gap-2 text-sm">
              <Badge
                variant={entry.level === "error" ? "destructive" : "secondary"}
              >
                {entry.level}
              </Badge>
              <span>{entry.source}</span>
            </CardTitle>
            <CardDescription>{entry.timestamp}</CardDescription>
          </CardHeader>
          <CardContent className="space-y-2 text-sm">
            <p>{entry.message}</p>
            {entry.captureId ? (
              <p className="text-muted-foreground">Capture {entry.captureId}</p>
            ) : null}
            {entry.details ? (
              <pre className="max-h-48 overflow-auto rounded-lg bg-muted p-3 text-xs whitespace-pre-wrap">
                {JSON.stringify(entry.details, null, 2)}
              </pre>
            ) : null}
          </CardContent>
        </Card>
      ))}
    </div>
  );
}
