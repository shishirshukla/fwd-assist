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
import type { StoredForwardCapture } from "@/lib/forward-capture";

type CapturesResponse = {
  count: number;
  captures: StoredForwardCapture[];
  pushUrlConfigured: boolean;
};

export function CapturesList() {
  const [data, setData] = useState<CapturesResponse | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);

  async function load() {
    setLoading(true);
    setError("");
    try {
      const response = await fetch("/api/captures");
      if (!response.ok) {
        throw new Error("Could not load captures.");
      }
      setData((await response.json()) as CapturesResponse);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not load captures.");
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
          <h1 className="text-lg font-semibold">Forward captures</h1>
          <p className="mt-1 text-sm text-muted-foreground">
            Each classified forward is stored in{" "}
            <code className="text-xs">data/forward-captures.txt</code> and
            submitted to the pensioner letter API. Edit{" "}
            <code className="text-xs">data/letter-lookup.json</code> to map TO
            addresses to department and FROM addresses to EntryBy.
          </p>
        </div>
        <div className="flex gap-2">
          <Button variant="outline" size="sm" onClick={() => void load()}>
            Refresh
          </Button>
          <Link
            href="/"
            className="inline-flex h-7 items-center rounded-lg border border-border px-2.5 text-[0.8rem] font-medium"
          >
            Simulator
          </Link>
        </div>
      </div>

      {data ? (
        <p className="text-sm text-muted-foreground">
          {data.count} stored · letter API{" "}
          {data.pushUrlConfigured ? "enabled" : "disabled"}
        </p>
      ) : null}

      {loading ? (
        <p className="text-sm text-muted-foreground">Loading captures…</p>
      ) : null}

      {error ? <p className="text-sm text-destructive">{error}</p> : null}

      {!loading && data && data.captures.length === 0 ? (
        <Card>
          <CardHeader>
            <CardTitle>No captures yet</CardTitle>
            <CardDescription>
              Classify a forwarded message in the simulator or Outlook. Save and
              send writes Original Email Date, Sender, Subject, Body, and TO/CC
              here.
            </CardDescription>
          </CardHeader>
        </Card>
      ) : null}

      {data?.captures.map((capture) => (
        <Card key={capture.id}>
          <CardHeader>
            <CardTitle className="break-words">{capture.subject || "(no subject)"}</CardTitle>
            <CardDescription>{capture.capturedAt}</CardDescription>
          </CardHeader>
          <CardContent className="space-y-3">
            <dl className="grid gap-2 text-sm sm:grid-cols-2">
              <div>
                <dt className="text-muted-foreground">Original Email Date</dt>
                <dd>{capture.originalEmailDate || "—"}</dd>
              </div>
              <div>
                <dt className="text-muted-foreground">Sender Name</dt>
                <dd>{capture.senderName || "—"}</dd>
              </div>
              <div className="sm:col-span-2">
                <dt className="text-muted-foreground">Sender Email ID</dt>
                <dd className="break-all">{capture.senderEmailId || "—"}</dd>
              </div>
              <div>
                <dt className="text-muted-foreground">TO</dt>
                <dd className="break-all">
                  {(capture.toEmailAddresses || []).join(", ") || "—"}
                </dd>
              </div>
              <div>
                <dt className="text-muted-foreground">Forwarded by</dt>
                <dd className="break-all">{capture.forwardedByEmail || "—"}</dd>
              </div>
              <div>
                <dt className="text-muted-foreground">CC</dt>
                <dd className="break-all">
                  {(capture.ccEmailAddresses || []).join(", ") || "—"}
                </dd>
              </div>
            </dl>
            {capture.classification ? (
              <div className="flex flex-wrap gap-2">
                <Badge variant="secondary">
                  Priority {capture.classification.priority}
                </Badge>
                <Badge variant="secondary">
                  End {capture.classification.endDate}
                </Badge>
                <Badge variant="secondary">
                  {capture.classification.category}
                </Badge>
              </div>
            ) : null}
            {capture.letterSubmit ? (
              <dl className="grid gap-2 rounded-lg bg-muted/50 p-3 text-sm sm:grid-cols-2">
                <div className="sm:col-span-2 font-medium">Letter API fields</div>
                <div>
                  <dt className="text-muted-foreground">receivingDate / letterDate</dt>
                  <dd>{capture.letterSubmit.receivingDate}</dd>
                </div>
                <div>
                  <dt className="text-muted-foreground">priority</dt>
                  <dd>{capture.letterSubmit.priority}</dd>
                </div>
                <div>
                  <dt className="text-muted-foreground">department</dt>
                  <dd>{capture.letterSubmit.department}</dd>
                </div>
                <div>
                  <dt className="text-muted-foreground">EntryBy</dt>
                  <dd>{capture.letterSubmit.EntryBy}</dd>
                </div>
                <div>
                  <dt className="text-muted-foreground">senderOffice</dt>
                  <dd className="break-all">{capture.letterSubmit.senderOffice}</dd>
                </div>
                <div>
                  <dt className="text-muted-foreground">sendName</dt>
                  <dd>{capture.letterSubmit.sendName}</dd>
                </div>
              </dl>
            ) : null}
            {capture.remotePush?.attempted ? (
              <p className="text-xs text-muted-foreground break-all">
                Letter API {capture.remotePush.ok ? "accepted" : "failed"}
                {capture.remotePush.status ? ` (${capture.remotePush.status})` : ""}
                {capture.remotePush.error ? ` — ${capture.remotePush.error}` : ""}
              </p>
            ) : null}
            <pre className="max-h-48 overflow-auto rounded-lg bg-muted p-3 text-xs whitespace-pre-wrap">
              {capture.messageBody || "(empty body)"}
            </pre>
          </CardContent>
        </Card>
      ))}
    </div>
  );
}
