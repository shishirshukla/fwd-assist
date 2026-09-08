"use client";

import { useState } from "react";
import { AlertCircle, CheckCircle2, Forward, Send } from "lucide-react";

import { ForwardMetadataForm } from "@/components/forward-metadata-form";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  formatClassificationLine,
  isForwardedSubject,
  type ForwardMetadata,
} from "@/lib/forward-metadata";

const ORIGINAL_MESSAGE = `---------- Forwarded message ---------
From: Alex Chen <alex.chen@contoso.com>
Sent: Monday, September 7, 2026 4:12 PM
To: you@contoso.com
Subject: Q3 vendor contract

Please review the attached vendor terms before Friday.`;

type ComposeMode = "new" | "forward";
type SendStatus = "idle" | "blocked" | "sent";

export function OutlookSimulator() {
  const [mode, setMode] = useState<ComposeMode>("forward");
  const [to, setTo] = useState("legal@contoso.com");
  const [subject, setSubject] = useState("FW: Q3 vendor contract");
  const [body, setBody] = useState(
    `Hi Legal,\n\nPassing this along for review.\n\n${ORIGINAL_MESSAGE}`,
  );
  const [formOpen, setFormOpen] = useState(false);
  const [sendStatus, setSendStatus] = useState<SendStatus>("idle");
  const [classification, setClassification] = useState<ForwardMetadata | null>(
    null,
  );
  const [sendError, setSendError] = useState<string | null>(null);

  const forwarded = mode === "forward" || isForwardedSubject(subject);

  function loadNewMessage() {
    setMode("new");
    setTo("");
    setSubject("");
    setBody("");
    setFormOpen(false);
    setSendStatus("idle");
    setClassification(null);
    setSendError(null);
  }

  function loadForward() {
    setMode("forward");
    setTo("legal@contoso.com");
    setSubject("FW: Q3 vendor contract");
    setBody(`Hi Legal,\n\nPassing this along for review.\n\n${ORIGINAL_MESSAGE}`);
    setFormOpen(false);
    setSendStatus("idle");
    setClassification(null);
    setSendError(null);
  }

  function handleSend() {
    setSendError(null);

    if (!to.trim()) {
      setSendError("Add at least one recipient before sending.");
      setSendStatus("idle");
      return;
    }

    if (forwarded && !classification) {
      setSendStatus("blocked");
      setFormOpen(true);
      return;
    }

    setSendStatus("sent");
    setFormOpen(false);
  }

  function handleClassify(data: ForwardMetadata) {
    const line = formatClassificationLine(data);
    setClassification(data);
    setBody((current) =>
      current.includes("[Forward classification:")
        ? current
        : `${line}\n\n${current}`,
    );
    setFormOpen(false);
    setSendStatus("sent");
    setSendError(null);
  }

  return (
    <div className="flex min-h-screen flex-col bg-[#f3f2f1] text-[#201f1e]">
      <header className="flex h-12 items-center gap-3 bg-[#0f6cbd] px-3 text-white sm:px-4">
        <Forward className="size-5 shrink-0" />
        <span className="text-sm font-semibold tracking-wide">Outlook</span>
        <span className="hidden text-xs text-white/80 sm:inline">
          Forward Guard add-in preview
        </span>
      </header>

      <div className="flex min-h-0 flex-1">
        <nav className="hidden w-52 shrink-0 border-r border-[#edebe9] bg-white p-3 md:block">
          <p className="px-2 text-xs font-semibold uppercase tracking-wide text-[#605e5c]">
            Folders
          </p>
          <ul className="mt-2 space-y-1 text-sm">
            {["Inbox", "Sent Items", "Drafts"].map((folder) => (
              <li
                key={folder}
                className="rounded px-2 py-1.5 text-[#201f1e] hover:bg-[#f3f2f1]"
              >
                {folder}
              </li>
            ))}
          </ul>
        </nav>

        <main className="flex min-w-0 flex-1 flex-col p-3 sm:p-5">
          <div className="mb-3 flex flex-wrap items-center gap-2">
            <Button
              type="button"
              variant={mode === "forward" ? "default" : "outline"}
              size="sm"
              onClick={loadForward}
            >
              Load forwarded email
            </Button>
            <Button
              type="button"
              variant={mode === "new" ? "default" : "outline"}
              size="sm"
              onClick={loadNewMessage}
            >
              Load new email
            </Button>
            {forwarded ? (
              <Badge variant="secondary">Forward detected</Badge>
            ) : (
              <Badge variant="outline">New message</Badge>
            )}
          </div>

          <div className="flex min-h-0 flex-1 flex-col gap-3 lg:flex-row">
            <section className="flex min-h-0 min-w-0 flex-1 flex-col overflow-hidden rounded-lg border border-[#edebe9] bg-white shadow-sm">
              <div className="flex flex-wrap items-center gap-2 border-b border-[#edebe9] bg-[#faf9f8] px-3 py-2">
                <button
                  type="button"
                  data-testid="send-email"
                  className="inline-flex h-8 items-center gap-1.5 rounded-lg bg-[#0f6cbd] px-2.5 text-sm font-medium text-white hover:bg-[#115ea3]"
                  onClick={handleSend}
                >
                  <Send className="size-4" />
                  Send
                </button>
                <span className="text-xs text-[#605e5c]">
                  On Send, the add-in checks whether this is a forward.
                </span>
              </div>

              <label className="flex items-center gap-3 border-b border-[#edebe9] px-4 py-2 text-sm">
                <span className="w-12 shrink-0 text-[#605e5c]">To</span>
                <input
                  value={to}
                  onChange={(event) => setTo(event.target.value)}
                  className="min-w-0 flex-1 bg-transparent outline-none"
                  placeholder="Add recipients"
                  autoComplete="off"
                />
              </label>
              <label className="flex items-center gap-3 border-b border-[#edebe9] px-4 py-2 text-sm">
                <span className="w-12 shrink-0 text-[#605e5c]">Subject</span>
                <input
                  value={subject}
                  onChange={(event) => {
                    setSubject(event.target.value);
                    setSendStatus("idle");
                    setClassification(null);
                  }}
                  className="min-w-0 flex-1 bg-transparent outline-none"
                  placeholder="Subject"
                />
              </label>

              <textarea
                value={body}
                onChange={(event) => setBody(event.target.value)}
                className="min-h-48 flex-1 resize-none px-4 py-3 text-sm leading-6 outline-none"
                placeholder="Write your message"
              />

              <div className="border-t border-[#edebe9] px-4 py-3">
                {sendError ? (
                  <p className="flex items-start gap-2 text-sm text-destructive">
                    <AlertCircle className="mt-0.5 size-4 shrink-0" />
                    {sendError}
                  </p>
                ) : null}
                {sendStatus === "blocked" ? (
                  <p className="flex items-start gap-2 text-sm text-[#8a3707]">
                    <AlertCircle className="mt-0.5 size-4 shrink-0" />
                    Send was stopped. This is a forwarded email — classify it
                    before sending.
                  </p>
                ) : null}
                {sendStatus === "sent" && classification ? (
                  <p className="flex items-start gap-2 text-sm text-[#0e7c3a]">
                    <CheckCircle2 className="mt-0.5 size-4 shrink-0" />
                    Sent after classification: {classification.priority} ·{" "}
                    {classification.endDate} · {classification.category}
                  </p>
                ) : null}
                {sendStatus === "sent" && !classification ? (
                  <p className="flex items-start gap-2 text-sm text-[#0e7c3a]">
                    <CheckCircle2 className="mt-0.5 size-4 shrink-0" />
                    Sent. This was not a forwarded message, so the form was
                    skipped.
                  </p>
                ) : null}
              </div>
            </section>

            {formOpen ? (
              <aside
                aria-label="Classify this forward"
                className="w-full shrink-0 rounded-lg border border-[#edebe9] bg-white p-4 shadow-sm lg:w-72"
              >
                <h2 className="text-sm font-semibold">Classify this forward</h2>
                <p className="mt-1 mb-3 text-xs text-[#605e5c]">
                  Send is paused. Fill in these three fields, then the message can
                  go out.
                </p>
                <ForwardMetadataForm
                  compact
                  submitLabel="Save and send"
                  onSubmit={handleClassify}
                />
              </aside>
            ) : null}
          </div>
        </main>
      </div>
    </div>
  );
}
