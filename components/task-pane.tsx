"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { AlertCircle, CheckCircle2 } from "lucide-react";

import { ForwardMetadataForm } from "@/components/forward-metadata-form";
import {
  CUSTOM_PROP_KEYS,
  formatClassificationLine,
  type ForwardMetadata,
} from "@/lib/forward-metadata";

type PaneState =
  | { kind: "loading" }
  | { kind: "browser" }
  | { kind: "ready" }
  | { kind: "saved"; data: ForwardMetadata }
  | { kind: "sent"; data: ForwardMetadata }
  | { kind: "error"; message: string };

export function TaskPane() {
  const [state, setState] = useState<PaneState>({ kind: "loading" });
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    let cancelled = false;
    const Office = window.Office;

    if (!Office) {
      const timer = window.setTimeout(() => {
        if (!cancelled) {
          setState({ kind: "browser" });
        }
      }, 0);
      return () => {
        cancelled = true;
        window.clearTimeout(timer);
      };
    }

    Office.onReady((info) => {
      if (cancelled) {
        return;
      }
      setState(info.host === "Outlook" ? { kind: "ready" } : { kind: "browser" });
    });

    return () => {
      cancelled = true;
    };
  }, []);

  function sendAfterSave(
    item: Office.MailboxItem,
    data: ForwardMetadata,
  ) {
    if (typeof item.sendAsync !== "function") {
      setSaving(false);
      setState({ kind: "saved", data });
      return;
    }
    item.sendAsync((sendResult) => {
      setSaving(false);
      if (sendResult.status === Office.AsyncResultStatus.Succeeded) {
        setState({ kind: "sent", data });
        return;
      }
      setState({ kind: "saved", data });
    });
  }

  function saveMetadata(data: ForwardMetadata) {
    const Office = window.Office;
    const item = Office?.context?.mailbox?.item;
    if (!Office || !item) {
      setState({ kind: "error", message: "No compose item is available." });
      return;
    }

    setSaving(true);
    item.loadCustomPropertiesAsync((propResult) => {
      if (propResult.status !== Office.AsyncResultStatus.Succeeded) {
        setSaving(false);
        setState({
          kind: "error",
          message: "Could not save classification on this message.",
        });
        return;
      }

      const props = propResult.value;
      props.set(CUSTOM_PROP_KEYS.complete, "true");
      props.set(CUSTOM_PROP_KEYS.priority, data.priority);
      props.set(CUSTOM_PROP_KEYS.endDate, data.endDate);
      props.set(CUSTOM_PROP_KEYS.category, data.category);
      props.saveAsync((saveResult) => {
        if (saveResult.status !== Office.AsyncResultStatus.Succeeded) {
          setSaving(false);
          setState({
            kind: "error",
            message: "Could not persist classification. Try again.",
          });
          return;
        }

        const line = formatClassificationLine(data);
        item.body.prependAsync(
          `<p>${line}</p>`,
          { coercionType: Office.CoercionType.Html },
          () => {
            if (item.internetHeaders?.setAsync) {
              item.internetHeaders.setAsync(
                {
                  "X-Forward-Priority": data.priority,
                  "X-Forward-End-Date": data.endDate,
                  "X-Forward-Category": data.category,
                },
                () => sendAfterSave(item, data),
              );
              return;
            }
            sendAfterSave(item, data);
          },
        );
      });
    });
  }

  if (state.kind === "loading") {
    return (
      <p className="text-sm text-muted-foreground">Connecting to Outlook…</p>
    );
  }

  if (state.kind === "browser") {
    return (
      <div className="space-y-3">
        <p className="text-sm text-muted-foreground">
          This pane is designed to run inside Outlook. Use the simulator on the
          home page to try the Send intercept in a browser.
        </p>
        <Link className="text-sm text-[#0f6cbd] underline" href="/">
          Open Outlook simulator
        </Link>
      </div>
    );
  }

  if (state.kind === "error") {
    return (
      <p className="flex items-start gap-2 text-sm text-destructive">
        <AlertCircle className="mt-0.5 size-4 shrink-0" />
        {state.message}
      </p>
    );
  }

  if (state.kind === "saved" || state.kind === "sent") {
    return (
      <div className="space-y-3 text-sm">
        <p className="flex items-start gap-2 text-[#0e7c3a]">
          <CheckCircle2 className="mt-0.5 size-4 shrink-0" />
          {state.kind === "sent"
            ? "Classification saved. Message sent."
            : "Classification saved. Click Send if the message is still open."}
        </p>
        <ul className="space-y-1 text-muted-foreground">
          <li>Priority: {state.data.priority}</li>
          <li>End Date: {state.data.endDate}</li>
          <li>Category: {state.data.category}</li>
        </ul>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <p className="text-sm text-muted-foreground">
        Send is blocked for this forwarded email until these fields are saved.
        Saving sends the message.
      </p>
      <ForwardMetadataForm
        compact
        submitting={saving}
        submitLabel="Save and send"
        onSubmit={saveMetadata}
      />
    </div>
  );
}
