"use client";

import { useMemo, useState } from "react";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  CATEGORIES,
  PRIORITIES,
  type ForwardMetadata,
  todayIsoDate,
  validateForwardMetadata,
} from "@/lib/forward-metadata";

type ForwardMetadataFormProps = {
  compact?: boolean;
  submitting?: boolean;
  submitLabel?: string;
  onSubmit: (data: ForwardMetadata) => void;
};

export function ForwardMetadataForm({
  compact = false,
  submitting = false,
  submitLabel = "Save classification",
  onSubmit,
}: ForwardMetadataFormProps) {
  const [priority, setPriority] = useState<string>("");
  const [endDate, setEndDate] = useState("");
  const [category, setCategory] = useState<string>("");
  const [errors, setErrors] = useState<
    Partial<Record<keyof ForwardMetadata, string>>
  >({});

  const minDate = useMemo(() => todayIsoDate(), []);

  function handleSubmit(event: React.FormEvent) {
    event.preventDefault();
    const draft = {
      priority: priority as ForwardMetadata["priority"],
      endDate,
      category: category as ForwardMetadata["category"],
    };
    const nextErrors = validateForwardMetadata(draft);
    setErrors(nextErrors);
    if (Object.keys(nextErrors).length > 0) {
      return;
    }
    onSubmit(draft as ForwardMetadata);
  }

  const fieldClass =
    "h-8 w-full rounded-lg border border-input bg-background px-2.5 text-sm outline-none focus-visible:border-ring focus-visible:ring-3 focus-visible:ring-ring/50";

  return (
    <form onSubmit={handleSubmit} className={compact ? "space-y-3" : "space-y-4"}>
      <div className="space-y-1.5">
        <Label htmlFor="forward-priority">Priority</Label>
        <select
          id="forward-priority"
          value={priority}
          onChange={(event) => setPriority(event.target.value)}
          aria-invalid={Boolean(errors.priority)}
          className={fieldClass}
        >
          <option value="">Select priority</option>
          {PRIORITIES.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
        {errors.priority ? (
          <p className="text-xs text-destructive">{errors.priority}</p>
        ) : null}
      </div>

      <div className="space-y-1.5">
        <Label htmlFor="forward-end-date">End Date</Label>
        <Input
          id="forward-end-date"
          type="date"
          min={minDate}
          value={endDate}
          onChange={(event) => setEndDate(event.target.value)}
          aria-invalid={Boolean(errors.endDate)}
        />
        {errors.endDate ? (
          <p className="text-xs text-destructive">{errors.endDate}</p>
        ) : null}
      </div>

      <div className="space-y-1.5">
        <Label htmlFor="forward-category">Category</Label>
        <select
          id="forward-category"
          value={category}
          onChange={(event) => setCategory(event.target.value)}
          aria-invalid={Boolean(errors.category)}
          className={fieldClass}
        >
          <option value="">Select category</option>
          {CATEGORIES.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
        {errors.category ? (
          <p className="text-xs text-destructive">{errors.category}</p>
        ) : null}
      </div>

      <Button type="submit" className="w-full" disabled={submitting}>
        {submitting ? "Saving…" : submitLabel}
      </Button>
    </form>
  );
}
