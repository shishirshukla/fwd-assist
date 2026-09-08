export const PRIORITIES = ["High", "Medium", "Low"] as const;
export const CATEGORIES = [
  "Work",
  "Personal",
  "Follow-up",
  "Legal",
  "Finance",
  "Other",
] as const;

export type Priority = (typeof PRIORITIES)[number];
export type Category = (typeof CATEGORIES)[number];

export type ForwardMetadata = {
  priority: Priority;
  endDate: string;
  category: Category;
};

export const FORWARD_SUBJECT_PATTERN = /^(fw|fwd)\s*:/i;

export function isForwardedSubject(subject: string): boolean {
  return FORWARD_SUBJECT_PATTERN.test(subject.trim());
}

export function todayIsoDate(): string {
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${now.getFullYear()}-${month}-${day}`;
}

export function validateForwardMetadata(
  value: Partial<ForwardMetadata>,
): Partial<Record<keyof ForwardMetadata, string>> {
  const errors: Partial<Record<keyof ForwardMetadata, string>> = {};

  if (!value.priority) {
    errors.priority = "Choose a priority.";
  }
  if (!value.category) {
    errors.category = "Choose a category.";
  }
  if (!value.endDate) {
    errors.endDate = "Choose an end date.";
  } else if (value.endDate < todayIsoDate()) {
    errors.endDate = "End date cannot be in the past.";
  }

  return errors;
}

export function formatClassificationLine(data: ForwardMetadata): string {
  return `[Forward classification: Priority=${data.priority} | End Date=${data.endDate} | Category=${data.category}]`;
}

export const CUSTOM_PROP_KEYS = {
  complete: "forwardMetadataComplete",
  priority: "forwardPriority",
  endDate: "forwardEndDate",
  category: "forwardCategory",
} as const;
