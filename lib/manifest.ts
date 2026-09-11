import { readFileSync } from "node:fs";
import { join } from "node:path";

export const DEFAULT_BASE = "https://localhost:43123";
const manifestTemplatePath = join(process.cwd(), "public", "manifest.template.xml");

export function normalizeBaseUrl(input: string): string {
  const trimmed = input.trim().replace(/\/+$/, "");
  if (!/^https:\/\/.+/i.test(trimmed)) {
    throw new Error("Base URL must start with https://");
  }
  return trimmed;
}

export function resolvePublicBaseUrl(): string | null {
  const candidates = [
    process.env.PUBLIC_BASE_URL,
    process.env.RAILWAY_STATIC_URL,
    process.env.RAILWAY_PUBLIC_DOMAIN
      ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN.replace(/^https?:\/\//, "")}`
      : undefined,
  ].filter(Boolean) as string[];

  for (const value of candidates) {
    try {
      const withScheme = /^https?:\/\//i.test(value)
        ? value
        : `https://${value}`;
      return normalizeBaseUrl(withScheme);
    } catch {
      continue;
    }
  }

  return null;
}

export function getPublicBaseUrl(): string {
  return resolvePublicBaseUrl() ?? DEFAULT_BASE;
}

export function hostFromBaseUrl(baseUrl: string): string {
  return new URL(baseUrl).hostname;
}

export function injectManifestUrls(template: string, baseUrl: string): string {
  const base = normalizeBaseUrl(baseUrl);
  const host = hostFromBaseUrl(base);

  let xml = template.replace(/https:\/\/[^"<]+/g, (match) => {
    const path = match.replace(/^https:\/\/[^/]+/, "");
    return base + path;
  });

  xml = xml.replace(
    /<AppDomain>[^<]+<\/AppDomain>/,
    `<AppDomain>${host}</AppDomain>`,
  );

  return xml;
}

export function isRailwayRuntime(): boolean {
  return Boolean(
    process.env.RAILWAY_ENVIRONMENT ||
      process.env.RAILWAY_PROJECT_ID ||
      process.env.RAILWAY_SERVICE_ID,
  );
}

export function loadManifestTemplate(): string {
  return readFileSync(/* turbopackIgnore: true */ manifestTemplatePath, "utf8");
}

export function buildManifestXml(baseUrl = getPublicBaseUrl()): string {
  return injectManifestUrls(loadManifestTemplate(), baseUrl);
}
