import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

export const DEFAULT_BASE = "https://localhost:43123";
export const manifestTemplatePath = resolve("public/manifest.template.xml");

export function normalizeBaseUrl(input) {
  const trimmed = input.trim().replace(/\/+$/, "");
  if (!/^https:\/\/.+/i.test(trimmed)) {
    throw new Error("Base URL must start with https:// (Outlook requires HTTPS).");
  }
  return trimmed;
}

export function hostFromUrl(url) {
  return new URL(url).hostname;
}

export function replaceManifestUrls(xml, nextBase) {
  const host = hostFromUrl(nextBase);
  let updated = xml.replace(/https:\/\/[^"<]+/g, (match) => {
    const path = match.replace(/^https:\/\/[^/]+/, "");
    return nextBase + (path || "");
  });
  updated = updated.replace(
    /<AppDomain>[^<]+<\/AppDomain>/,
    `<AppDomain>${host}</AppDomain>`,
  );
  return updated;
}

export function resolvePublicBaseUrl() {
  const candidates = [
    process.env.PUBLIC_BASE_URL,
    process.env.RAILWAY_STATIC_URL,
    process.env.RAILWAY_PUBLIC_DOMAIN
      ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN.replace(/^https?:\/\//, "")}`
      : undefined,
  ].filter(Boolean);

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

export function getPublicBaseUrl() {
  return resolvePublicBaseUrl() ?? DEFAULT_BASE;
}

export function isRailwayRuntime() {
  return Boolean(
    process.env.RAILWAY_ENVIRONMENT ||
      process.env.RAILWAY_PROJECT_ID ||
      process.env.RAILWAY_SERVICE_ID,
  );
}

export function loadManifestTemplate() {
  return readFileSync(manifestTemplatePath, "utf8");
}

export function buildManifestXml(baseUrl = getPublicBaseUrl()) {
  return replaceManifestUrls(loadManifestTemplate(), baseUrl);
}

export function assertManifestConfigured() {
  if (!isRailwayRuntime()) {
    return;
  }

  if (!resolvePublicBaseUrl()) {
    console.warn(
      "PUBLIC_BASE_URL is not set. Manifest URLs will fall back to localhost. " +
        "Set PUBLIC_BASE_URL=https://your-app.up.railway.app in Railway Variables.",
    );
  }
}

export function writeManifestForBase(nextBase, outputPath = resolve("public/manifest.xml")) {
  const updated = buildManifestXml(nextBase);
  writeFileSync(outputPath, updated, "utf8");
  return { nextBase, host: hostFromUrl(nextBase) };
}
