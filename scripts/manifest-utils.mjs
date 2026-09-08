import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

export const DEFAULT_BASE = "https://localhost:43123";
export const manifestPath = resolve("public/manifest.xml");

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
  if (process.env.PUBLIC_BASE_URL) {
    return normalizeBaseUrl(process.env.PUBLIC_BASE_URL);
  }
  if (process.env.RAILWAY_PUBLIC_DOMAIN) {
    return normalizeBaseUrl(`https://${process.env.RAILWAY_PUBLIC_DOMAIN}`);
  }
  return null;
}

export function writeManifestForBase(nextBase) {
  const current = readFileSync(manifestPath, "utf8");
  const updated = replaceManifestUrls(current, nextBase);
  if (updated !== current) {
    writeFileSync(manifestPath, updated, "utf8");
  }
  return { nextBase, host: hostFromUrl(nextBase), changed: updated !== current };
}
