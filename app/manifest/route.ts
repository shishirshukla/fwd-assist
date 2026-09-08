import { buildManifestXml } from "@/lib/manifest";

export const dynamic = "force-dynamic";

export async function GET() {
  const xml = buildManifestXml();

  return new Response(xml, {
    headers: {
      "Content-Type": "application/xml; charset=utf-8",
      "Cache-Control": "no-store, max-age=0",
    },
  });
}
