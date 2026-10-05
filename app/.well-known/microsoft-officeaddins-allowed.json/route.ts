import { getPublicBaseUrl } from "@/lib/manifest";

export const dynamic = "force-dynamic";

export function GET() {
  return Response.json(
    { allowed: [`${getPublicBaseUrl()}/launchevent.js`] },
    { headers: { "Cache-Control": "no-store, max-age=0" } },
  );
}
