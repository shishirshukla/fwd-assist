import {
  getPublicBaseUrl,
  isRailwayRuntime,
  resolvePublicBaseUrl,
} from "@/lib/manifest";

export const dynamic = "force-dynamic";

export async function GET() {
  return Response.json({
    publicBaseUrl: getPublicBaseUrl(),
    configuredBaseUrl: resolvePublicBaseUrl(),
    railway: isRailwayRuntime(),
    manifestUrl: `${getPublicBaseUrl()}/manifest.xml`,
  });
}
