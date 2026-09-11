import {
  buildLetterSubmitUrl,
  dummyLetterSubmitFields,
  letterSubmitBaseUrl,
  letterSubmitEnabled,
  letterSubmitFromServer,
  loadLetterLookup,
} from "@/lib/letter-submit";

export const dynamic = "force-dynamic";

const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
};

export function OPTIONS() {
  return new Response(null, { status: 204, headers: cors });
}

export async function GET() {
  return Response.json(
    {
      url: buildLetterSubmitUrl(letterSubmitBaseUrl()),
      lookup: loadLetterLookup(),
      dummyPayload: dummyLetterSubmitFields(),
      submitEnabled: letterSubmitEnabled(),
      submitFromServer: letterSubmitFromServer(),
      submitFrom: letterSubmitFromServer() ? "server" : "browser",
    },
    { headers: cors },
  );
}
