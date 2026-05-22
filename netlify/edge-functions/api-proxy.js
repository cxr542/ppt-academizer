// Proxy API routes to PPT_ACADEMIZER_API_URL (localtunnel bypass + CORS for GitHub Pages).
const CORS_ORIGIN = "*";

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": CORS_ORIGIN,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers":
      "Content-Type, bypass-tunnel-reminder, X-Requested-With",
    "Access-Control-Expose-Headers":
      "X-Academize-Warnings, X-Academize-Slide-Count, X-Academize-Profile, X-Academize-Pipeline, X-Academize-Source-Format, Content-Disposition",
  };
}

function withCors(response) {
  const headers = new Headers(response.headers);
  for (const [k, v] of Object.entries(corsHeaders())) {
    headers.set(k, v);
  }
  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

export default async (request, context) => {
  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: corsHeaders() });
  }

  const apiBase = Netlify.env.get("PPT_ACADEMIZER_API_URL");
  if (!apiBase) {
    return context.next();
  }

  const incoming = new URL(request.url);
  const target =
    apiBase.replace(/\/$/, "") + incoming.pathname + incoming.search;

  const headers = new Headers(request.headers);
  headers.set("bypass-tunnel-reminder", "true");
  headers.delete("host");

  const upstream = await fetch(target, {
    method: request.method,
    headers,
    body:
      request.method === "GET" || request.method === "HEAD"
        ? undefined
        : request.body,
  });

  return withCors(upstream);
};
