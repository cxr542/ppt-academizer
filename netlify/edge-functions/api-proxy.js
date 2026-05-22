// Proxy API routes to PPT_ACADEMIZER_API_URL (localtunnel bypass header).
export default async (request, context) => {
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

  return fetch(target, {
    method: request.method,
    headers,
    body:
      request.method === "GET" || request.method === "HEAD"
        ? undefined
        : request.body,
  });
};
