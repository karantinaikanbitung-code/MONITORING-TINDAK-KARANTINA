const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwMquEGFezdJZqG9ci814UsF01BVc4oYn5DS1mCWHTL6YBX1kVjA2hCNCJBr2UYyrC4yg/exec";

export default {
    async fetch(request, env) {
        const url = new URL(request.url);

        // Handle CORS preflight
        if (request.method === "OPTIONS") {
            return new Response(null, {
                headers: {
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
                    "Access-Control-Allow-Headers": "Content-Type",
                },
            });
        }

        const corsHeaders = {
            "Access-Control-Allow-Origin": "*",
        };

        // 1. Check if it's a file operation (R2)
        // Usually these are detected by specific path patterns or headers
        // In previous versions, the pathname was directly used as the R2 key
        const path = url.pathname.slice(1);

        // IF it's a PUT/DELETE or a GET with a filename, handle via R2
        if (request.method === "PUT" || request.method === "DELETE" || (request.method === "GET" && path !== "")) {
            try {
                if (request.method === "PUT") {
                    await env.LHU_BUCKET.put(path, request.body);
                    return new Response(JSON.stringify({ success: true, file: path }), {
                        status: 200,
                        headers: { ...corsHeaders, "Content-Type": "application/json" }
                    });
                }

                if (request.method === "DELETE") {
                    await env.LHU_BUCKET.delete(path);
                    return new Response(JSON.stringify({ success: true, deleted: path }), {
                        status: 200,
                        headers: { ...corsHeaders, "Content-Type": "application/json" }
                    });
                }

                if (request.method === "GET") {
                    const obj = await env.LHU_BUCKET.get(path);
                    if (!obj) return new Response("Not found", { status: 404, headers: corsHeaders });

                    const headers = new Headers();
                    obj.writeHttpMetadata(headers);
                    headers.set("Access-Control-Allow-Origin", "*");
                    headers.set("etag", obj.httpEtag);

                    return new Response(obj.body, { headers });
                }
            } catch (err) {
                return new Response(`R2 Error: ${err.message}`, { status: 500, headers: corsHeaders });
            }
        }

        // 2. Otherwise, treat as Proxy to Google Apps Script
        try {
            const fetchOptions = {
                method: request.method,
                headers: {
                    "Content-Type": request.headers.get("Content-Type") || "application/json"
                }
            };

            if (request.method === "POST") {
                fetchOptions.body = await request.text();
            }

            // Forward to Apps Script
            const response = await fetch(APPS_SCRIPT_URL, fetchOptions);
            const data = await response.text();

            return new Response(data, {
                status: response.status,
                headers: {
                    ...corsHeaders,
                    "Content-Type": "application/json"
                }
            });

        } catch (err) {
            return new Response(`Proxy Error: ${err.message}`, { status: 500, headers: corsHeaders });
        }
    }
};
