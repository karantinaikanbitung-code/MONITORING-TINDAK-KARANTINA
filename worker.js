export default {
    async fetch(request, env) {
        const url = new URL(request.url);
        const path = url.pathname.split('/').filter(p => p);

        // Usage: /upload/:docId/:context/:filename
        // Usage: /view/:docId/:context/:filename
        // Usage: /check/:docId/:context

        let [action, docId, context, filename] = path;

        // Decode parameters to handle spaces and special characters
        docId = docId ? decodeURIComponent(docId) : null;
        context = context ? decodeURIComponent(context) : null;
        filename = filename ? decodeURIComponent(filename) : null;

        // Handle CORS
        if (request.method === "OPTIONS") {
            return new Response(null, {
                headers: {
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "GET, PUT, OPTIONS",
                    "Access-Control-Allow-Headers": "Content-Type",
                },
            });
        }

        if (action === "upload" && request.method === "PUT") {
            if (!docId || !context || !filename) {
                return new Response("Missing parameters", { status: 400 });
            }

            const key = `${docId}/${context}/${filename}`;
            await env.BUCKET.put(key, request.body, {
                httpMetadata: { contentType: request.headers.get("Content-Type") || "application/octet-stream" }
            });

            return new Response(JSON.stringify({ success: true, key }), {
                headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
            });
        }

        if (action === "view" && request.method === "GET") {
            if (!docId || !context || !filename) {
                return new Response("Missing parameters", { status: 400 });
            }

            const key = `${docId}/${context}/${filename}`;
            const object = await env.BUCKET.get(key);

            if (!object) {
                return new Response(`Object Not Found: ${key}`, { status: 404, headers: { "Access-Control-Allow-Origin": "*" } });
            }

            const headers = new Headers();
            object.writeHttpMetadata(headers);
            headers.set("etag", object.httpEtag);
            headers.set("Access-Control-Allow-Origin", "*");

            return new Response(object.body, { headers });
        }

        if (action === "check" && request.method === "GET") {
            if (!docId || !context) {
                return new Response("Missing parameters", { status: 400 });
            }

            const prefix = `${docId}/${context}/`;
            const list = await env.BUCKET.list({ prefix });

            if (list.objects.length > 0) {
                const latest = list.objects[0]; // Assuming one file per context
                const actualFilename = latest.key.split('/').pop();
                return new Response(JSON.stringify({
                    exists: true,
                    filename: actualFilename,
                    url: `/view/${encodeURIComponent(docId)}/${encodeURIComponent(context)}/${encodeURIComponent(actualFilename)}`
                }), {
                    headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
                });
            }

            return new Response(JSON.stringify({ exists: false }), {
                headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
            });
        }

        return new Response("Not Found", { status: 404 });
    }
};
