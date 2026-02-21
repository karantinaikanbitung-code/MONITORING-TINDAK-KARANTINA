export default {
    async fetch(request, env) {
        const url = new URL(request.url);
        const path = url.pathname.split('/').filter(p => p);

        const corsHeaders = {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, PUT, POST, DELETE, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
        };

        if (request.method === "OPTIONS") {
            return new Response(null, { headers: corsHeaders });
        }

        const respond = (data, status = 200, type = "application/json") => {
            const body = type === "application/json" ? JSON.stringify(data) : data;
            return new Response(body, {
                status,
                headers: {
                    ...corsHeaders,
                    "Content-Type": type
                }
            });
        };

        try {
            let [action, docId, context, filename] = path;

            // Decode parameters to handle spaces and special characters
            docId = docId ? decodeURIComponent(docId) : null;
            context = context ? decodeURIComponent(context) : null;
            filename = filename ? decodeURIComponent(filename) : null;

            if (action === "upload" && request.method === "PUT") {
                if (!docId || !context || !filename) {
                    return respond({ error: "Missing parameters" }, 400);
                }

                // Clear entire prefix before upload to enforce one-file-per-context
                const prefix = `${docId}/${context}/`;
                const list = await env.BUCKET.list({ prefix });
                for (const obj of list.objects) {
                    await env.BUCKET.delete(obj.key);
                }

                const key = `${docId}/${context}/${filename}`;
                await env.BUCKET.put(key, request.body, {
                    httpMetadata: { contentType: request.headers.get("Content-Type") || "application/octet-stream" }
                });

                return respond({ success: true, key });
            }

            if (action === "view" && request.method === "GET") {
                if (!docId || !context || !filename) {
                    return respond({ error: "Missing parameters" }, 400);
                }

                const key = `${docId}/${context}/${filename}`;
                const object = await env.BUCKET.get(key);

                if (!object) {
                    return respond(`Object Not Found: ${key}`, 404, "text/plain");
                }

                const headers = new Headers(corsHeaders);
                object.writeHttpMetadata(headers);
                headers.set("etag", object.httpEtag);

                return new Response(object.body, { headers });
            }

            if (action === "check" && request.method === "GET") {
                if (!docId || !context) {
                    return respond({ error: "Missing parameters" }, 400);
                }

                const prefix = `${docId}/${context}/`;
                const list = await env.BUCKET.list({ prefix });

                if (list.objects.length > 0) {
                    const latest = list.objects[0];
                    const actualFilename = latest.key.split('/').pop();
                    return respond({
                        exists: true,
                        filename: actualFilename,
                        url: `/view/${encodeURIComponent(docId)}/${encodeURIComponent(context)}/${encodeURIComponent(actualFilename)}`
                    });
                }

                return respond({ exists: false });
            }

            if (action === "delete" || (action === "delete" && request.method === "DELETE")) {
                if (!docId || !context) {
                    return respond({ error: "Missing parameters" }, 400);
                }

                // Delete entire prefix to ensure even orphaned/old files are removed
                const prefix = `${docId}/${context}/`;
                const list = await env.BUCKET.list({ prefix });
                for (const obj of list.objects) {
                    await env.BUCKET.delete(obj.key);
                }

                return respond({ success: true, message: "Prefix deleted" });
            }

            return respond({ error: "Not Found" }, 404);
        } catch (err) {
            return respond({ error: err.message }, 500);
        }
    }
};
