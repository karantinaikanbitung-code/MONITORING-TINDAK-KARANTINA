export default {
    async fetch(request, env) {
        const url = "https://script.google.com/macros/s/AKfycbwMquEGFezdJZqG9ci814UsF01BVc4oYn5DS1mCWHTL6YBX1kVjA2hCNCJBr2UYyrC4yg/exec";

        // Add CORS headers for preflight
        if (request.method === "OPTIONS") {
            return new Response(null, {
                headers: {
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
                    "Access-Control-Allow-Headers": "Content-Type",
                },
            });
        }

        // Proxy request to GAS
        const response = await fetch(url, {
            method: request.method,
            headers: request.headers,
            body: request.body,
        });

        // Mirror response with CORS
        const newHeaders = new Headers(response.headers);
        newHeaders.set("Access-Control-Allow-Origin", "*");

        return new Response(response.body, {
            status: response.status,
            statusText: response.statusText,
            headers: newHeaders,
        });
    },
};
