import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";

// Enable streaming for 2026 AI workflows
app.setup({ enableHttpStream: true });

const CLAUDE_API_KEY = process.env.CLAUDE_API_KEY;
const FOUNDRY_API_KEY = process.env.FOUNDRY_API_KEY;
const AZURE_RESOURCE = process.env.AZURE_RESOURCE;

/**
 * MAIN DISPATCHER 
 * This mimics your GCF 'main' export.
 */
export async function main(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    const url = new URL(request.url);
    const path = url.pathname; // e.g., /api/proxy/azure or /api/proxy/claude

    // 1. Handle CORS (Your setCorsHeaders logic)
    if (request.method === "OPTIONS") {
        return { status: 204, headers: getCorsHeaders() };
    }

    // 2. Original Dispatching Logic
    if (request.method === "GET") {
        return { status: 200, body: "Claude proxy is up and running." };
    }

    if (path === "/api/proxy/azure" && request.method === "POST") {
        return await handleClaudeProxy(request, context, true);
    }

    if (path === "/api/proxy/claude" && request.method === "POST") {
        return await handleClaudeProxy(request, context, false);
    }

    // Future-proofing: Add more paths here
    // if (path === "/api/proxy/other") { return await handleOther(request); }

    return {
        status: 404,
        jsonBody: { error: `No handler for path: ${path}` },
        headers: getCorsHeaders()
    };
}

// ─── Handlers ─────────────────────────────────────────────────────────────────

async function handleClaudeProxy(request: HttpRequest, context: InvocationContext, azure: boolean): Promise<HttpResponseInit> {
    try {
        const pwaBody: any = await request.json();

        // Validate the "baby" exists
        if (!pwaBody?.path || !pwaBody?.messages) {
            return { status: 400, jsonBody: { error: "Missing body parameter: path or messages" } };
        }

        let upstreamResponse: Response;
        if (azure) {
            upstreamResponse = await callAzureFoundry(pwaBody, request);
        } else {
            upstreamResponse = await callClaudeAPI(pwaBody, request);
        }

        // Mirror headers & setup response
        const responseHeaders = getCorsHeaders();
        const skip = new Set(["transfer-encoding", "connection", "content-encoding"]);
        upstreamResponse.headers.forEach((v, k) => {
            if (!skip.has(k.toLowerCase())) responseHeaders[k] = v;
        });

        return {
            status: upstreamResponse.status,
            headers: responseHeaders,
            body: upstreamResponse.body // Pass-thru the stream automatically
        };

    } catch (err: any) {
        context.error("Proxy error:", err.message);
        return {
            status: 502,
            jsonBody: { error: "Bad gateway", detail: err.message },
            headers: getCorsHeaders()
        };
    }
}

async function callClaudeAPI(pwaBody: any, request: HttpRequest) {
    if (!CLAUDE_API_KEY) throw new Error("Missing CLAUDE Key");
    const url = `https://api.anthropic.com/${pwaBody.path}`;
    return await fetchAPI(url, CLAUDE_API_KEY, pwaBody, request, false);
}

async function callAzureFoundry(pwaBody: any, request: HttpRequest) {
    if (!FOUNDRY_API_KEY) throw new Error("Missing AZURE Key");
    const url = `https://${AZURE_RESOURCE}.services.ai.azure.com/anthropic/${pwaBody.path}`;
    return await fetchAPI(url, FOUNDRY_API_KEY, pwaBody, request, true);
}

async function fetchAPI(url: string, apiKey: string, pwaBody: any, request: HttpRequest, isAzure: boolean) {
    const headers: Record<string, string> = {
        "Content-Type": "application/json",
        "anthropic-version": request.headers.get("anthropic-version") || "2023-06-01"
    };

    // Swap header key based on provider
    if (isAzure) {
        headers["api-key"] = apiKey;
    } else {
        headers["x-api-key"] = apiKey;
    }

    return await fetch(url, {
        method: "POST",
        headers: headers,
        body: JSON.stringify(pwaBody) // The "Baby": Sent exactly as received
    });
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getCorsHeaders(): Record<string, string> {
    return {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, anthropic-version, anthropic-beta, x-api-key, api-key"
    };
}

// Register the function with a wildcard route to allow the dispatcher to work
app.http('dispatcher', {
    methods: ['GET', 'POST', 'OPTIONS'],
    authLevel: 'anonymous',
    route: 'api/proxy/{*remainder}',
    handler: main
});