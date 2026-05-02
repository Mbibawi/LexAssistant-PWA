const CLAUDE_API_KEY = process.env.CLAUDE_API_KEY;
const FOUNDRY_API_KEY = process.env.FOUNDRY_API_KEY;

export const main = async (req:any, res:any) => {
    setCorsHeaders(res);

    if (req.method === "OPTIONS") {
        res.status(204).send("");
        return;
    }

    const path = req.path || "/";

    if (req.method === "GET") {
        res.status(200).send("Claude proxy is up and running.");
    } else if (path === "/api/proxy/azure" && req.method === "POST") {
        await handleClaudeProxy(req, res, true);
    } else if (path === "/api/proxy/claude" && req.method === "POST") {
        await handleClaudeProxy(req, res);
    } else {
        res.status(404).json({ error: `No handler for path: ${path}` });
    }
};

// ─── Handlers ─────────────────────────────────────────────────────────────────

async function handleClaudeProxy(req:any, res:any, azure:boolean = false) {
    const claudePath = req.body?.path ;
    if (!claudePath) {
        res.status(400).json({ error: "Missing body parameter: path" });
        return;
    }
    if (!req.body.path || !req.body.messages) throw new Error("Missing body parameter: path or messages");
    try {
        const upstream = azure ? await callAzureFoundry(req) : await callClaudeAPI(req);
        
        res.status(upstream.status);
        mirrorResponseHeaders(upstream.headers, res);

        const contentType = upstream.headers.get("content-type") || "";
        if (contentType.includes("text/event-stream")) {
            const reader = upstream.body!.getReader();
            const decoder = new TextDecoder();
            while (true) {
                const { done, value } = await reader!.read();
                if (done) break;
                res.write(decoder.decode(value, { stream: true }));
            }
            res.send();
        } else {
            const body = await upstream.text();
            res.send(body);
        }
    } catch (err:any) {
        console.error("Proxy error:", err.message);
        console.error("Cause:", err.cause?.message);
        console.error("Cause code:", err.cause?.code);
        res.status(502).json({ error: "Bad gateway", detail: err.message });
    }
}

async function callClaudeAPI(req: any) {
    //Claude API
    if (!CLAUDE_API_KEY) throw new Error("Missing CLAUDE Key");
    
    const url = `https://api.anthropic.com/${req.body.path}`;
    return await fetchAPI(url, CLAUDE_API_KEY, req);
}

async function callAzureFoundry(req:any) {
    //Azure AI Foundry
    if (!FOUNDRY_API_KEY) throw new Error("Missing AZURE Key");
    
    const url = `https://${process.env.AZURE_RESSOURCE}.services.ai.azure.com/anthropic/${req.body.path}`;
    return await fetchAPI(url, FOUNDRY_API_KEY, req);

}

async function fetchAPI(url:string, apiKey:string, req:any) {
    const { method, headers, body } = req;
    headers["x-api-key"] = apiKey;
    return await fetch(url, {
        method: method,
        headers: headers,
        body: body.messages
    });
}


function mirrorResponseHeaders(upstreamHeaders:any, res:any) {
    const skip = new Set(["transfer-encoding", "connection"]);
    for (const [key, value] of upstreamHeaders.entries()) {
        if (!skip.has(key.toLowerCase())) {
            res.setHeader(key, value);
        }
    }
}

function setCorsHeaders(res:any) {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, anthropic-version, anthropic-beta");
}