const CLAUDE_API_KEY = process.env.CLAUDE_API_KEY;
const CLAUDE_BASE_URL = "https://api.anthropic.com";

export const main = async (req, res) => {
    setCorsHeaders(res);

    if (req.method === "OPTIONS") {
        res.status(204).send("");
        return;
    }

    const path = req.path || "/";

    if (path === "/" && req.method === "GET") {
        res.status(200).send("Claude proxy is up and running.");
    } else if (path.startsWith("/api/proxy")) {
        await handleClaudeProxy(req, res);
    } else {
        res.status(404).json({ error: `No handler for path: ${path}` });
    }
};

// ─── Handlers ─────────────────────────────────────────────────────────────────

async function handleClaudeProxy(req, res) {
    const body = JSON.parse(req.body);
    const { claudePath, messages } = body;
    if (!claudePath) {
        res.status(400).json({ error: "Missing body parameter: path" });
        return;
    }

    const claudeUrl = `${CLAUDE_BASE_URL}${claudePath}`;

    try {
        const upstream = await fetch(claudeUrl, {
            method: req.method,
            headers: buildForwardHeaders(req.headers),
            body: JSON.stringify(messages)
        });

        res.status(upstream.status);
        mirrorResponseHeaders(upstream.headers, res);

        const contentType = upstream.headers.get("content-type") || "";
        if (contentType.includes("text/event-stream")) {
            const reader = upstream.body.getReader();
            const decoder = new TextDecoder();
            while (true) {
                const { done, value } = await reader.read();
                if (done) break;
                res.write(decoder.decode(value, { stream: true }));
            }
            res.end();
        } else {
            const body = await upstream.text();
            res.send(body);
        }
    } catch (err) {
        console.error("Proxy error:", err.message);
        console.error("Cause:", err.cause?.message);
        console.error("Cause code:", err.cause?.code);
        res.status(502).json({ error: "Bad gateway", detail: err.message, headers: buildForwardHeaders(req.headers) });
    }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function buildForwardHeaders(incoming) {
    const strip = new Set([
        "host", "connection", "transfer-encoding",
        "te", "trailer", "upgrade",
        "x-forwarded-for", "x-forwarded-host", "x-forwarded-proto",
        "x-cloud-trace-context", "traceparent",
        "content-length",
        "x-path"
    ]);

    const headers = {};
    for (const [key, value] of Object.entries(incoming)) {
        if (strip.has(key.toLowerCase())) continue;
        headers[key] = value;
    }

    headers["x-api-key"] = CLAUDE_API_KEY;
    headers["content-type"] = headers["content-type"] || "application/json";

    return headers;
}

function mirrorResponseHeaders(upstreamHeaders, res) {
    const skip = new Set(["transfer-encoding", "connection"]);
    for (const [key, value] of upstreamHeaders.entries()) {
        if (!skip.has(key.toLowerCase())) {
            res.setHeader(key, value);
        }
    }
}

function setCorsHeaders(res) {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, anthropic-version, anthropic-beta");
}