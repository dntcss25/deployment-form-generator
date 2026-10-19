require("dotenv").config();

const http = require("node:http");
const { generateDeploymentReport } = require("./index");

const PORT = Number(process.env.PORT) || 3001;
const ALLOWED_ORIGIN = process.env.CCC_SUPPORT_ORIGIN || "http://localhost:5173";
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const server = http.createServer(async (request, response) => {
    const origin = request.headers.origin;
    if (!origin || origin === ALLOWED_ORIGIN) {
        response.setHeader("Access-Control-Allow-Origin", origin || ALLOWED_ORIGIN);
        response.setHeader("Vary", "Origin");
        response.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
        response.setHeader("Access-Control-Allow-Headers", "Content-Type");
    }

    if (request.method === "OPTIONS") {
        response.writeHead(204).end();
        return;
    }

    if (request.method === "GET" && request.url === "/api/health") {
        response.writeHead(200, { "Content-Type": "application/json" });
        response.end(JSON.stringify({ ok: true }));
        return;
    }

    if (request.method !== "POST" || request.url !== "/api/deployment-report") {
        response.writeHead(404, { "Content-Type": "application/json" });
        response.end(JSON.stringify({ error: "Not found." }));
        return;
    }

    try {
        let rawBody = "";
        for await (const chunk of request) {
            rawBody += chunk;
            if (rawBody.length > 10_000) throw new Error("Request body is too large.");
        }

        const { application, tag } = JSON.parse(rawBody || "{}");
        if (!application || !["PROD", "STAGING", "ALL"].includes(tag)) {
            response.writeHead(400, { "Content-Type": "application/json" });
            response.end(JSON.stringify({ error: "Choose an application and deployment tag." }));
            return;
        }

        const result = await generateDeploymentReport(application, tag === "ALL" ? null : tag);
        response.writeHead(200, {
            "Content-Type": XLSX_TYPE,
            "Content-Disposition": `attachment; filename="${result.filename}"`,
            "Content-Length": result.buffer.length,
            "Cache-Control": "no-store"
        });
        response.end(Buffer.from(result.buffer));
    } catch (error) {
        const status = error.message === "Missing TOKEN in .env" ? 503 : 500;
        console.error("Deployment report request failed:", error.response?.status || error.message);
        response.writeHead(status, { "Content-Type": "application/json" });
        response.end(JSON.stringify({
            error: status === 503
                ? "The report service is missing its API token."
                : "Unable to generate the report. Check the backend logs."
        }));
    }
});

server.listen(PORT, () => {
    console.log(`Deployment report API listening on http://localhost:${PORT}`);
});
