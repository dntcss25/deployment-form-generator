const { generateDeploymentReport } = require("../index");

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

module.exports = async function handler(request, response) {
    const allowedOrigin = process.env.CCC_SUPPORT_ORIGIN;
    const origin = request.headers.origin;

    if (allowedOrigin && origin === allowedOrigin) {
        response.setHeader("Access-Control-Allow-Origin", allowedOrigin);
        response.setHeader("Vary", "Origin");
        response.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
        response.setHeader("Access-Control-Allow-Headers", "Content-Type");
    }

    if (request.method === "OPTIONS") {
        response.status(204).end();
        return;
    }

    if (request.method !== "POST") {
        response.status(405).json({ error: "Method not allowed." });
        return;
    }

    const { application, tag } = request.body ?? {};
    if (!application || !["PROD", "STAGING", "CLOUD", "ALL"].includes(tag)) {
        response.status(400).json({ error: "Choose an application and deployment tag." });
        return;
    }

    try {
        const report = await generateDeploymentReport(application, tag === "ALL" ? null : tag);
        response.setHeader("Content-Type", XLSX_TYPE);
        response.setHeader("Content-Disposition", `attachment; filename="${report.filename}"`);
        response.setHeader("Cache-Control", "no-store");
        response.status(200).send(Buffer.from(report.buffer));
    } catch (error) {
        const status = error.message === "Missing TOKEN in .env" ? 503 : 500;
        console.error("Deployment report request failed:", error.response?.status || error.message);
        response.status(status).json({
            error: status === 503
                ? "The report service is missing its API token."
                : "Unable to generate the report. Check the backend logs."
        });
    }
};
