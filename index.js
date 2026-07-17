require("dotenv").config();

const axios = require("axios");
const ExcelJS = require("exceljs");

const api = axios.create({
    baseURL: "https://api.phub.centralizedinc.com/api",
    headers: {
        Authorization: `Bearer ${process.env.TOKEN}`
    }
});

async function getIssues() {

    let page = 1;
    let issues = [];

    while (true) {

        const { data } = await api.get("/dev/issues", {
            params: {
                includes: "assignees,tags,comments",
                search: `status:FOR DEPLOYMENT,clientId:${process.env.CLIENT_ID}`,
                page,
                sort: "updatedAt:desc"
            }
        });

        const rows = data.response.rows;

        if (!rows.length) break;

        issues.push(...rows);

        page++;
    }

    return issues;
}

function stripHtml(text = "") {

    return text
        .replace(/<[^>]*>/g, "")
        .replace(/&nbsp;/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function extractResolution(comments = []) {

    const EXCLUDED_USERS = [
        144 // Shiela
    ];

    const cleanedComments = comments
        .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt))
        .filter(c =>
            c.comment?.trim() &&
            !EXCLUDED_USERS.includes(c.createdById)
        )
        .map(c => cleanComment(c.comment));

    const extracted = [];

    for (const comment of cleanedComments) {

        const sections = extractTaggedSections(comment);

        if (sections) {
            extracted.push(sections);
        }

    }

    if (extracted.length) {
        return extracted.join("\n\n");
    }

    return cleanedComments.join("\n\n");
}

function extractTaggedSections(text = "") {

    const regex =
        /(Root Cause|Solution|Resolution)\s*:?\s*([\s\S]*?)(?=(Root Cause|Solution|Resolution)\s*:|$)/gi;

    const sections = [];

    let match;

    while ((match = regex.exec(text)) !== null) {

        const title =
            match[1]
                .trim()
                .toLowerCase()
                .replace(/\b\w/g, c => c.toUpperCase());

        const content =
            match[2]
                .trim();

        sections.push(`${title}:\n${content}`);
    }

    return sections.length
        ? sections.join("\n\n")
        : null;
}

function cleanComment(text = "") {

    return text

        // Remove HTML tags
        .replace(/<[^>]*>/g, " ")

        // Remove URLs
        .replace(/https?:\/\/\S+/gi, "")

        // Remove Merge Request links
        .replace(/MR\s*:?\s*https?:\/\/\S+/gi, "")

        // Remove "Merge Request: ..."
        .replace(/Merge Request\s*:?.*/gi, "")

        // Remove GitLab Merge Request references
        .replace(/!?\d+\s*\(?(merge request|mr)\)?/gi, "")

        // Remove Jira/PHub ticket hyperlinks left after HTML stripping
        .replace(/Reference Ticket\s*:?.*/gi, "")

        // Remove extra spaces/newlines
        .replace(/\n\s*\n/g, "\n")
        .replace(/[ \t]+/g, " ")
        .trim();
}

async function generateExcel() {

    const workbook = new ExcelJS.Workbook();

    const sheet = workbook.addWorksheet("Deployment");

    sheet.columns = [
        { header: "Issue No.", key: "issueNo", width: 22 },
        { header: "Title", key: "title", width: 40 },
        // { header: "Description", key: "description", width: 70 },
        // { header: "Assignee", key: "assignee", width: 25 },
        // { header: "Priority", key: "priority", width: 15 },
        // { header: "Severity", key: "severity", width: 15 },
        // { header: "Tags", key: "tags", width: 20 },
        { header: "Resolution", key: "resolution", width: 80 },
        // { header: "Updated At", key: "updatedAt", width: 25 }
    ];

    const issues = await getIssues();

    for (const issue of issues) {

        sheet.addRow({

            issueNo: issue.code,

            title: issue.name,

            description: stripHtml(issue.description),

            assignee: issue.assignees
                ?.map(a => `${a.fname} ${a.lname}`)
                .join(", "),

            priority: issue.priority,

            severity: issue.severity,

            tags: issue.tags
                ?.map(t => t.name)
                .join(", "),

            resolution: extractResolution(issue.comments),

            updatedAt: issue.updatedAt
        });

    }

    await workbook.xlsx.writeFile("./output/Deployment_Report.xlsx");

    console.log("Deployment_Report.xlsx generated successfully.");
}

generateExcel();