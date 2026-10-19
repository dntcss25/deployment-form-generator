
require("dotenv").config();

const axios = require("axios");
const ExcelJS = require("exceljs");
const path = require("path");
const readline = require("readline/promises");

// API configuration
const api = axios.create({
    baseURL: "https://api.phub.centralizedinc.com/api",
    headers: {
        Authorization: `Bearer ${process.env.TOKEN}`
    }
});

// Application configuration
const APPLICATIONS = {
    MARSU: {
        clientId: "6",
        projectId: "82"
    },
    MSUN: {
        clientId: "4",
        projectId: null
    },
    MSUG: {
        clientId: "7",
        projectId: null
    }
};


// Ask user to select an application
async function selectApplication(rl) {
    const apps = Object.keys(APPLICATIONS);

    console.log("\n========== SELECT APPLICATION ==========");

    apps.forEach((app, index) => {
        console.log(`${index + 1}. ${app}`);
    });

    while (true) {
        const answer = (
            await rl.question("\nSelect application (1-3): ")
        ).trim();

        const index = Number(answer) - 1;

        if (Number.isInteger(index) && index >= 0 && index < apps.length) {
            const appName = apps[index];
            const app = APPLICATIONS[appName];

            if (!app.clientId) {
                throw new Error(
                    `Missing CLIENT_ID for ${appName}. ` +
                    "Please configure MSUG_CLIENT_ID in .env."
                );
            }

            return {
                name: appName,
                ...app
            };
        }

        console.log("Invalid selection. Please choose 1, 2, or 3.");
    }
}

// Ask user to select a deployment tag
async function selectDeployTag(rl) {
    console.log("\n========== SELECT DEPLOY TAG ==========");
    console.log("1. PROD");
    console.log("2. STAGING");
    console.log("3. NONE (all tags)");

    const tagOptions = {
        "1": "PROD",
        "2": "STAGING",
        "3": null
    };

    while (true) {
        const answer = (
            await rl.question(
                "\nSelect deployment tag (1-3): "
            )
        ).trim();

        if (Object.prototype.hasOwnProperty.call(tagOptions, answer)) {
            return tagOptions[answer];
        }

        console.log("Invalid selection. Please choose 1, 2, or 3.");
    }
}

// Fetch deployment issues
async function getIssues(app, deployTag) {
    let page = 1;
    const issues = [];

    while (true) {
        const searchParts = [
            "status:FOR DEPLOYMENT",
            `clientId:${app.clientId}`
        ];

        if (app.projectId) {
            searchParts.push(`projectId:${app.projectId}`);
        }

        const { data } = await api.get("/dev/issues", {
            params: {
                includes: "assignees,tags,comments",
                search: searchParts.join(","),
                page,
                sort: "updatedAt:desc"
            }
        });

        const rows = data?.response?.rows ?? [];

        if (!rows.length) {
            break;
        }

        issues.push(...rows);
        console.log(`Fetched page ${page}: ${rows.length} issues`);

        page++;
    }

    // Filter by tag only when a specific tag was selected.
    // NONE means all tags, including issues without tags.
    if (deployTag !== null) {
        return issues.filter(issue =>
            issue.tags?.some(tag =>
                tag.name?.toLowerCase() === deployTag.toLowerCase()
            )
        );
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


async function getLinkedIssue(issueId) {
    const { data } = await api.get(`/dev/issues/${issueId}`, {
        params: {
            includes: "assignees,tags,comments"
        }
    });

    return data?.response ?? data;
}

async function extractResolution(comments = [], visited = new Set()) {
    const EXCLUDED_USERS = [144];

    const sortedComments = [...comments]
        .sort((a, b) =>
            new Date(a.createdAt) - new Date(b.createdAt)
        )
        .filter(comment =>
            comment.comment?.trim() &&
            !EXCLUDED_USERS.includes(comment.createdById)
        );

    const results = [];

    for (const item of sortedComments) {
        const text = item.comment;

        // Detect PHub issue links.
        const linkRegex =
            /https?:\/\/phub\.centralizedinc\.com\/app\/issues\/details\?[^ \s"'<>]*\bid=(\d+)/gi;

        const linkedIds = [
            ...text.matchAll(linkRegex)
        ].map(match => match[1]);

        if (linkedIds.length > 0) {
            for (const linkedId of [...new Set(linkedIds)]) {
                if (visited.has(linkedId)) continue;

                visited.add(linkedId);

                try {
                    const linkedIssue = await getLinkedIssue(linkedId);

                    const linkedResolution = await extractResolution(
                        linkedIssue.comments ?? [],
                        visited
                    );

                    if (linkedResolution.trim()) {
                        results.push(
                            `Linked Ticket #${linkedId}:\n${linkedResolution}`
                        );
                    }
                } catch (error) {
                    console.error(
                        `Unable to fetch linked ticket #${linkedId}:`,
                        error.response?.status ?? error.message
                    );

                    results.push(
                        `Linked Ticket #${linkedId}: Unable to retrieve resolution.`
                    );
                }
            }

            // Do not include the reference-only comment as a resolution.
            continue;
        }

        const cleaned = cleanComment(text);
        const sections = extractTaggedSections(cleaned);

        if (sections) {
            results.push(sections);
        } else if (cleaned) {
            results.push(cleaned);
        }
    }

    return results.join("\n\n");
}


function extractTaggedSections(text = "") {
    const regex =
        /(Root Cause|Solution|Resolution)\s*:?\s*([\s\S]*?)(?=(?:Root Cause|Solution|Resolution)\s*:|$)/gi;

    const sections = [];
    let match;

    while ((match = regex.exec(text)) !== null) {
        const title = match[1]
            .trim()
            .toLowerCase()
            .replace(/\b\w/g, char => char.toUpperCase());

        const content = match[2].trim();

        if (content) {
            sections.push(`${title}:\n${content}`);
        }
    }

    return sections.length ? sections.join("\n\n") : null;
}


function extractMergeRequests(text = "") {
    const regex =
        /https?:\/\/git\.centralizedinc\.com\/([^\s"'<>()[\]]+?)\/-\/merge_requests\/(\d+)/gi;

    const mergeRequests = [];
    let match;

    while ((match = regex.exec(text)) !== null) {
        const projectPath = match[1].replace(/\/+$/, "");
        const pathParts = projectPath.split("/");
        const application = pathParts
            .map(part => part.toLowerCase())
            .find(part => ["cms", "hris", "lms"].includes(part));

        if (!application) continue;

        mergeRequests.push({
            projectPath,
            application: application.toUpperCase(),
            image: pathParts[pathParts.length - 1],
            mrNumber: match[2],
            url: match[0].replace(/[.,;]+$/, "")
        });
    }

    return mergeRequests;
}


function collectTriggeredProjects(issues) {
    const projects = new Map();

    for (const issue of issues) {
        for (const comment of issue.comments ?? []) {
            const mergeRequests = extractMergeRequests(
                comment.comment ?? ""
            );

            for (const mr of mergeRequests) {
                // One row per unique full repository path.
                if (!projects.has(mr.projectPath)) {
                    projects.set(mr.projectPath, {
                        application: mr.application,
                        image: mr.image,
                        projectPath: mr.projectPath,
                        mergeRequests: new Set()
                    });
                }

                projects.get(mr.projectPath).mergeRequests.add(mr.url);
            }
        }
    }

    return [...projects.values()].sort((a, b) =>
        a.application.localeCompare(b.application) ||
        a.image.localeCompare(b.image)
    );
}



function cleanComment(text = "") {
    return text
        .replace(/<[^>]*>/g, " ")
        .replace(/https?:\/\/\S+/gi, "")
        .replace(/MR\s*:?\s*https?:\/\/\S+/gi, "")
        .replace(/Merge Request\s*:?.*/gi, "")
        .replace(/!?\d+\s*\(?\s*(merge request|mr)\s*\)?/gi, "")
        .replace(/Reference Ticket\s*:?.*/gi, "")
        .replace(/\n\s*\n/g, "\n")
        .replace(/[ \t]+/g, " ")
        .trim();
}


async function generateExcel(app, deployTag, issues) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Deployment");

    // Deployment issues columns
    sheet.columns = [
        { header: "Issue No.", key: "issueNo", width: 22 },
        { header: "Title", key: "title", width: 40 },
        { header: "Resolution", key: "resolution", width: 80 }
    ];

    // Add deployment issues
    for (const issue of issues) {
        sheet.addRow({
            issueNo: issue.code,
            title: issue.name,
            resolution: await extractResolution(
                issue.comments ?? [],
                new Set([String(issue.id)])
            )
        });
    }

    // Collect unique repositories from MR links
    const triggeredProjects = collectTriggeredProjects(issues);

    // Add triggered projects summary
    sheet.addRow([]);
    const titleRow = sheet.addRow(["TRIGGERED PROJECTS"]);
    titleRow.font = { bold: true, size: 12 };

    const headerRow = sheet.addRow([
        "APPLICATION",
        "IMAGE"
    ]);

    headerRow.font = { bold: true };

    for (const project of triggeredProjects) {
        const row = sheet.addRow([
            project.application,
            project.image
        ]);

        // Store project path and unique MR links in the IMAGE cell note
        row.getCell(2).note = [
            `Project Path: ${project.projectPath}`,
            `Unique MR Count: ${project.mergeRequests.size}`,
            ...project.mergeRequests
        ].join("\n");
    }

    // Format triggered projects table
    for (
        let rowIndex = headerRow.number;
        rowIndex <= sheet.rowCount;
        rowIndex++
    ) {
        const row = sheet.getRow(rowIndex);

        for (let col = 1; col <= 2; col++) {
            const cell = row.getCell(col);

            cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" }
            };

            cell.alignment = {
                horizontal: "center",
                vertical: "middle",
                wrapText: true
            };
        }
    }

    const tagLabel = deployTag ?? "ALL";
    const filename =
        `${app.name}_${tagLabel}_Deployment_Report.xlsx`;
    const buffer = await workbook.xlsx.writeBuffer();

    return { buffer, filename, issueCount: issues.length };
}

async function generateDeploymentReport(appName, deployTag) {
    if (!process.env.TOKEN) {
        throw new Error("Missing TOKEN in .env");
    }

    const app = APPLICATIONS[appName];
    if (!app) throw new Error("Unsupported application.");
    if (!["PROD", "STAGING", null].includes(deployTag)) {
        throw new Error("Unsupported deployment tag.");
    }

    const issues = await getIssues({ name: appName, ...app }, deployTag);
    return generateExcel({ name: appName, ...app }, deployTag, issues);
}


async function main() {
    if (!process.env.TOKEN) {
        throw new Error("Missing TOKEN in .env");
    }

    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    try {
        const app = await selectApplication(rl);
        const deployTag = await selectDeployTag(rl);

        console.log("\nFetching deployment issues...");

        const { buffer, filename, issueCount } = await generateDeploymentReport(app.name, deployTag);
        const outputDir = path.join(__dirname, "output");
        require("fs").mkdirSync(outputDir, { recursive: true });
        const outputPath = path.join(outputDir, filename);
        require("fs").writeFileSync(outputPath, buffer);

        console.log("\n========== REPORT GENERATED ==========");
        console.log(`Application : ${app.name}`);
        console.log(`Client ID   : ${app.clientId}`);
        console.log(`Deploy Tag  : ${deployTag ?? "ALL"}`);
        console.log(`Total Issues: ${issueCount}`);
        console.log(`Output File : ${outputPath}`);
    } catch (error) {
        console.error(
            "\nFailed to generate report:",
            error.response?.data ?? error.message
        );
        process.exitCode = 1;
    } finally {
        rl.close();
    }
}

if (require.main === module) main();

module.exports = { generateDeploymentReport };
