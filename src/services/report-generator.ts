import { getGraphClient } from "../lib/graph.js";
import { getClassificationLog, getClassificationStats } from "../tools/email.js";

export interface ReportRequest {

    reportType: "monthly" | "weekly" | "on_demand";

    dataSources: ("planner" | "email_metrics" | "both")[];

    output: ("teams" | "email" | "both")[];

    emailTo?: string;

}

interface BucketSummary {
    vip: number;
    urgent: number;
    client: number;
    operations: number;
    support: number;
    backlog: number;
}

interface TaskSummary {

    total: number;

    completed: number;

    completionRate: string;

    byBucket: BucketSummary;

    overdue: number;

}

interface TopClient {
    domain: string;
    count: number;
}

interface EmailMetricsSummary {

    totalClassified: number;

    byCategory: Record<string, number>;

    byBucket: Record<string, number>;

    topClients: TopClient[];

    potentialFalsePositives: number;

}

interface ReportData {
    generatedAt: string;
    reportType: string;
    taskSummary?: TaskSummary;
    emailMetrics?: EmailMetricsSummary;
}

export interface ReportResult {

    success: boolean;

    report: ReportData;

    deliveredTo: string[];

    errors: string[];

}

interface BucketIds {
    vip: string;
    urgent: string;
    client: string;
    operations: string;
    support: string;
    backlog: string;
}

interface PlannerTask {
    percentComplete: number;
    dueDateTime?: string;
    bucketId?: string;
}

interface PlannerTasksResponse {
    value?: PlannerTask[];
}

interface ErrorLike {
    message: string;
}

// IDs from environment

const GROUP_ID: string = process.env.PLANNER_GROUP_ID || process.env.TEAMS_GROUP_ID || "8c1829c3-aeb7-4af6-905d-a81023b3bebd";

const PLAN_ID: string = process.env.PLANNER_PLAN_ID || "isN5lApe9UKwJzUcBXpiWsgABOwz";

const TEAMS_CHANNEL_ID: string = process.env.TEAMS_CHANNEL_ID || "19:dCAbGRmp2ZNxjwjPCVC7iaKRp6XbJWz6fvO9uBEgalM1@thread.tacv2";

const DEFAULT_EMAIL: string = process.env.DEFAULT_EMAIL || "Danielb@naviafreight.com";

// Bucket IDs

const BUCKETS: BucketIds = {

    vip: process.env.BUCKET_VIP || "MR_zXOcdqE6RIF6Yd6wPjMgAEaxq",

    urgent: process.env.BUCKET_URGENT || "r1uZj-Zh7E63PTEPYqPmIsgANefB",

    client: process.env.BUCKET_CLIENT || "RGDYN7r6oEWXkIA0L01o5sgAKG4S",

    operations: process.env.BUCKET_OPERATIONS || "3hPTGvtl7kC5QGaKItSyDMgAD15d",

    support: process.env.BUCKET_SUPPORT || "UBHJXpRD50qAPuqJODCDm8gAFBEB",

    backlog: process.env.BUCKET_BACKLOG || "-ve7obEGl0GRm3nQC7ebN8gAPac6",

};

// Get Planner tasks and calculate summary

async function getPlannerSummary(): Promise<TaskSummary> {

    try {

        const client = await getGraphClient();

        // Get all tasks from the plan

        const response = await client

            .api(`/planner/plans/${PLAN_ID}/tasks`)

            .get() as PlannerTasksResponse;

        const tasks: PlannerTask[] = response.value || [];

        const total: number = tasks.length;

        const completed: number = tasks.filter((t: PlannerTask) => t.percentComplete === 100).length;

        const now: Date = new Date();

        const overdue: number = tasks.filter((t: PlannerTask) => t.dueDateTime &&

            new Date(t.dueDateTime) < now &&

            t.percentComplete !== 100).length;

        // Count by bucket

        const byBucket: BucketSummary = {

            vip: tasks.filter((t: PlannerTask) => t.bucketId === BUCKETS.vip).length,

            urgent: tasks.filter((t: PlannerTask) => t.bucketId === BUCKETS.urgent).length,

            client: tasks.filter((t: PlannerTask) => t.bucketId === BUCKETS.client).length,

            operations: tasks.filter((t: PlannerTask) => t.bucketId === BUCKETS.operations).length,

            support: tasks.filter((t: PlannerTask) => t.bucketId === BUCKETS.support).length,

            backlog: tasks.filter((t: PlannerTask) => t.bucketId === BUCKETS.backlog).length,

        };

        return {

            total,

            completed,

            completionRate: total > 0 ? `${((completed / total) * 100).toFixed(1)}%` : "N/A",

            byBucket,

            overdue,

        };

    }

    catch (error: unknown) {

        const err = error as ErrorLike;

        console.error("Error getting Planner summary:", err.message);

        throw error;

    }

}

// Get email classification metrics

function getEmailMetricsSummary(): EmailMetricsSummary {

    const stats = getClassificationStats();

    const log = getClassificationLog(500);

    // Get top clients

    const clientCounts: Record<string, number> = {};

    for (const entry of log) {

        if (entry.domain) {

            clientCounts[entry.domain] = (clientCounts[entry.domain] || 0) + 1;

        }

    }

    const topClients: TopClient[] = Object.entries(clientCounts)

        .sort((a, b) => b[1] - a[1])

        .slice(0, 5)

        .map(([domain, count]) => ({ domain, count }));

    return {

        totalClassified: stats.total,

        byCategory: stats.byCategory,

        byBucket: stats.byBucket,

        topClients,

        potentialFalsePositives: stats.potentialFalsePositives.length,

    };

}

// Post message to Teams channel

async function postToTeams(message: string): Promise<void> {

    try {

        const client = await getGraphClient();

        await client

            .api(`/teams/${GROUP_ID}/channels/${TEAMS_CHANNEL_ID}/messages`)

            .post({

            body: {

                contentType: "html",

                content: message,

            },

        });

        console.log("[REPORT] Posted to Teams channel");

    }

    catch (error: unknown) {

        const err = error as ErrorLike;

        console.error("Error posting to Teams:", err.message);

        throw error;

    }

}

// Send email via Graph API

async function sendEmail(to: string, subject: string, body: string): Promise<void> {

    try {

        const client = await getGraphClient();

        // Using sendMail which requires Mail.Send permission

        // For app-only, we need to send as a specific user or use a shared mailbox

        await client

            .api("/users/Danielb@naviafreight.com/sendMail")

            .post({

            message: {

                subject,

                body: {

                    contentType: "HTML",

                    content: body,

                },

                toRecipients: [

                    {

                        emailAddress: {

                            address: to,

                        },

                    },

                ],

            },

        });

        console.log(`[REPORT] Email sent to ${to}`);

    }

    catch (error: unknown) {

        const err = error as ErrorLike;

        console.error("Error sending email:", err.message);

        throw error;

    }

}

// Build HTML report

function buildHtmlReport(reportType: ReportRequest["reportType"], taskSummary?: TaskSummary, emailMetrics?: EmailMetricsSummary): string {

    const now: string = new Date().toLocaleString("en-AU", { timeZone: "Australia/Sydney" });

    let html = `

<h2>Operations Report - ${reportType}</h2>

<p><em>Generated: ${now}</em></p>

<hr/>

`;

    if (taskSummary) {

        html += `

<h3>Planner Task Summary</h3>

<table border="1" cellpadding="8" style="border-collapse: collapse;">

  <tr style="background-color: #f0f0f0;">

    <th>Metric</th>

    <th>Value</th>

  </tr>

  <tr>

    <td>Total Tasks</td>

    <td><strong>${taskSummary.total}</strong></td>

  </tr>

  <tr>

    <td>Completed</td>

    <td>${taskSummary.completed}</td>

  </tr>

  <tr>

    <td>Completion Rate</td>

    <td><strong>${taskSummary.completionRate}</strong></td>

  </tr>

  <tr>

    <td>Overdue</td>

    <td style="color: ${taskSummary.overdue > 0 ? "red" : "green"};">${taskSummary.overdue}</td>

  </tr>

</table>



<h4>Tasks by Bucket</h4>

<table border="1" cellpadding="8" style="border-collapse: collapse;">

  <tr style="background-color: #f0f0f0;">

    <th>Bucket</th>

    <th>Count</th>

  </tr>

  <tr><td>VIP Follow-up</td><td>${taskSummary.byBucket.vip}</td></tr>

  <tr><td>Urgent</td><td>${taskSummary.byBucket.urgent}</td></tr>

  <tr><td>Client Communications</td><td>${taskSummary.byBucket.client}</td></tr>

  <tr><td>Operations</td><td>${taskSummary.byBucket.operations}</td></tr>

  <tr><td>Support</td><td>${taskSummary.byBucket.support}</td></tr>

  <tr><td>Backlog</td><td>${taskSummary.byBucket.backlog}</td></tr>

</table>

`;

    }

    if (emailMetrics) {

        html += `

<h3>Email Classification Metrics</h3>

<p>Total Emails Classified: <strong>${emailMetrics.totalClassified}</strong></p>



<h4>By Category</h4>

<ul>

${Object.entries(emailMetrics.byCategory).map(([cat, count]) => `<li>${cat}: ${count}</li>`).join("\n")}

</ul>



<h4>Top 5 Clients by Volume</h4>

<ol>

${emailMetrics.topClients.map((c: TopClient) => `<li>${c.domain}: ${c.count} emails</li>`).join("\n")}

</ol>



<p>Potential False Positives: <strong>${emailMetrics.potentialFalsePositives}</strong></p>

`;

    }

    html += `

<hr/>

<p style="color: #666; font-size: 12px;">Generated by Navia MCP Server</p>

`;

    return html;

}

// Main report generation function

export async function generateReport(request: ReportRequest): Promise<ReportResult> {

    const result: ReportResult = {

        success: false,

        report: {

            generatedAt: new Date().toISOString(),

            reportType: request.reportType,

        },

        deliveredTo: [],

        errors: [],

    };

    try {

        // Determine what data to collect

        const includePlanner: boolean = request.dataSources.includes("planner") ||

            request.dataSources.includes("both");

        const includeEmail: boolean = request.dataSources.includes("email_metrics") ||

            request.dataSources.includes("both");

        // Collect data

        if (includePlanner) {

            try {

                result.report.taskSummary = await getPlannerSummary();

                console.log("[REPORT] Planner data collected");

            }

            catch (e: unknown) {

                const err = e as ErrorLike;

                result.errors.push(`Planner error: ${err.message}`);

            }

        }

        if (includeEmail) {

            try {

                result.report.emailMetrics = getEmailMetricsSummary();

                console.log("[REPORT] Email metrics collected");

            }

            catch (e: unknown) {

                const err = e as ErrorLike;

                result.errors.push(`Email metrics error: ${err.message}`);

            }

        }

        // Build report HTML

        const reportHtml: string = buildHtmlReport(request.reportType, result.report.taskSummary, result.report.emailMetrics);

        // Determine output destinations

        const sendToTeams: boolean = request.output.includes("teams") ||

            request.output.includes("both");

        const sendToEmail: boolean = request.output.includes("email") ||

            request.output.includes("both");

        // Deliver report

        if (sendToTeams) {

            try {

                await postToTeams(reportHtml);

                result.deliveredTo.push("teams");

            }

            catch (e: unknown) {

                const err = e as ErrorLike;

                result.errors.push(`Teams error: ${err.message}`);

            }

        }

        if (sendToEmail) {

            try {

                const emailTo: string = request.emailTo || DEFAULT_EMAIL;

                await sendEmail(emailTo, `Operations Report - ${request.reportType} - ${new Date().toLocaleDateString()}`, reportHtml);

                result.deliveredTo.push(`email:${emailTo}`);

            }

            catch (e: unknown) {

                const err = e as ErrorLike;

                result.errors.push(`Email error: ${err.message}`);

            }

        }

        result.success = result.errors.length === 0;

    }

    catch (error: unknown) {

        const err = error as ErrorLike;

        result.errors.push(`General error: ${err.message}`);

    }

    return result;

}
