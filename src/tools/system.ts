import type { Tool } from "@modelcontextprotocol/sdk/types.js";

type ToolResult = Promise<{ content: { type: string; text: string }[] }>;

// In-memory error log (last 100 entries)
const MAX_ERROR_LOG = 100;
const errorLog: Array<{
    timestamp: string;
    tool: string;
    error: string;
    context?: string;
}> = [];

export const systemTools: Tool[] = [
    {
        name: "system_health_check",
        description: "Check the health status of the MCP server, including uptime, registered tool count, memory usage, and connectivity to external services (Graph API, Dataverse, CargoWise)",
        inputSchema: {
            type: "object",
            properties: {
                includeDetails: {
                    type: "boolean",
                    description: "Include detailed diagnostics (memory, env vars status). Default: false",
                },
            },
        },
    },
    {
        name: "system_log_error",
        description: "Log an error encountered during tool execution for diagnostics and self-healing. Returns the error log history.",
        inputSchema: {
            type: "object",
            properties: {
                tool: {
                    type: "string",
                    description: "Name of the tool that encountered the error",
                },
                error: {
                    type: "string",
                    description: "Error message or description",
                },
                context: {
                    type: "string",
                    description: "Additional context about what was being attempted",
                },
            },
            required: ["tool", "error"],
        },
    },
    {
        name: "system_get_config",
        description: "Get the current server configuration including enabled modules, environment variable status (names only, not values), and feature flags",
        inputSchema: {
            type: "object",
            properties: {},
        },
    },
];

const startTime = Date.now();

export async function handleSystemTool(name: string, args: Record<string, unknown>): ToolResult {
    switch (name) {
        case "system_health_check": {
            const includeDetails = (args.includeDetails as boolean) || false;
            const uptimeMs = Date.now() - startTime;
            const uptimeHours = Math.floor(uptimeMs / 3600000);
            const uptimeMinutes = Math.floor((uptimeMs % 3600000) / 60000);

            const health: Record<string, unknown> = {
                status: "healthy",
                uptime: `${uptimeHours}h ${uptimeMinutes}m`,
                uptimeMs,
                timestamp: new Date().toISOString(),
                services: {
                    graphApi: {
                        configured: !!(process.env.AZURE_TENANT_ID && process.env.AZURE_CLIENT_ID && process.env.AZURE_CLIENT_SECRET),
                        status: (process.env.AZURE_TENANT_ID && process.env.AZURE_CLIENT_ID && process.env.AZURE_CLIENT_SECRET) ? "ready" : "not_configured",
                    },
                    dataverse: {
                        configured: !!process.env.DATAVERSE_URL,
                        url: process.env.DATAVERSE_URL ? "configured" : "not_set",
                        status: process.env.DATAVERSE_URL ? "ready" : "not_configured",
                    },
                    cargowise: {
                        configured: !!(process.env.CARGOWISE_ENDPOINT && process.env.CARGOWISE_USERNAME && process.env.CARGOWISE_PASSWORD),
                        status: (process.env.CARGOWISE_ENDPOINT && process.env.CARGOWISE_USERNAME && process.env.CARGOWISE_PASSWORD) ? "ready" : "not_configured",
                    },
                    memory: {
                        configured: !!(process.env.AZURE_STORAGE_CONNECTION_STRING || process.env.AZURE_STORAGE_ACCOUNT_NAME),
                        status: (process.env.AZURE_STORAGE_CONNECTION_STRING || process.env.AZURE_STORAGE_ACCOUNT_NAME) ? "ready" : "not_configured",
                    },
                },
                recentErrors: errorLog.length,
            };

            if (includeDetails) {
                const memUsage = process.memoryUsage();
                health.details = {
                    memory: {
                        heapUsedMB: Math.round(memUsage.heapUsed / 1048576),
                        heapTotalMB: Math.round(memUsage.heapTotal / 1048576),
                        rssMB: Math.round(memUsage.rss / 1048576),
                    },
                    nodeVersion: process.version,
                    platform: process.platform,
                    pid: process.pid,
                    toolExposure: process.env.MCP_TOOL_EXPOSURE || "FULL",
                    recentErrors: errorLog.slice(-5),
                };
            }

            return {
                content: [{ type: "text", text: JSON.stringify(health, null, 2) }],
            };
        }

        case "system_log_error": {
            const tool = args.tool as string;
            const error = args.error as string;
            const context = args.context as string | undefined;

            const entry = {
                timestamp: new Date().toISOString(),
                tool,
                error,
                context,
            };

            errorLog.push(entry);
            if (errorLog.length > MAX_ERROR_LOG) {
                errorLog.shift();
            }

            console.error(`[SYSTEM_ERROR] tool=${tool} error=${error}${context ? ` context=${context}` : ""}`);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            logged: true,
                            entry,
                            totalErrors: errorLog.length,
                            recentErrors: errorLog.slice(-5),
                        }, null, 2),
                    },
                ],
            };
        }

        case "system_get_config": {
            const envStatus = (key: string): string => process.env[key] ? "set" : "not_set";

            const config = {
                server: {
                    name: "navia-mcp-server",
                    version: "2.1.0",
                    port: process.env.PORT || 88,
                    toolExposure: process.env.MCP_TOOL_EXPOSURE || "FULL",
                },
                modules: {
                    email: "enabled",
                    actionPlanner: "enabled",
                    teams: "enabled",
                    outlook: "enabled",
                    dataverse: "enabled",
                    planner: "enabled",
                    memory: "enabled",
                    cargowise: "enabled",
                    system: "enabled",
                },
                environment: {
                    AZURE_TENANT_ID: envStatus("AZURE_TENANT_ID"),
                    AZURE_CLIENT_ID: envStatus("AZURE_CLIENT_ID"),
                    AZURE_CLIENT_SECRET: envStatus("AZURE_CLIENT_SECRET"),
                    DATAVERSE_URL: envStatus("DATAVERSE_URL"),
                    OUTLOOK_USER_ID: envStatus("OUTLOOK_USER_ID"),
                    CARGOWISE_ENDPOINT: envStatus("CARGOWISE_ENDPOINT"),
                    CARGOWISE_USERNAME: envStatus("CARGOWISE_USERNAME"),
                    CARGOWISE_PASSWORD: envStatus("CARGOWISE_PASSWORD"),
                    AZURE_STORAGE_CONNECTION_STRING: envStatus("AZURE_STORAGE_CONNECTION_STRING"),
                    AZURE_STORAGE_ACCOUNT_NAME: envStatus("AZURE_STORAGE_ACCOUNT_NAME"),
                    TEAMS_WEBHOOK_URL: envStatus("TEAMS_WEBHOOK_URL"),
                },
            };

            return {
                content: [{ type: "text", text: JSON.stringify(config, null, 2) }],
            };
        }

        default:
            throw new Error(`Unknown system tool: ${name}`);
    }
}
