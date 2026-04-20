import { ClientSecretCredential } from "@azure/identity";
import type { Tool } from "@modelcontextprotocol/sdk/types.js";

// Dataverse client setup
let dataverseToken: string | null = null;
let tokenExpiry: Date | null = null;

async function getDataverseToken(): Promise<string> {
    // Check if we have a valid token
    if (dataverseToken && tokenExpiry && tokenExpiry > new Date()) {
        return dataverseToken;
    }

    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const dataverseUrl = process.env.DATAVERSE_URL || "https://dcbnavia.crm6.dynamics.com";

    if (!tenantId || !clientId || !clientSecret) {
        throw new Error("Missing Azure credentials for Dataverse");
    }

    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    const token = (await credential.getToken(`${dataverseUrl}/.default`)) as {
        token: string;
        expiresOnTimestamp: number;
    };

    dataverseToken = token.token;

    // Set expiry 5 minutes before actual expiry for safety
    tokenExpiry = new Date(token.expiresOnTimestamp - 5 * 60 * 1000);

    return dataverseToken;
}

async function dataverseRequest(
    method: "GET" | "POST" | "PATCH" | "DELETE",
    endpoint: string,
    body?: Record<string, unknown>,
): Promise<Record<string, unknown>> {
    const token = await getDataverseToken();
    const dataverseUrl = process.env.DATAVERSE_URL || "https://dcbnavia.crm6.dynamics.com";
    const url = `${dataverseUrl}/api/data/v9.2${endpoint}`;

    const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        Accept: "application/json",
        Prefer: "return=representation",
    };

    if (body) {
        headers["Content-Type"] = "application/json";
    }

    const response = await fetch(url, {
        method,
        headers,
        body: body ? JSON.stringify(body) : undefined,
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Dataverse API error: ${response.status} - ${errorText}`);
    }

    // DELETE returns no content
    if (response.status === 204) {
        return { success: true };
    }

    return (await response.json()) as Record<string, unknown>;
}

export const dataverseTools: Tool[] = [
    {
        name: "dataverse_list_records",
        description: "List records from a Dataverse table with optional filtering",
        inputSchema: {
            type: "object",
            properties: {
                tableName: {
                    type: "string",
                    description: "The logical name of the table (e.g., 'accounts', 'contacts', 'nf_vipclient1s')",
                },
                select: {
                    type: "string",
                    description: "Comma-separated list of columns to return (e.g., 'name,emailaddress1')",
                },
                filter: {
                    type: "string",
                    description: "OData filter expression (e.g., \"statecode eq 0\")",
                },
                top: {
                    type: "number",
                    description: "Maximum records to return. Default: 50, Max: 5000",
                },
                orderby: {
                    type: "string",
                    description: "Sort order (e.g., 'createdon desc')",
                },
            },
            required: ["tableName"],
        },
    },
    {
        name: "dataverse_get_record",
        description: "Get a single record by ID from a Dataverse table",
        inputSchema: {
            type: "object",
            properties: {
                tableName: {
                    type: "string",
                    description: "The logical name of the table",
                },
                recordId: {
                    type: "string",
                    description: "The GUID of the record",
                },
                select: {
                    type: "string",
                    description: "Comma-separated list of columns to return",
                },
            },
            required: ["tableName", "recordId"],
        },
    },
    {
        name: "dataverse_create_record",
        description: "Create a new record in a Dataverse table",
        inputSchema: {
            type: "object",
            properties: {
                tableName: {
                    type: "string",
                    description: "The logical name of the table",
                },
                data: {
                    type: "object",
                    description: "The record data as key-value pairs",
                },
            },
            required: ["tableName", "data"],
        },
    },
    {
        name: "dataverse_update_record",
        description: "Update an existing record in a Dataverse table",
        inputSchema: {
            type: "object",
            properties: {
                tableName: {
                    type: "string",
                    description: "The logical name of the table",
                },
                recordId: {
                    type: "string",
                    description: "The GUID of the record to update",
                },
                data: {
                    type: "object",
                    description: "The fields to update as key-value pairs",
                },
            },
            required: ["tableName", "recordId", "data"],
        },
    },
    {
        name: "dataverse_delete_record",
        description: "Delete a record from a Dataverse table",
        inputSchema: {
            type: "object",
            properties: {
                tableName: {
                    type: "string",
                    description: "The logical name of the table",
                },
                recordId: {
                    type: "string",
                    description: "The GUID of the record to delete",
                },
            },
            required: ["tableName", "recordId"],
        },
    },
    {
        name: "dataverse_query",
        description: "Execute a custom OData query against Dataverse",
        inputSchema: {
            type: "object",
            properties: {
                query: {
                    type: "string",
                    description: "The full OData query path (e.g., '/accounts?$select=name&$top=10')",
                },
            },
            required: ["query"],
        },
    },
    {
        name: "dataverse_query_with_expand",
        description: "Query Dataverse records with $expand to include related/linked records. Use this to retrieve parent-child or lookup relationships in a single call.",
        inputSchema: {
            type: "object",
            properties: {
                tableName: {
                    type: "string",
                    description: "The logical name of the table (e.g., 'accounts', 'contacts', 'nf_emaillogs')",
                },
                select: {
                    type: "string",
                    description: "Comma-separated list of columns to return from the main table",
                },
                expand: {
                    type: "string",
                    description: "OData $expand expression for related records (e.g., 'primarycontactid($select=fullname,emailaddress1)' or 'nf_RelatedEmails($select=nf_subject,nf_sender;$top=5)')",
                },
                filter: {
                    type: "string",
                    description: "OData filter expression",
                },
                top: {
                    type: "number",
                    description: "Maximum records to return. Default: 50, Max: 5000",
                },
                orderby: {
                    type: "string",
                    description: "Sort order (e.g., 'createdon desc')",
                },
            },
            required: ["tableName", "expand"],
        },
    },
    {
        name: "dataverse_create_3pl_integration",
        description: "Create a new record in the Customer 3PL API Integrations table",
        inputSchema: {
            type: "object",
            properties: {
                name: {
                    type: "string",
                    description: "Integration name/title",
                },
                customer: {
                    type: "string",
                    description: "Customer account ID (lookup)",
                },
                apiType: {
                    type: "string",
                    enum: ["Shopify", "WooCommerce", "Cin7", "Webhook", "Custom"],
                    description: "Type of API integration",
                },
                status: {
                    type: "string",
                    enum: ["Active", "Inactive", "Testing", "Error"],
                    description: "Integration status",
                },
                configuration: {
                    type: "string",
                    description: "JSON configuration for the integration",
                },
                notes: {
                    type: "string",
                    description: "Additional notes about the integration",
                },
            },
            required: ["name"],
        },
    },
];

// Table name to logical name mapping for common custom tables
const TABLE_ALIASES: Record<string, string> = {
    // Pre-existing nf_ tables
    "vip_clients": "nf_vipclient1s",
    "email_triage_logs": "nf_emailtriagelog1s",
    "3pl_integrations": "nf_customer3plapiintegrations",
    "integration_errors": "nf_integrationerror1s",
    "agent_config": "nf_agentconfigurations",
    "agent_memory": "nf_agentmemorys",
    "email_triage_events": "nf_emailtriageevents",
    "email_rules": "nf_emailrules",
    "vip_tracking": "nf_vipemailtrackings",
    "resources": "nf_resources",
    // Vibe-created nf_ tables (primary source of truth)
    "triage_logs": "nf_triagelogs",
    "email_logs": "nf_emaillogs",
    "routing_rules": "nf_routingrules",
    "extracted_entities": "nf_extractedentitys",
    "audit_logs": "nf_auditlogs",
    "planner_tasks": "nf_plannertasks",
    "priority_levels": "nf_prioritylevels",
    "keyword_rules": "nf_keywordrules",
    "vip_domains": "nf_vipdomains",
    "folder_mappings": "nf_foldermappings",
    "channel_mappings": "nf_teamschannelmappings",
    "emails": "nf_emails",
    // New tables
    "allowed_targets": "nf_allowedtargets",
    "group_plans": "nf_groupplans",
    // Vibe cr4b9_ tables (kept where no nf_ equivalent)
    "p1_escalations": "cr4b9_p1escalations",
    "insight_summaries": "cr4b9_insightsummaries",
    "task_links": "cr4b9_tasklinks",
    // AI training data
    "training_data": "cr4b9_aitrainingdatas",
};

function resolveTableName(tableName: string): string {
    return TABLE_ALIASES[tableName.toLowerCase()] || tableName;
}

export async function handleDataverseTool(
    name: string,
    args: Record<string, unknown>,
): Promise<{ content: { type: string; text: string }[] }> {
    switch (name) {
        case "dataverse_list_records": {
            const tableNameArg = args.tableName as string;
            const topArg = args.top as number | undefined;
            const selectArg = args.select as string | undefined;
            const filterArg = args.filter as string | undefined;
            const orderbyArg = args.orderby as string | undefined;

            const tableName = resolveTableName(tableNameArg);
            const top = Math.min(topArg || 50, 5000);
            let endpoint = `/${tableName}?$top=${top}`;

            if (selectArg) {
                endpoint += `&$select=${selectArg}`;
            }

            if (filterArg) {
                endpoint += `&$filter=${encodeURIComponent(filterArg)}`;
            }

            if (orderbyArg) {
                endpoint += `&$orderby=${encodeURIComponent(orderbyArg)}`;
            }

            const result = await dataverseRequest("GET", endpoint);
            const value = (result.value as unknown[] | undefined) || [];

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                count: value.length || 0,
                                records: value || [],
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "dataverse_get_record": {
            const tableNameArg = args.tableName as string;
            const recordIdArg = args.recordId as string;
            const selectArg = args.select as string | undefined;

            const tableName = resolveTableName(tableNameArg);
            const recordId = recordIdArg;
            let endpoint = `/${tableName}(${recordId})`;

            if (selectArg) {
                endpoint += `?$select=${selectArg}`;
            }

            const result = await dataverseRequest("GET", endpoint);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(result, null, 2),
                    },
                ],
            };
        }

        case "dataverse_create_record": {
            const tableNameArg = args.tableName as string;
            const dataArg = args.data as Record<string, unknown>;

            const tableName = resolveTableName(tableNameArg);
            const data = dataArg;

            const result = await dataverseRequest("POST", `/${tableName}`, data);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                created: result,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "dataverse_update_record": {
            const tableNameArg = args.tableName as string;
            const recordIdArg = args.recordId as string;
            const dataArg = args.data as Record<string, unknown>;

            const tableName = resolveTableName(tableNameArg);
            const recordId = recordIdArg;
            const data = dataArg;

            const result = await dataverseRequest("PATCH", `/${tableName}(${recordId})`, data);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                updated: result,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "dataverse_delete_record": {
            const tableNameArg = args.tableName as string;
            const recordIdArg = args.recordId as string;

            const tableName = resolveTableName(tableNameArg);
            const recordId = recordIdArg;

            await dataverseRequest("DELETE", `/${tableName}(${recordId})`);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                deleted: recordId,
                                table: tableName,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "dataverse_query_with_expand": {
            const tableNameArg = args.tableName as string;
            const selectArg = args.select as string | undefined;
            const expandArg = args.expand as string;
            const filterArg = args.filter as string | undefined;
            const topArg = args.top as number | undefined;
            const orderbyArg = args.orderby as string | undefined;

            const tableName = resolveTableName(tableNameArg);
            const top = Math.min(topArg || 50, 5000);
            let endpoint = `/${tableName}?$top=${top}&$expand=${encodeURIComponent(expandArg)}`;

            if (selectArg) {
                endpoint += `&$select=${selectArg}`;
            }

            if (filterArg) {
                endpoint += `&$filter=${encodeURIComponent(filterArg)}`;
            }

            if (orderbyArg) {
                endpoint += `&$orderby=${encodeURIComponent(orderbyArg)}`;
            }

            const result = await dataverseRequest("GET", endpoint);
            const value = (result.value as unknown[] | undefined) || [];

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                count: value.length || 0,
                                records: value || [],
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "dataverse_query": {
            const queryArg = args.query as string;
            const query = queryArg;

            const result = await dataverseRequest("GET", query);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(result, null, 2),
                    },
                ],
            };
        }

        case "dataverse_create_3pl_integration": {
            const nameArg = args.name as string;
            const customerArg = args.customer as string | undefined;
            const apiTypeArg = args.apiType as string | undefined;
            const statusArg = args.status as string | undefined;
            const configurationArg = args.configuration as string | undefined;
            const notesArg = args.notes as string | undefined;

            // Map friendly field names to actual column names
            const data: Record<string, string | number> = {
                nf_name: nameArg,
            };

            if (customerArg) {
                data["nf_Customer@odata.bind"] = `/accounts(${customerArg})`;
            }

            if (apiTypeArg) {
                // Map to option set value if needed
                const apiTypeMap: Record<string, number> = {
                    Shopify: 1,
                    WooCommerce: 2,
                    Cin7: 3,
                    Webhook: 4,
                    Custom: 5,
                };

                data.nf_apitype = apiTypeMap[apiTypeArg] || 5;
            }

            if (statusArg) {
                const statusMap: Record<string, number> = {
                    Active: 1,
                    Inactive: 2,
                    Testing: 3,
                    Error: 4,
                };

                data.nf_status = statusMap[statusArg] || 1;
            }

            if (configurationArg) {
                data.nf_configuration = configurationArg;
            }

            if (notesArg) {
                data.nf_notes = notesArg;
            }

            const result = await dataverseRequest("POST", "/nf_customer3plapiintegrations", data);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                created: result,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        default:
            throw new Error(`Unknown dataverse tool: ${name}`);
    }
}
