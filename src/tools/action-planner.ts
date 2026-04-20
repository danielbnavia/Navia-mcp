/**
 * Action Planner Tool
 *
 * Generates action plans for emails based on routing rules.
 * Returns structured actions for Teams posts, DMs, and Planner tasks.
 * Also handles Outlook email rules for automatic filing and forwarding.
 */
import type { Tool } from "@modelcontextprotocol/sdk/types.js";
import {
    DM_RECIPIENTS,
    extractEntities,
    generateActionPlan,
    PLANNER_PLANS,
    REGEX_PATTERNS,
    TEAMS_CHANNELS,
    VIP_LIST,
} from "../config/routing.js";
import {
    FORWARD_RECIPIENTS,
    evaluateRules,
    getRulesSummary,
    OUTLOOK_FOLDERS,
    OUTLOOK_RULES,
    type OutlookRule,
    type RuleEvaluationResult,
} from "../config/outlook-rules.js";

type ToolResult = Promise<{ content: { type: string; text: string }[] }>;

export const actionPlannerTools: Tool[] = [
    {
        name: "email_plan_actions",
        description: `Generate an action plan for an email. Returns structured actions including Teams channel posts, direct messages, and Planner task creation. Uses routing rules to determine:
- P1.Urgent-Escalation: DM to Daniel + high-priority task due today
- P1.AP-Posting-Failure: Post to Ops/AP channel + Raft Desk Tickets plan
- P2.Raft-NotPosting: Post to Ops/AP + Raft Desk plan
- P2.Integration-Wove: Post to Wove channel + Wove plan
- P2.Integration-Project: Post to IT Integrations + Integrations plan
- P3.Warehouse-GHI: Daily digest unless due today
- P3.Task-Digests: Daily digest (muted)`,
        inputSchema: {
            type: "object",
            properties: {
                sender: {
                    type: "string",
                    description: "Email sender address",
                },
                subject: {
                    type: "string",
                    description: "Email subject line",
                },
                body: {
                    type: "string",
                    description: "Email body content",
                },
                outlookItemUrl: {
                    type: "string",
                    description: "Optional Outlook web link to the email for attachments",
                },
            },
            required: ["sender", "subject"],
        },
    },
    {
        name: "email_get_routing_config",
        description: "Get the current routing configuration including Teams channels, Planner plans, and DM recipients",
        inputSchema: {
            type: "object",
            properties: {},
        },
    },
    {
        name: "email_extract_cw1_entities",
        description: "Extract CargoWise/CW1-specific entities from email content: shipment IDs (S######), service request IDs (SR######), warehouse order IDs (W#######), creditor codes, transaction numbers, invoice numbers, and amounts",
        inputSchema: {
            type: "object",
            properties: {
                content: {
                    type: "string",
                    description: "Email content to analyze",
                },
            },
            required: ["content"],
        },
    },
    {
        name: "email_check_required_date",
        description: "Check if a warehouse order email has a Required Date of today (for urgent routing)",
        inputSchema: {
            type: "object",
            properties: {
                content: {
                    type: "string",
                    description: "Email content containing Required Date",
                },
            },
            required: ["content"],
        },
    },
    // OUTLOOK RULES TOOLS
    {
        name: "email_apply_outlook_rules",
        description: `Apply Outlook email rules to determine folder routing, forwarding, and mark-as-read actions. 
Rules include:
- Security/2FA codes → Forward to IT staff + file to 99-Security
- AP Posting failures → Mark read + file to TPA checking
- VIP senders (Brendan, Chris, Bruno, etc.) → File to 00-Action
- Integration emails (Wove, Raft, Cin7) → Route to dedicated folders
- Warehouse orders → File to Warehouse folder
- 3PL client domains → File to 3PL inbox
Returns matched rules and recommended actions.`,
        inputSchema: {
            type: "object",
            properties: {
                sender: {
                    type: "string",
                    description: "Email sender address",
                },
                subject: {
                    type: "string",
                    description: "Email subject line",
                },
                isInCC: {
                    type: "boolean",
                    description: "Whether the user's email is in the CC field",
                },
            },
            required: ["sender", "subject"],
        },
    },
    {
        name: "email_list_outlook_rules",
        description: "List all configured Outlook email rules with their conditions and actions. Optionally filter by category or enabled status.",
        inputSchema: {
            type: "object",
            properties: {
                category: {
                    type: "string",
                    description: "Filter by category (Security, Automation, VIP, Integrations, Systems, Operations, Vendors, 3PL, Warehouses)",
                },
                enabledOnly: {
                    type: "boolean",
                    description: "Only return enabled rules (default: true)",
                },
            },
            required: [],
        },
    },
    {
        name: "email_get_outlook_rule",
        description: "Get detailed information about a specific Outlook email rule by ID or name",
        inputSchema: {
            type: "object",
            properties: {
                ruleId: {
                    type: "number",
                    description: "Rule ID (1-23)",
                },
                ruleName: {
                    type: "string",
                    description: "Rule name (partial match supported)",
                },
            },
            required: [],
        },
    },
    {
        name: "email_get_outlook_folders",
        description: "Get list of all Outlook folders used by the email rules, with their configured purpose",
        inputSchema: {
            type: "object",
            properties: {},
            required: [],
        },
    },
];

export async function handleActionPlannerTool(name: string, args: Record<string, unknown>): ToolResult {
    switch (name) {
        case "email_plan_actions": {
            const sender = (args.sender as string | undefined) || "";
            const subject = (args.subject as string | undefined) || "";
            const body = (args.body as string | undefined) || "";
            const outlookItemUrl = args.outlookItemUrl as string | undefined;
            const plan = generateActionPlan(sender, subject, body, outlookItemUrl);
            console.log(`[ACTION-PLAN] ${sender} -> Bucket: ${plan.bucket}, Priority: P${plan.priority}, Actions: ${plan.actions.length}, Triggers: ${plan.triggers.join(", ") || "none"}`);
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(plan, null, 2),
                    },
                ],
            };
        }
        case "email_get_routing_config": {
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                teams_channels: TEAMS_CHANNELS,
                                planner_plans: PLANNER_PLANS,
                                dm_recipients: DM_RECIPIENTS,
                                vip_list: VIP_LIST,
                                patterns: Object.keys(REGEX_PATTERNS),
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }
        case "email_extract_cw1_entities": {
            const content = (args.content as string | undefined) || "";
            const entities = extractEntities(content);
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                entities,
                                found: Object.keys(entities).length > 0,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }
        case "email_check_required_date": {
            const content = (args.content as string | undefined) || "";
            if (!content) {
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                found: false,
                                isToday: false,
                                message: "No content provided",
                            }),
                        },
                    ],
                };
            }
            const dateMatch = content.match(/Required Date\s+([0-9-: ]+)/i);
            if (!dateMatch) {
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                found: false,
                                isToday: false,
                                message: "No Required Date found in content",
                            }),
                        },
                    ],
                };
            }
            const requiredDateStr = dateMatch[1].trim();
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            // Try to parse the date (handles formats like "2026-01-27" or "27-01-2026")
            let requiredDate: Date | null = null;
            try {
                // Try ISO format first
                requiredDate = new Date(requiredDateStr);
                if (Number.isNaN(requiredDate.getTime())) {
                    // Try DD-MM-YYYY format
                    const parts = requiredDateStr.split(/[-/]/);
                    if (parts.length >= 3) {
                        requiredDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
                    }
                }
            } catch {
                requiredDate = null;
            }
            if (!requiredDate || Number.isNaN(requiredDate.getTime())) {
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                found: true,
                                requiredDateStr,
                                isToday: false,
                                message: "Could not parse Required Date",
                            }),
                        },
                    ],
                };
            }
            requiredDate.setHours(0, 0, 0, 0);
            const isToday = requiredDate.getTime() === today.getTime();
            const isPast = requiredDate.getTime() < today.getTime();
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            found: true,
                            requiredDateStr,
                            requiredDate: requiredDate.toISOString().split("T")[0],
                            isToday,
                            isPast,
                            urgentRouting: isToday || isPast,
                            message: isToday
                                ? "Required Date is TODAY - route to urgent"
                                : isPast
                                    ? "Required Date is PAST - route to urgent"
                                    : "Required Date is in the future - queue for digest",
                        }),
                    },
                ],
            };
        }
        // OUTLOOK RULES HANDLERS
        case "email_apply_outlook_rules": {
            const sender = (args.sender as string | undefined) || "";
            const subject = (args.subject as string | undefined) || "";
            const isInCC = (args.isInCC as boolean | undefined) || false;
            const result: RuleEvaluationResult = evaluateRules(sender, subject, isInCC);
            console.log(`[OUTLOOK-RULES] ${sender} | Subject: "${subject.substring(0, 40)}..." -> ` +
                `Matched: ${result.matched}, Folder: ${result.actions.moveToFolder || "None"}, ` +
                `Forward: ${result.actions.forwardTo.length > 0 ? result.actions.forwardTo.join(", ") : "None"}, ` +
                `Triggers: ${result.triggers.join(", ") || "none"}`);
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            matched: result.matched,
                            matchedRuleCount: result.matchedRules.length,
                            matchedRules: result.matchedRules.map((r: OutlookRule) => ({
                                id: r.id,
                                name: r.name,
                                priority: r.priority,
                            })),
                            actions: result.actions,
                            triggers: result.triggers,
                            summary: result.matched
                                ? `Apply: ${result.actions.markAsRead ? "Mark Read, " : ""}${result.actions.moveToFolder ? `Move to "${result.actions.moveToFolder}"` : ""}${result.actions.forwardTo.length > 0 ? `, Forward to ${result.actions.forwardTo.join(", ")}` : ""}`
                                : "No matching rules found",
                        }, null, 2),
                    },
                ],
            };
        }
        case "email_list_outlook_rules": {
            const category = args.category as string | undefined;
            const enabledOnly = args.enabledOnly !== false; // Default to true
            let rules: OutlookRule[] = [...OUTLOOK_RULES];
            // Filter by enabled status
            if (enabledOnly) {
                rules = rules.filter((r: OutlookRule) => r.enabled);
            }
            // Filter by category
            if (category) {
                const categoryLower = category.toLowerCase();
                rules = rules.filter((r: OutlookRule) => {
                    const ruleCategory = r.name.toLowerCase();
                    return ruleCategory.includes(categoryLower) ||
                        r.description.toLowerCase().includes(categoryLower);
                });
            }
            // Sort by priority
            rules.sort((a: OutlookRule, b: OutlookRule) => a.priority - b.priority);
            const summary = rules.map((r: OutlookRule) => ({
                id: r.id,
                name: r.name,
                priority: r.priority,
                enabled: r.enabled,
                folder: r.actions.moveToFolder ? OUTLOOK_FOLDERS[r.actions.moveToFolder] : null,
                forwards: r.actions.forwardTo?.map((f) => FORWARD_RECIPIENTS[f]) || [],
                markAsRead: r.actions.markAsRead || false,
            }));
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            totalRules: OUTLOOK_RULES.length,
                            filteredCount: rules.length,
                            filter: { category, enabledOnly },
                            rules: summary,
                        }, null, 2),
                    },
                ],
            };
        }
        case "email_get_outlook_rule": {
            const ruleId = args.ruleId as number | undefined;
            const ruleName = args.ruleName as string | undefined;
            let rule: OutlookRule | undefined;
            if (ruleId !== undefined) {
                rule = OUTLOOK_RULES.find((r: OutlookRule) => r.id === ruleId);
            }
            else if (ruleName) {
                const nameLower = ruleName.toLowerCase();
                rule = OUTLOOK_RULES.find((r: OutlookRule) => r.name.toLowerCase().includes(nameLower));
            }
            if (!rule) {
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                found: false,
                                message: `No rule found with ${ruleId !== undefined ? `ID ${ruleId}` : `name containing "${ruleName}"`}`,
                                availableRules: getRulesSummary(),
                            }, null, 2),
                        },
                    ],
                };
            }
            // Resolve folder and forward recipients to actual values
            const resolvedRule = {
                ...rule,
                resolvedFolder: rule.actions.moveToFolder ? OUTLOOK_FOLDERS[rule.actions.moveToFolder] : null,
                resolvedForwardTo: rule.actions.forwardTo?.map((f) => FORWARD_RECIPIENTS[f]) || [],
            };
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            found: true,
                            rule: resolvedRule,
                        }, null, 2),
                    },
                ],
            };
        }
        case "email_get_outlook_folders": {
            // Group rules by folder
            const folderUsage: Record<string, { folder: string; rules: string[] }> = {};
            for (const rule of OUTLOOK_RULES) {
                if (rule.actions.moveToFolder) {
                    const folderKey = rule.actions.moveToFolder;
                    const folderName = OUTLOOK_FOLDERS[folderKey];
                    if (!folderUsage[folderName]) {
                        folderUsage[folderName] = { folder: folderName, rules: [] };
                    }
                    folderUsage[folderName].rules.push(rule.name);
                }
            }
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            folders: OUTLOOK_FOLDERS,
                            forwardRecipients: FORWARD_RECIPIENTS,
                            folderUsage: Object.values(folderUsage).sort((a, b) => a.folder.localeCompare(b.folder)),
                        }, null, 2),
                    },
                ],
            };
        }
        default:
            throw new Error(`Unknown action planner tool: ${name}`);
    }
}
