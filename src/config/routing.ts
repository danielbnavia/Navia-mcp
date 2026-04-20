/**

 * Routing Configuration for Email Triage

 *

 * This module provides routing rules for Teams channels, Planner plans, and DMs

 * based on email content patterns.

 */

// VIP Client List (by company name, not domain)

export const VIP_LIST = [

    "Splash Blanket",

    "Yoni Pleasure Palace",

    "BF Global Warehouse",

    "Tiles of Ezra",

    "BDS Animal Health",

];

// Executive escalation senders
export const EXEC_SENDERS = [
    "brendan.borg",
    "ian.fleming",
    "alana.raitt",
    "david.burns",
];

// Regex patterns for content matching

export const REGEX_PATTERNS = {

    exec_escalation: /(?:urgent|issue|escalate)/i,

    urgent: /(?:urgent|asap|shipment delay|pickup (?:today|now)|expo|need(?:ed)?.*asap|information.*needed)/i,

    ap_posting: /(?:unable to post transaction|ATP trigger)/i,

    raft_not_posting: /(?:pushed but not posting|Awaiting response from CargoWise)/i,

    raft_domain: /(?:@raft\.ai)/i,

    wove_domain: /(?:@wove\.com)/i,

    warehouse: /(?:Warehouse Order\s+W\d+|Goods Handling Instructions|GHI|putaway|put away|inward.*confirmation|outward.*confirmation|stock receipt|stock transfer)/i,

    planner: /(?:You have late tasks|Incomplete Tasks|Overdue Task Alert)/i,

    wove: /\bWove\b/i,

    integrations_generic: /(?:integration|Shopify|NaviaFill|Cin7|Odoo|NetSuite|MachShip|Logic App|Webhook|CW1 Integration|transmission issue|transmission error|transmission fail|sync issue|sync fail|raft issues|API|JSON|XML|CodeCom|3PL)/i,

};

// Entity extraction patterns

export const ENTITY_PATTERNS = {

    cw1_ids: /(S\d{6,}|SR\d{6,}|W\d{7,})/gi,

    creditor_code: /Creditor Code:\s*(\w+)/i,

    transaction_no: /Transaction number:\s*([\w-]+)/i,

    invoice_number: /Invoice Number\s*[:#]?\s*([A-Za-z0-9-]+)/i,

    amount_incl_tax: /Amount \(incl tax\):\s*([\d.]+)/i,

    required_date: /Required Date\s+([0-9-: ]+)/i,

};

// Teams Channel Configuration - CORRECT Team + Channel Mapping

// Each channel belongs to a DIFFERENT team - must use correct teamId!

export const TEAMS_CHANNELS = {

    // Raft AI Desk team -> RAFT.ai Desk channel

    ops_ap: {

        teamId: "6cdbd2a2-7e69-492e-ad7e-a617ed1c6597",

        channelId: "19:825f894d32b8473a9739bc5363b37dd8@thread.tacv2",

        name: "RAFT.ai Desk",

    },

    // Customer 3PL API Integrations team -> Wove-Consol channel

    wove: {

        teamId: "8c1829c3-aeb7-4af6-905d-a81023b3bebd",

        channelId: "19:11b039e14c124f2cb42155ac01ea80cc@thread.tacv2",

        name: "Wove-Consol",

    },

    // "3pl Integration" team -> "General" channel (logical name: IT Integrations)
    // Validated via Graph API 2026-02-14: team="3pl Integration", channel="General", membershipType=standard

    it_integrations: {

        teamId: "9fd6ec21-e570-4455-8338-9979d75d87e7",

        channelId: "19:Wodl4fnvVItdRGcdLQCuYf8NbNMmZWkBUvGceDoSJp01@thread.tacv2",

        name: "IT Integrations",

    },

    // "Customer 3PL API Integrations" team -> "General" channel (logical name: Email Tasks)
    // Validated via Graph API 2026-02-14: team="Customer 3PL API Integrations", channel="General", membershipType=standard

    email_tasks: {

        teamId: "8c1829c3-aeb7-4af6-905d-a81023b3bebd",

        channelId: "19:dCAbGRmp2ZNxjwjPCVC7iaKRp6XbJWz6fvO9uBEgalM1@thread.tacv2",

        name: "Email Tasks",

    },

} as const;

// Planner Plan IDs (from Graph API - NOT GUIDs!)

export const PLANNER_PLANS = {

    raft_desk_tickets: "JFnyzChtLkW76LsN3fZ6X8gAGW-V", // NF RAFT plan

    wove_plan: "isN5lApe9UKwJzUcBXpiWsgABOwz", // Email Tasks (no dedicated Wove plan)

    integrations_plan: "isN5lApe9UKwJzUcBXpiWsgABOwz", // Email Tasks (no dedicated plan)

    email_tasks: "isN5lApe9UKwJzUcBXpiWsgABOwz", // Email Tasks

} as const;

// Direct Message Recipients

export const DM_RECIPIENTS = {

    daniel_breglia: "Danielb@naviafreight.com",

} as const;

export type ActionType = "teams.post" | "teams.dm" | "planner.create" | "digest.queue" | "mute";

export interface ActionParams {
    teamId?: string;
    channelId?: string;
    channelName?: string;
    title?: string;
    text?: string;
    to?: string;
    planId?: string;
    planName?: string;
    priority?: "Low" | "Medium" | "High" | "Urgent";
    due?: "today" | "tomorrow" | string;
    labels?: string[];
    attachments?: string[];
    checklist?: string[];
    digestType?: "daily" | "weekly";
    reason?: string;
}

export interface Action {
    type: ActionType;
    params: ActionParams;
}

export interface ActionPlanResponse {
    bucket: string;
    priority: number;
    priorityLabel: string;
    actions: Action[];
    entities: {
        cw1_ids?: string[];
        creditor_code?: string;
        transaction_no?: string;
        invoice_number?: string;
        amount?: string;
        required_date?: string;
    };
    triggers: string[];
    outlookCategories?: string[];
}

interface RoutingRule {
    pattern: keyof typeof REGEX_PATTERNS;
    priority: number;
    bucket: string;
    outlookCategories?: string[];
    actions: {
        teamsChannel?: keyof typeof TEAMS_CHANNELS;
        plannerPlan?: keyof typeof PLANNER_PLANS;
        dm?: keyof typeof DM_RECIPIENTS;
        due?: "today" | "tomorrow";
        priority?: "Low" | "Medium" | "High" | "Urgent";
        labels?: string[];
        checklist?: string[];
        digest?: "daily" | "weekly";
    };
}

// Ordered routing rules (first match wins for primary classification)

export const ROUTING_RULES: RoutingRule[] = [

    // P0: Exec escalation overlay -> DM + Risk/Escalation category

    {

        pattern: "exec_escalation",

        priority: 0,

        bucket: "Urgent",

        outlookCategories: ["Risk / Escalation"],

        actions: {

            dm: "daniel_breglia",

            plannerPlan: "email_tasks",

            priority: "Urgent",

            due: "today",

            labels: ["P0.Exec-Escalation"],

        },

    },

    // P1: Urgent escalation -> DM + due-today task

    {

        pattern: "urgent",

        priority: 1,

        bucket: "Urgent",

        outlookCategories: ["Action Required"],

        actions: {

            dm: "daniel_breglia",

            plannerPlan: "email_tasks",

            priority: "High",

            due: "today",

            labels: ["P1.Urgent-Escalation"],

            checklist: [

                "Confirm warehouse status",

                "Contact client with ETA",

                "Add carrier/POD link",

            ],

        },

    },

    // P1: AP Posting Failure

    {

        pattern: "ap_posting",

        priority: 1,

        bucket: "Urgent",

        outlookCategories: ["Action Required", "RAFT"],

        actions: {

            teamsChannel: "ops_ap",

            plannerPlan: "raft_desk_tickets",

            priority: "High",

            labels: ["P1.AP-Posting-Failure"],

            checklist: [

                "Check ATP trigger status",

                "Verify transaction in CW1",

                "Contact accounts if needed",

            ],

        },

    },

    // P2: Raft not posting to CW1

    {

        pattern: "raft_not_posting",

        priority: 2,

        bucket: "Operations",

        outlookCategories: ["Action Required", "RAFT"],

        actions: {

            teamsChannel: "ops_ap",

            plannerPlan: "raft_desk_tickets",

            priority: "Medium",

            labels: ["P2.Raft-NotPosting"],

        },

    },

    // P2: Wove integration

    {

        pattern: "wove",

        priority: 2,

        bucket: "Operations",

        outlookCategories: ["In Progress", "Wove"],

        actions: {

            teamsChannel: "wove",

            plannerPlan: "wove_plan",

            priority: "Medium",

            labels: ["P2.Integration-Wove"],

        },

    },

    // P2: Generic integrations (Shopify, Cin7, etc.)

    {

        pattern: "integrations_generic",

        priority: 2,

        bucket: "Operations",

        outlookCategories: ["In Progress", "Integrations \u2013 Active"],

        actions: {

            teamsChannel: "it_integrations",

            plannerPlan: "integrations_plan",

            priority: "Medium",

            labels: ["P2.Integration-Project"],

        },

    },

    // P3: Warehouse orders - queue for daily digest unless due today

    {

        pattern: "warehouse",

        priority: 3,

        bucket: "Operations",

        outlookCategories: ["In Progress"],

        actions: {

            digest: "daily",

            plannerPlan: "email_tasks",

            priority: "Low",

            labels: ["P3.Warehouse-GHI"],

        },

    },

    // P3: Planner task digests - mute/daily digest

    {

        pattern: "planner",

        priority: 4,

        bucket: "Backlog",

        actions: {

            digest: "daily",

            labels: ["P3.Task-Digests"],

        },

    },

];

/**

 * Extract entities from email content

 */

export function extractEntities(content: string): ActionPlanResponse["entities"] {

    const entities: ActionPlanResponse["entities"] = {};

    const cw1Match = content.match(ENTITY_PATTERNS.cw1_ids);

    if (cw1Match)

        entities.cw1_ids = [...new Set(cw1Match)];

    const creditorMatch = content.match(ENTITY_PATTERNS.creditor_code);

    if (creditorMatch)

        entities.creditor_code = creditorMatch[1];

    const txnMatch = content.match(ENTITY_PATTERNS.transaction_no);

    if (txnMatch)

        entities.transaction_no = txnMatch[1];

    const invoiceMatch = content.match(ENTITY_PATTERNS.invoice_number);

    if (invoiceMatch)

        entities.invoice_number = invoiceMatch[1];

    const amountMatch = content.match(ENTITY_PATTERNS.amount_incl_tax);

    if (amountMatch)

        entities.amount = amountMatch[1];

    const dateMatch = content.match(ENTITY_PATTERNS.required_date);

    if (dateMatch)

        entities.required_date = dateMatch[1];

    return entities;

}

/**

 * Generate action plan based on email content

 */

export function generateActionPlan(sender: string, subject: string, body: string, outlookItemUrl?: string): ActionPlanResponse {

    const content = `${subject} ${body}`.toLowerCase();

    const triggers: string[] = [];

    const actions: Action[] = [];

    let bucket = "Backlog";

    let priority = 5;

    let priorityLabel = "Medium";

    // Extract entities first

    const entities = extractEntities(`${subject} ${body}`);

    // Check sender to prevent self-DM loops

    const senderLower = sender.toLowerCase();

    const isDaniel = senderLower.includes("danielb@naviafreight.com");

    // Find matching rules

    for (const rule of ROUTING_RULES) {

        const pattern = REGEX_PATTERNS[rule.pattern];

        if (pattern.test(content)) {

            triggers.push(rule.pattern);

            // Apply rule settings

            if (rule.priority < priority || priority === 5) {

                priority = rule.priority;

                bucket = rule.bucket;

                priorityLabel = rule.priority === 1 ? "Urgent" : rule.priority === 2 ? "High" : rule.priority === 3 ? "Medium" : "Low";

            }

            // Build actions

            const { actions: ruleActions } = rule;

            // Teams DM (skip if sender is the DM recipient)

            if (ruleActions.dm && !isDaniel) {

                actions.push({

                    type: "teams.dm",

                    params: {

                        to: DM_RECIPIENTS[ruleActions.dm],

                        title: `P${rule.priority} ${rule.pattern.toUpperCase()}: ${subject.substring(0, 50)}`,

                        text: `From: ${sender}\nReason: ${rule.pattern}\nCW1 IDs: ${entities.cw1_ids?.join(", ") || "N/A"}${outlookItemUrl ? `\nLink: ${outlookItemUrl}` : ""}`,

                    },

                });

            }

            // Teams channel post

            if (ruleActions.teamsChannel) {

                const channel = TEAMS_CHANNELS[ruleActions.teamsChannel];

                actions.push({

                    type: "teams.post",

                    params: {

                        teamId: channel.teamId,

                        channelId: channel.channelId,

                        channelName: channel.name,

                        title: subject,

                        text: `From: ${sender}\nCategory: ${rule.pattern}\nCW1 IDs: ${entities.cw1_ids?.join(", ") || "N/A"}`,

                    },

                });

            }

            // Planner task

            if (ruleActions.plannerPlan) {

                actions.push({

                    type: "planner.create",

                    params: {

                        planId: PLANNER_PLANS[ruleActions.plannerPlan],

                        planName: ruleActions.plannerPlan,

                        title: subject,

                        priority: ruleActions.priority || "Medium",

                        due: ruleActions.due,

                        labels: ruleActions.labels,

                        attachments: outlookItemUrl ? [outlookItemUrl] : undefined,

                        checklist: ruleActions.checklist,

                    },

                });

            }

            // Digest queue

            if (ruleActions.digest) {

                actions.push({

                    type: "digest.queue",

                    params: {

                        digestType: ruleActions.digest,

                        reason: rule.pattern,

                    },

                });

            }

            // Only apply first matching rule for primary classification

            // but continue to collect all triggers

        }

    }

    // If no rules matched, default to email_tasks plan

    if (actions.length === 0) {

        actions.push({

            type: "planner.create",

            params: {

                planId: PLANNER_PLANS.email_tasks,

                planName: "email_tasks",

                title: subject,

                priority: "Low",

                labels: ["Unclassified"],

            },

        });

    }

    return {

        bucket,

        priority,

        priorityLabel,

        actions,

        entities,

        triggers,

    };

}
