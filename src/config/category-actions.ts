/**
 * Category → Action Mapping Configuration
 *
 * This maps classification categories to post-processing actions:
 * - Planner bucket assignment
 * - Teams channel notifications
 * - DM recipients
 * - SLA and priority
 * - Response templates
 */
/**
 * CATEGORY → ACTION MAPPING
 *
 * Edit this to customize what happens after an email is classified.
 */
export interface CategoryAction {
    displayName: string;
    description: string;
    priority: number;
    priorityLabel: "Urgent" | "High" | "Medium" | "Low";
    sla: string;
    bucket: string;
    plannerPlan: string;
    labels: string[];
    checklist?: string[];
    teamsChannel?: string;
    dmRecipient?: string;
    createTask: boolean;
    sendNotification: boolean;
    queueForDigest: boolean;
    responseTemplate?: string;
    outlookCategories: string[];
}

export type CategoryActionsMap = Record<string, CategoryAction>;

export const CATEGORY_ACTIONS: CategoryActionsMap = {
    // ═══════════════════════════════════════════════════════════════════════════
    // P1: CRITICAL - Immediate action required
    // ═══════════════════════════════════════════════════════════════════════════
    vip_integration_error: {
        displayName: "VIP Integration Error",
        description: "Integration error from a VIP client - highest priority",
        priority: 1,
        priorityLabel: "Urgent",
        sla: "2hr",
        bucket: "Urgent",
        plannerPlan: "raft_desk_tickets",
        labels: ["P1.VIP-Integration", "VIP", "Integration"],
        checklist: [
            "Acknowledge to client within 30 mins",
            "Identify root cause",
            "Escalate to tech lead if not resolved in 1hr",
            "Update client with resolution",
        ],
        teamsChannel: "ops_ap",
        dmRecipient: "daniel_breglia",
        createTask: true,
        sendNotification: true,
        queueForDigest: false,
        responseTemplate: "vip_integration_error",
        outlookCategories: ["P1 - Urgent", "VIP Client", "Integration Error"],
    },
    escalation: {
        displayName: "Escalation",
        description: "Customer escalation or complaint requiring urgent attention",
        priority: 1,
        priorityLabel: "Urgent",
        sla: "1hr",
        bucket: "Urgent",
        plannerPlan: "email_tasks",
        labels: ["P1.Escalation", "Customer"],
        checklist: [
            "Acknowledge receipt immediately",
            "Review complaint details",
            "Prepare resolution plan",
            "Follow up within SLA",
        ],
        teamsChannel: "email_tasks",
        dmRecipient: "daniel_breglia",
        createTask: true,
        sendNotification: true,
        queueForDigest: false,
        responseTemplate: "escalation",
        outlookCategories: ["P1 - Urgent", "Escalation"],
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // P2: HIGH - Same-day response required
    // ═══════════════════════════════════════════════════════════════════════════
    vip: {
        displayName: "VIP Client",
        description: "Request from VIP client - priority handling",
        priority: 2,
        priorityLabel: "High",
        sla: "4hr",
        bucket: "VIP Follow-up",
        plannerPlan: "email_tasks",
        labels: ["P2.VIP", "VIP"],
        teamsChannel: "email_tasks",
        createTask: true,
        sendNotification: true,
        queueForDigest: false,
        responseTemplate: "vip_acknowledgement",
        outlookCategories: ["P2 - High Priority", "VIP Client"],
    },
    integration_error: {
        displayName: "Integration Error",
        description: "System integration error (API, CargoWise, Raft, etc.)",
        priority: 2,
        priorityLabel: "High",
        sla: "4hr",
        bucket: "Urgent",
        plannerPlan: "raft_desk_tickets",
        labels: ["P2.Integration", "Technical"],
        checklist: [
            "Check error logs",
            "Identify affected shipments",
            "Verify CW1 status",
            "Escalate if critical",
        ],
        teamsChannel: "ops_ap",
        createTask: true,
        sendNotification: true,
        queueForDigest: false,
        responseTemplate: "integration_error",
        outlookCategories: ["P2 - High Priority", "Integration Error"],
    },
    ap_posting: {
        displayName: "AP Posting Failure",
        description: "AP posting or ATP trigger failure - critical for accounting",
        priority: 1,
        priorityLabel: "Urgent",
        sla: "2hr",
        bucket: "Urgent",
        plannerPlan: "raft_desk_tickets",
        labels: ["P1.AP-Posting-Failure", "Accounting"],
        checklist: [
            "Check transaction details",
            "Verify creditor code",
            "Review ATP trigger logs",
            "Reprocess if needed",
        ],
        teamsChannel: "ops_ap",
        createTask: true,
        sendNotification: true,
        queueForDigest: false,
        responseTemplate: "ap_posting",
        outlookCategories: ["P1 - Urgent", "AP Posting"],
    },
    raft_cw1: {
        displayName: "Raft/CW1 Sync Issue",
        description: "Raft pushed but not posting to CargoWise",
        priority: 2,
        priorityLabel: "High",
        sla: "4hr",
        bucket: "P2.Raft→CW1-NotPosting",
        plannerPlan: "raft_desk_tickets",
        labels: ["P2.Raft-NotPosting", "Integration"],
        checklist: [
            "Check Raft sync status",
            "Verify CW1 connection",
            "Review shipment IDs",
            "Retry sync if needed",
        ],
        teamsChannel: "ops_ap",
        createTask: true,
        sendNotification: true,
        queueForDigest: false,
        responseTemplate: "raft_cw1",
        outlookCategories: ["P2 - High Priority", "Raft/CW1"],
    },
    wove: {
        displayName: "Wove Integration",
        description: "Wove-specific integration issue",
        priority: 2,
        priorityLabel: "High",
        sla: "4hr",
        bucket: "P2.Integration-Wove",
        plannerPlan: "wove_plan",
        labels: ["P2.Integration-Wove", "Wove"],
        teamsChannel: "wove",
        createTask: true,
        sendNotification: true,
        queueForDigest: false,
        responseTemplate: "wove",
        outlookCategories: ["P2 - High Priority", "Wove"],
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // P3: MEDIUM - Next business day
    // ═══════════════════════════════════════════════════════════════════════════
    warehouse: {
        displayName: "Warehouse Order/GHI",
        description: "Warehouse orders and Goods Handling Instructions",
        priority: 3,
        priorityLabel: "Medium",
        sla: "1day",
        bucket: "P3.Warehouse-Orders-GHI",
        plannerPlan: "raft_desk_tickets",
        labels: ["P3.Warehouse-GHI", "Warehouse"],
        teamsChannel: "email_tasks",
        createTask: true,
        sendNotification: false,
        queueForDigest: true,
        responseTemplate: "warehouse",
        outlookCategories: ["P3 - Medium", "Warehouse"],
    },
    order_status: {
        displayName: "Order/Shipment Status",
        description: "Order tracking, shipment status, delivery inquiries",
        priority: 3,
        priorityLabel: "Medium",
        sla: "1day",
        bucket: "Operations",
        plannerPlan: "email_tasks",
        labels: ["P3.Order-Status", "Operations"],
        teamsChannel: undefined, // No notification, just task
        createTask: true,
        sendNotification: false,
        queueForDigest: true,
        responseTemplate: "order_status",
        outlookCategories: ["P3 - Medium", "Order Status"],
    },
    client_inquiry: {
        displayName: "Client Inquiry",
        description: "General questions, support requests, information requests",
        priority: 3,
        priorityLabel: "Medium",
        sla: "1day",
        bucket: "Client Communications",
        plannerPlan: "email_tasks",
        labels: ["P3.Inquiry", "Client"],
        createTask: true,
        sendNotification: false,
        queueForDigest: true,
        responseTemplate: "client_inquiry",
        outlookCategories: ["P3 - Medium", "Client Inquiry"],
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // P4: LOW - Backlog
    // ═══════════════════════════════════════════════════════════════════════════
    general: {
        displayName: "General",
        description: "Unclassified emails - review and categorize manually",
        priority: 5,
        priorityLabel: "Low",
        sla: "2day",
        bucket: "Backlog",
        plannerPlan: "email_tasks",
        labels: ["P4.General", "Unclassified"],
        createTask: true,
        sendNotification: false,
        queueForDigest: true,
        responseTemplate: undefined, // No auto-response
        outlookCategories: ["P5 - Low/FYI"],
    },
};
/**
 * Outlook category mapping based on email content/sender
 * Uses custom descriptive categories (created automatically when first used)
 */
export interface OutlookCategoryRule {
    pattern: RegExp;
    category: string;
}

export const OUTLOOK_CATEGORY_RULES: OutlookCategoryRule[] = [
    // VIP clients
    { pattern: /tilesofezra|medifab|splash|yoni|bfglobal|bds/i, category: "VIP Client" },
    // Urgent keywords
    { pattern: /urgent|critical|asap|emergency|escalat/i, category: "P1 - Urgent" },
    // Integration/technical issues
    { pattern: /error|fail|exception|crash|down|not posting/i, category: "Integration Error" },
    // Raft/CW1 specific
    { pattern: /raft|cw1|cargowise|pushed but not/i, category: "Raft/CW1" },
    // Wove specific
    { pattern: /\bwove\b/i, category: "Wove" },
    // AP Posting
    { pattern: /ap posting|unable to post|atp trigger/i, category: "AP Posting" },
    // Warehouse/GHI
    { pattern: /warehouse order|goods handling|ghi|W\d{7}/i, category: "Warehouse" },
];
/**
 * Get Outlook categories for an email based on classification + content
 */
export function getOutlookCategories(classification: string, sender: string, subject: string, body?: string): string[] {
    const action = getCategoryAction(classification);
    const categories = new Set(action.outlookCategories);
    const content = `${sender} ${subject} ${body || ""}`.toLowerCase();
    // Add context-specific categories based on content
    for (const rule of OUTLOOK_CATEGORY_RULES) {
        if (rule.pattern.test(content)) {
            categories.add(rule.category);
        }
    }
    return Array.from(categories);
}
/**
 * Get action configuration for a classification category
 */
export function getCategoryAction(category: string): CategoryAction {
    return CATEGORY_ACTIONS[category] || CATEGORY_ACTIONS.general;
}
/**
 * Get all categories sorted by priority
 */
export function getCategoriesByPriority(): { category: string; action: CategoryAction; }[] {
    return Object.entries(CATEGORY_ACTIONS)
        .map(([category, action]) => ({ category: category, action }))
        .sort((a, b) => a.action.priority - b.action.priority);
}
/**
 * Summary for logging/display
 */
export function getCategorySummary(): string {
    const lines = ["Category → Action Mapping:", ""];
    for (const { category, action } of getCategoriesByPriority()) {
        lines.push(`P${action.priority} ${category.toUpperCase()}`);
        lines.push(`   Bucket: ${action.bucket} | SLA: ${action.sla}`);
        lines.push(`   Task: ${action.createTask ? "Yes" : "No"} | Notify: ${action.sendNotification ? "Yes" : "No"}`);
        if (action.teamsChannel)
            lines.push(`   Teams: #${action.teamsChannel}`);
        if (action.dmRecipient)
            lines.push(`   DM: ${action.dmRecipient}`);
        lines.push("");
    }
    return lines.join("\n");
}
