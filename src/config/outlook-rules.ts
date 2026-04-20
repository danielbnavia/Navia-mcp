/**
 * Outlook Email Rules Configuration
 *
 * These rules define automatic email filing, forwarding, and marking based on
 * sender and subject patterns. Rules are evaluated in priority order (lower = higher priority).
 *
 * Actions:
 * - markAsRead: Mark email as read
 * - moveToFolder: Move to specified Outlook folder
 * - forwardTo: Forward to specified recipient(s)
 */
// Forward recipients (email addresses)
export const FORWARD_RECIPIENTS = {
    ian_fleming: "ian.fleming@naviafreight.com",
    jeremy_teschky: "jeremy.teschky@naviafreight.com",
    lerma_razal: "lerma.razal@naviafreight.com",
    melissa_rakai: "melissa.rakai@naviafreight.com",
} as const;
// Outlook folder names (must match actual mailbox folder names)
export const OUTLOOK_FOLDERS = {
    tpa_checking: "TPA checking",
    security: "99-Security",
    action: "00-Action",
    wove: "Wove",
    systems: "05-Systems",
    tests: "_Tests",
    planning: "Planning",
    read: "read",
    customer_ticket: "Customer Ticket",
    sea_cargo: "Sea Cargo Messages",
    warehouse: "Warehouse",
    vendors_carriers: "02-Vendors & Carriers",
    shipping_line_cld: "Shipping Line / CLD Invoice",
    threePL_inbox: "3PL inbox",
    europe_warehouse: "Europe Warehouse",
    chicago_warehouse: "Chicago Warehouse",
    sydney_warehouse: "Sydney Warehouse",
    raft_emails: "RAFT emails",
    cin7: "Cin7",
    integrations: "04-Integrations",
    netsuite: "NetSuite",
    shopify: "Shopify",
    jira_notifications: "Jira Notifications",
    notifications: "Notifications",
} as const;

export interface RuleCondition {
    senderEmail?: string | string[];
    senderContains?: string | string[];
    senderDomain?: string | string[];
    senderDomainNot?: string | string[];
    senderEmailNot?: string | string[];
    subjectContains?: string | string[];
    subjectContainsAll?: string[];
    and?: RuleCondition[];
    or?: RuleCondition[];
    inCC?: boolean;
}

export interface RuleAction {
    markAsRead?: boolean;
    moveToFolder?: keyof typeof OUTLOOK_FOLDERS;
    forwardTo?: (keyof typeof FORWARD_RECIPIENTS)[];
}

export interface OutlookRule {
    id: number;
    name: string;
    description: string;
    enabled: boolean;
    priority: number;
    conditions: RuleCondition;
    actions: RuleAction;
    stopProcessing?: boolean;
}
/**
 * OUTLOOK RULES - Ordered by priority
 *
 * Rule evaluation: First matching rule wins (unless stopProcessing is false)
 */
export const OUTLOOK_RULES: OutlookRule[] = [
    // ═══════════════════════════════════════════════════════════════════════════
    // SECURITY & 2FA RULES (Highest Priority)
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 2,
        name: "ShipStation Login Code",
        description: "Forward ShipStation login codes to Ian and file to security",
        enabled: true,
        priority: 1,
        conditions: {
            senderEmail: "noreply@shipstation.com",
            subjectContains: "login code",
        },
        actions: {
            forwardTo: ["ian_fleming"],
            moveToFolder: "security",
        },
        stopProcessing: true,
    },
    {
        id: 3,
        name: "Shippo 2FA Code",
        description: "Forward Shippo verification codes to Jeremy and file to security",
        enabled: true,
        priority: 2,
        conditions: {
            senderEmail: "noreply@goshippo.com",
            subjectContains: ["verification", "code"],
        },
        actions: {
            forwardTo: ["jeremy_teschky"],
            moveToFolder: "security",
        },
        stopProcessing: true,
    },
    {
        id: 4,
        name: "Security - Generic OTP/Verification",
        description: "File all OTP and verification code emails to security folder",
        enabled: true,
        priority: 3,
        conditions: {
            subjectContains: ["OTP", "verification code", "security code", "one-time"],
        },
        actions: {
            moveToFolder: "security",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // AUTOMATION & SYSTEM FAILURES
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 1,
        name: "AP Posting Failures",
        description: "Mark and file AP posting failures from CargoWise/Edinlpmel",
        enabled: true,
        priority: 10,
        conditions: {
            senderContains: ["edinlpmel", "Edinlpmel Navia", "AP Posting"],
            subjectContains: "Unable to post transaction",
        },
        actions: {
            markAsRead: true,
            moveToFolder: "tpa_checking",
        },
        stopProcessing: true,
    },
    {
        id: 5,
        name: "Automation Failures",
        description: "Route Power Automate and workflow failures to action folder",
        enabled: true,
        priority: 11,
        conditions: {
            subjectContains: ["WorkflowManager", "Trigger Action Failure", "flow failed", "run failed"],
        },
        actions: {
            moveToFolder: "action",
        },
        stopProcessing: true,
    },
    {
        id: 29,
        name: "Power Automate Failures",
        description: "Route Power Automate failure notifications to action folder",
        enabled: true,
        priority: 12,
        conditions: {
            senderEmail: "PowerAutomateNoReply@microsoft.com",
            subjectContains: ["failed", "error", "couldn't"],
        },
        actions: {
            moveToFolder: "action",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // VIP & APPROVALS
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 6,
        name: "VIP - External Contacts",
        description: "Route external VIP emails to 3PL inbox with VIP tracking (SLA escalation via Power Automate)",
        enabled: true,
        priority: 20,
        conditions: {
            or: [
                { senderEmail: "emily@tilesofezra.com" },
                { senderEmail: "may.zhou@apsystems.com" },
                { senderEmail: "darrel@tilesofezra.com" },
                { senderEmail: "charlie@lbedesign.com" },
                { senderEmail: "janis.habner@apsystems.com" },
            ],
        },
        actions: {
            moveToFolder: "threePL_inbox", // Route to 3PL inbox, Power Automate handles SLA escalation
        },
        stopProcessing: true,
        // Note: Power Automate flow will track VIP emails and escalate to 00-Action 30 min before SLA
    },
    {
        id: 38,
        name: "VIP - Internal Executives",
        description: "Route emails from internal executives to action folder",
        enabled: true,
        priority: 21,
        conditions: {
            senderEmail: [
                "brendanb@naviafreight.com",
                "chrisf@naviafreight.com",
                "brunos@naviafreight.com",
                "mandos@naviafreight.com",
            ],
        },
        actions: {
            moveToFolder: "action",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // INTEGRATIONS
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 7,
        name: "Integrations - Wove",
        description: "Route Wove emails to Wove folder",
        enabled: true,
        priority: 30,
        conditions: {
            senderDomain: "wove.com",
        },
        actions: {
            moveToFolder: "wove",
        },
        stopProcessing: true,
    },
    {
        id: 22,
        name: "Integrations - Raft",
        description: "Route Raft.ai and Vector AI Jira tickets to RAFT folder",
        enabled: true,
        priority: 31,
        conditions: {
            senderDomain: ["raft.ai", "vectorai.atlassian.net"],
        },
        actions: {
            moveToFolder: "raft_emails",
        },
        stopProcessing: true,
    },
    {
        id: 31,
        name: "Raft Pack Notifications",
        description: "Route Raft pack/shipment notifications to RAFT folder",
        enabled: true,
        priority: 30, // Before domain-based Raft rule (31)
        conditions: {
            senderEmail: "no-reply@raft.ai",
            subjectContains: ["pack", "shipment", "order"],
        },
        actions: {
            moveToFolder: "raft_emails",
        },
        stopProcessing: true,
    },
    {
        id: 23,
        name: "Integrations - Cin7",
        description: "Route Cin7 integration emails (DISABLED - now goes to 3PL inbox)",
        enabled: false, // Disabled - Cin7 emails now route to 3PL inbox via rule 18
        priority: 32,
        conditions: {
            senderDomain: "cin7.com",
        },
        actions: {
            moveToFolder: "cin7",
        },
        stopProcessing: true,
    },
    {
        id: 24,
        name: "Integrations - Cin7 Support",
        description: "Route Cin7 Omni support tickets (DISABLED - now goes to 3PL inbox)",
        enabled: false, // Disabled - Cin7 emails now route to 3PL inbox via rule 18
        priority: 31,
        conditions: {
            senderEmail: "omnihelpresponse@cin7.com",
        },
        actions: {
            moveToFolder: "cin7",
        },
        stopProcessing: true,
    },
    {
        id: 25,
        name: "Integrations - MachShip",
        description: "Route MachShip integration emails",
        enabled: true,
        priority: 34,
        conditions: {
            senderDomain: "machship.com",
        },
        actions: {
            moveToFolder: "integrations",
        },
        stopProcessing: true,
    },
    {
        id: 26,
        name: "Integrations - NetSuite",
        description: "Route NetSuite/Oracle notifications",
        enabled: true,
        priority: 35,
        conditions: {
            senderDomain: ["netsuite.com", "oracle.com"],
            subjectContains: ["NetSuite", "workflow", "notification"],
        },
        actions: {
            moveToFolder: "netsuite",
        },
        stopProcessing: true,
    },
    {
        id: 27,
        name: "Integrations - Shopify Flow",
        description: "Route Shopify Flow notifications",
        enabled: true,
        priority: 36,
        conditions: {
            senderEmail: "flow@shopify.com",
        },
        actions: {
            moveToFolder: "shopify",
        },
        stopProcessing: true,
    },
    {
        id: 28,
        name: "Integrations - Shopify General",
        description: "Route general Shopify notifications",
        enabled: true,
        priority: 37,
        conditions: {
            senderDomain: "shopify.com",
        },
        actions: {
            moveToFolder: "shopify",
        },
        stopProcessing: true,
    },
    {
        id: 37,
        name: "Jira Notifications",
        description: "Route Jira/Atlassian notifications to Jira Notifications folder",
        enabled: true,
        priority: 38,
        conditions: {
            senderDomain: ["atlassian.net", "atlassian.com", "jira.com"],
        },
        actions: {
            moveToFolder: "jira_notifications",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // SYSTEM DIGESTS & AUTOMATED NOTIFICATIONS
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 8,
        name: "Systems Digests",
        description: "Mark and file ShipStation/Stamps.com system notifications",
        enabled: true,
        priority: 40,
        conditions: {
            senderEmail: ["no-reply@shipstation.com", "no-reply@stamps.com"],
        },
        actions: {
            markAsRead: true,
            moveToFolder: "systems",
        },
        stopProcessing: true,
    },
    {
        id: 13,
        name: "Customs & ICS",
        description: "Route customs declarations and status messages to systems folder",
        enabled: true,
        priority: 41,
        conditions: {
            subjectContains: [
                "Declaration Status Advice",
                "Import Declaration Response",
                "Authority to Deal",
                "Status Advice Message",
                "SAM",
                "Cargo Status Advice",
                "1-Stop",
                "Vessel Difference",
                "CARST",
                "IMDR",
                "ATD",
                "DSA",
            ],
        },
        actions: {
            moveToFolder: "systems",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // SELF-MAIL & TESTING
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 9,
        name: "Self-mail / Tests",
        description: "Route emails from Daniel Breglia to tests folder",
        enabled: true,
        priority: 150, // Lower priority - let other rules match first
        conditions: {
            senderContains: "Daniel Breglia",
        },
        actions: {
            moveToFolder: "tests",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // PLANNING & TEAMS
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 10,
        name: "Planning (Teams channel)",
        description: "Route Teams/Microsoft planning notifications",
        enabled: true,
        priority: 60,
        conditions: {
            and: [
                { senderContains: ["teams", "microsoft"] },
                { subjectContains: "Planning" },
            ],
        },
        actions: {
            moveToFolder: "planning",
        },
        stopProcessing: true,
    },
    {
        id: 30,
        name: "Planner Late Tasks",
        description: "Route Planner late/overdue task notifications",
        enabled: true,
        priority: 61,
        conditions: {
            senderEmail: "planner.office365.com",
            subjectContains: ["late", "overdue", "due today"],
        },
        actions: {
            markAsRead: true,
            moveToFolder: "planning",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // CC FORWARDING
    // ═══════════════════════════════════════════════════════════════════════════
    // ═══════════════════════════════════════════════════════════════════════════
    // SUBJECT-BASED ROUTING (for internal replies to specific topics)
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 39,
        name: "Subject - RAFT/Jira Replies",
        description: "Route replies about RAFT/Jira tickets to RAFT folder",
        enabled: true,
        priority: 65,
        conditions: {
            subjectContains: ["CS-", "RAFT", "raft.ai", "Vector AI", "pack", "Navia-"],
        },
        actions: {
            moveToFolder: "raft_emails",
        },
        stopProcessing: true,
    },
    {
        id: 40,
        name: "Subject - 3PL/Integration Replies",
        description: "Route replies about 3PL/integration topics to 3PL inbox",
        enabled: true,
        priority: 66,
        conditions: {
            subjectContains: ["3PL", "Cin7", "integration", "Shopify", "CargoWise", "CST0000"],
        },
        actions: {
            moveToFolder: "threePL_inbox",
        },
        stopProcessing: true,
    },
    {
        id: 41,
        name: "Marketing/Notifications",
        description: "Route marketing and notification emails to Notifications folder (excludes known vendors)",
        enabled: true,
        priority: 67,
        conditions: {
            and: [
                { senderDomainNot: ["vanguardlogistics.com", "fedex.com", "auspost.com.au", "dhl.com", "shipstation.com"] }, // Exclude vendors
                {
                    or: [
                        { subjectContains: ["unsubscribe", "newsletter", "webinar", "invitation"] },
                        { senderContains: ["marketing", "news@"] }, // Removed noreply/no-reply - too broad
                    ],
                },
            ],
        },
        actions: {
            markAsRead: true,
            moveToFolder: "notifications",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // CC FORWARDING (Last priority)
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 11,
        name: "CC Forward to Lerma",
        description: "Forward emails where user is in CC to Lerma",
        enabled: true,
        priority: 200, // Last priority - let other rules match first
        conditions: {
            inCC: true,
        },
        actions: {
            forwardTo: ["lerma_razal"],
            moveToFolder: "read",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // CUSTOMER SUPPORT
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 12,
        name: "Customer Tickets",
        description: "Route customer support tickets to dedicated folder",
        enabled: true,
        priority: 80,
        conditions: {
            subjectContains: ["ticket", "support request"],
        },
        actions: {
            moveToFolder: "customer_ticket",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // OPERATIONS
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 14,
        name: "Sea Cargo / Ops Reports",
        description: "Route sea cargo and operations reports",
        enabled: true,
        priority: 90,
        conditions: {
            subjectContains: ["sea cargo", "ops report", "vessel"],
        },
        actions: {
            moveToFolder: "sea_cargo",
        },
        stopProcessing: true,
    },
    {
        id: 15,
        name: "Warehouse Orders",
        description: "Route warehouse order notifications (KFL/TTCC/General)",
        enabled: true,
        priority: 91,
        conditions: {
            subjectContains: ["Warehouse Order", "KFL Warehouse"],
        },
        actions: {
            moveToFolder: "warehouse",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // VENDORS & CARRIERS
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 16,
        name: "Vendors - Vanguard",
        description: "Forward Vanguard emails to Melissa and file to vendors",
        enabled: true,
        priority: 100,
        conditions: {
            senderContains: "vanguard",
        },
        actions: {
            forwardTo: ["melissa_rakai"],
            moveToFolder: "vendors_carriers",
        },
        stopProcessing: true,
    },
    {
        id: 17,
        name: "Shipping Lines / CLD / Arrival",
        description: "Route shipping line notifications and arrival notices",
        enabled: true,
        priority: 101,
        conditions: {
            subjectContains: ["arrival notice", "CLD", "bill of lading"],
        },
        actions: {
            moveToFolder: "shipping_line_cld",
        },
        stopProcessing: true,
    },
    {
        id: 32,
        name: "Carriers - FedEx",
        description: "Route FedEx tracking and delivery notifications",
        enabled: true,
        priority: 102,
        conditions: {
            senderDomain: "fedex.com",
        },
        actions: {
            moveToFolder: "vendors_carriers",
        },
        stopProcessing: true,
    },
    {
        id: 33,
        name: "Carriers - Australia Post",
        description: "Route Australia Post notifications",
        enabled: true,
        priority: 103,
        conditions: {
            senderDomain: ["auspost.com.au", "startrack.com.au"],
        },
        actions: {
            moveToFolder: "vendors_carriers",
        },
        stopProcessing: true,
    },
    {
        id: 34,
        name: "Carriers - CMA CGM",
        description: "Route CMA CGM shipping line notifications",
        enabled: true,
        priority: 104,
        conditions: {
            senderDomain: ["cma-cgm.com", "cmacgm.com"],
        },
        actions: {
            moveToFolder: "shipping_line_cld",
        },
        stopProcessing: true,
    },
    {
        id: 35,
        name: "Carriers - DHL",
        description: "Route DHL tracking and delivery notifications",
        enabled: true,
        priority: 105,
        conditions: {
            senderDomain: ["dhl.com", "dhl.com.au"],
        },
        actions: {
            moveToFolder: "vendors_carriers",
        },
        stopProcessing: true,
    },
    {
        id: 36,
        name: "Carriers - TNT",
        description: "Route TNT/FedEx Express notifications",
        enabled: true,
        priority: 106,
        conditions: {
            senderDomain: ["tnt.com", "tnt.com.au"],
        },
        actions: {
            moveToFolder: "vendors_carriers",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // 3PL INTEGRATION CLIENTS
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 18,
        name: "3PL Integration Leads",
        description: "Route 3PL client integration emails",
        enabled: true,
        priority: 110,
        conditions: {
            senderDomain: [
                "medifab.com",
                "lbedesign.com",
                "apsystems.com",
                "tilesofezra.com",
                "uskidsgolf.com",
                "itransition.com",
                "splashblanket.com",
                "yonipleasurepalace.com",
                "fitonportal.com",
                "bgiworldwide.com",
                "thedevelopment.com.au",
                "cin7.com", // Added - Cin7 integration emails now go to 3PL inbox
            ],
        },
        actions: {
            moveToFolder: "threePL_inbox",
        },
        stopProcessing: true,
    },
    // ═══════════════════════════════════════════════════════════════════════════
    // REGIONAL WAREHOUSES
    // ═══════════════════════════════════════════════════════════════════════════
    {
        id: 19,
        name: "Warehouses - Europe",
        description: "Route European warehouse emails",
        enabled: true,
        priority: 120,
        conditions: {
            senderContains: ["europe", "eu warehouse"],
        },
        actions: {
            moveToFolder: "europe_warehouse",
        },
        stopProcessing: true,
    },
    {
        id: 20,
        name: "Warehouses - Chicago",
        description: "Route Chicago warehouse emails",
        enabled: true,
        priority: 121,
        conditions: {
            senderContains: "chicago",
        },
        actions: {
            moveToFolder: "chicago_warehouse",
        },
        stopProcessing: true,
    },
    {
        id: 21,
        name: "Warehouses - Sydney",
        description: "Route Sydney warehouse emails",
        enabled: true,
        priority: 122,
        conditions: {
            senderContains: "sydney",
        },
        actions: {
            moveToFolder: "sydney_warehouse",
        },
        stopProcessing: true,
    },
];
// ═══════════════════════════════════════════════════════════════════════════
// RULE EVALUATION FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════
/**
 * Check if a string matches any of the patterns (case-insensitive)
 */
function matchesAny(value: string, patterns: string | string[]): boolean {
    const patternList = Array.isArray(patterns) ? patterns : [patterns];
    const valueLower = value.toLowerCase();
    return patternList.some(p => valueLower.includes(p.toLowerCase()));
}
/**
 * Check if a string matches all patterns (case-insensitive)
 */
function matchesAll(value: string, patterns: string[]): boolean {
    const valueLower = value.toLowerCase();
    return patterns.every(p => valueLower.includes(p.toLowerCase()));
}
/**
 * Extract domain from email address
 */
function extractDomain(email: string): string {
    const match = email.match(/@([a-zA-Z0-9.-]+)/);
    return match ? match[1].toLowerCase() : "";
}
/**
 * Evaluate a single condition against email data
 */
function evaluateCondition(condition: RuleCondition, sender: string, subject: string, isInCC: boolean = false): boolean {
    // Handle OR conditions
    if (condition.or) {
        return condition.or.some(c => evaluateCondition(c, sender, subject, isInCC));
    }
    // Handle AND conditions
    if (condition.and) {
        return condition.and.every(c => evaluateCondition(c, sender, subject, isInCC));
    }
    // Check inCC condition
    if (condition.inCC !== undefined) {
        if (condition.inCC && !isInCC)
            return false;
    }
    // Check sender email (exact match)
    if (condition.senderEmail) {
        const emails = Array.isArray(condition.senderEmail) ? condition.senderEmail : [condition.senderEmail];
        const senderLower = sender.toLowerCase();
        if (!emails.some(e => senderLower.includes(e.toLowerCase()))) {
            return false;
        }
    }
    // Check sender contains
    if (condition.senderContains) {
        if (!matchesAny(sender, condition.senderContains)) {
            return false;
        }
    }
    // Check sender domain
    if (condition.senderDomain) {
        const domains = Array.isArray(condition.senderDomain) ? condition.senderDomain : [condition.senderDomain];
        const senderDomain = extractDomain(sender);
        if (!domains.some(d => senderDomain === d.toLowerCase() || senderDomain.endsWith(`.${d.toLowerCase()}`))) {
            return false;
        }
    }
    // Check sender domain exclusion (NOT)
    if (condition.senderDomainNot) {
        const excludeDomains = Array.isArray(condition.senderDomainNot) ? condition.senderDomainNot : [condition.senderDomainNot];
        const senderDomain = extractDomain(sender);
        if (excludeDomains.some(d => senderDomain === d.toLowerCase() || senderDomain.endsWith(`.${d.toLowerCase()}`))) {
            return false; // Excluded domain matched, rule doesn't apply
        }
    }
    // Check sender email exclusion (NOT)
    if (condition.senderEmailNot) {
        const excludeEmails = Array.isArray(condition.senderEmailNot) ? condition.senderEmailNot : [condition.senderEmailNot];
        const senderLower = sender.toLowerCase();
        if (excludeEmails.some(e => senderLower.includes(e.toLowerCase()))) {
            return false; // Excluded email matched, rule doesn't apply
        }
    }
    // Check subject contains (any)
    if (condition.subjectContains) {
        if (!matchesAny(subject, condition.subjectContains)) {
            return false;
        }
    }
    // Check subject contains all
    if (condition.subjectContainsAll) {
        if (!matchesAll(subject, condition.subjectContainsAll)) {
            return false;
        }
    }
    return true;
}
/**
 * Evaluate all rules against an email and return matching actions
 */
export interface RuleEvaluationResult {
    matched: boolean;
    matchedRules: OutlookRule[];
    actions: {
        markAsRead: boolean;
        moveToFolder: string | null;
        forwardTo: string[];
    };
    triggers: string[];
}

/**
 * Evaluate all rules against an email and return matching actions
 */
export function evaluateRules(sender: string, subject: string, isInCC: boolean = false): RuleEvaluationResult {
    const matchedRules: OutlookRule[] = [];
    const triggers: string[] = [];
    let markAsRead = false;
    let moveToFolder: string | null = null;
    const forwardTo: string[] = [];
    // Sort rules by priority (lower = higher priority)
    const sortedRules = [...OUTLOOK_RULES]
        .filter(r => r.enabled)
        .sort((a, b) => a.priority - b.priority);
    for (const rule of sortedRules) {
        if (evaluateCondition(rule.conditions, sender, subject, isInCC)) {
            matchedRules.push(rule);
            triggers.push(rule.name);
            // Apply actions
            if (rule.actions.markAsRead) {
                markAsRead = true;
            }
            if (rule.actions.moveToFolder && !moveToFolder) {
                moveToFolder = OUTLOOK_FOLDERS[rule.actions.moveToFolder];
            }
            if (rule.actions.forwardTo) {
                for (const recipient of rule.actions.forwardTo) {
                    const email = FORWARD_RECIPIENTS[recipient];
                    if (!forwardTo.includes(email)) {
                        forwardTo.push(email);
                    }
                }
            }
            // Stop processing if rule says so
            if (rule.stopProcessing !== false) {
                break;
            }
        }
    }
    return {
        matched: matchedRules.length > 0,
        matchedRules,
        actions: {
            markAsRead,
            moveToFolder,
            forwardTo,
        },
        triggers,
    };
}
/**
 * Get all enabled rules sorted by priority
 */
export function getEnabledRules(): OutlookRule[] {
    return [...OUTLOOK_RULES]
        .filter(r => r.enabled)
        .sort((a, b) => a.priority - b.priority);
}
/**
 * Get a summary of all rules for documentation
 */
export function getRulesSummary(): { id: number; name: string; description: string; priority: number; enabled: boolean; }[] {
    return OUTLOOK_RULES.map(r => ({
        id: r.id,
        name: r.name,
        description: r.description,
        priority: r.priority,
        enabled: r.enabled,
    }));
}
