import type { Resource } from "@modelcontextprotocol/sdk/types.js";

type EmailResourceUri =
    | "navia://email/vip-domains"
    | "navia://email/classification-rules"
    | "navia://email/response-templates"
    | "navia://planner/bucket-mapping"
    | "navia://clients/directory";

interface VipDomain {
    domain: string;
    name: string;
    sla: string;
}

interface ClassificationCategory {
    name: string;
    priority: number;
    sla: string;
    bucket: string;
    keywords: string[];
    senders: string[];
}

interface ClassificationDefaultCategory {
    name: string;
    priority: number;
    sla: string;
    bucket: string;
}

interface ClassificationRules {
    categories: ClassificationCategory[];
    defaultCategory: ClassificationDefaultCategory;
}

interface ResponseTemplate {
    subject: string;
    body: string;
}

interface ResponseTemplates {
    integration_error: ResponseTemplate;
    vip_request: ResponseTemplate;
    order_status: ResponseTemplate;
    general: ResponseTemplate;
}

interface BucketConfig {
    id: string;
    categories: string[];
    defaultPriority: number;
}

interface BucketMapping {
    "VIP Follow-up": BucketConfig;
    "Urgent": BucketConfig;
    "Client Communications": BucketConfig;
    "Operations": BucketConfig;
    "Support": BucketConfig;
    "Backlog": BucketConfig;
}

type ClientTier = "Platinum" | "Gold" | "Standard";

interface ClientDirectoryEntry extends VipDomain {
    tier: ClientTier;
    contacts: string[];
    notes: string;
}

interface EmailResourceContent<TUri extends EmailResourceUri> {
    uri: TUri;
    mimeType: string;
    text: string;
}

type EmailResourceReadResult =
    | { contents: EmailResourceContent<"navia://email/vip-domains">[] }
    | { contents: EmailResourceContent<"navia://email/classification-rules">[] }
    | { contents: EmailResourceContent<"navia://email/response-templates">[] }
    | { contents: EmailResourceContent<"navia://planner/bucket-mapping">[] }
    | { contents: EmailResourceContent<"navia://clients/directory">[] };

export const emailResources: Resource[] = [

    {

        uri: "navia://email/vip-domains",

        name: "VIP Client Domains",

        description: "List of VIP client email domains that require priority handling",

        mimeType: "application/json",

    },

    {

        uri: "navia://email/classification-rules",

        name: "Email Classification Rules",

        description: "Rules for classifying incoming emails by category, priority, and SLA",

        mimeType: "application/json",

    },

    {

        uri: "navia://email/response-templates",

        name: "Email Response Templates",

        description: "Pre-approved response templates for common email categories",

        mimeType: "application/json",

    },

    {

        uri: "navia://planner/bucket-mapping",

        name: "Planner Bucket Mapping",

        description: "Mapping of email categories to Planner buckets",

        mimeType: "application/json",

    },

    {

        uri: "navia://clients/directory",

        name: "Client Directory",

        description: "Directory of known clients with their tier, SLA, and contact info",

        mimeType: "application/json",

    },

];

const VIP_DOMAINS: VipDomain[] = [

    { domain: "tilesofezra.com", name: "Tiles of Ezra", sla: "2hr" },

    { domain: "medifab.com", name: "Medifab", sla: "2hr" },

    { domain: "raft.ai", name: "Raft.ai", sla: "4hr" },

    { domain: "wove.com", name: "Wove", sla: "4hr" },

    { domain: "apsystems.com", name: "AP Systems", sla: "4hr" },

    { domain: "lbedesign.com", name: "LBE Design", sla: "1day" },

    { domain: "uskidsgolf.com", name: "US Kids Golf", sla: "1day" },

    { domain: "bdsanimalhealth.com", name: "BDS Animal Health", sla: "1day" },

    { domain: "mateauto.com", name: "Mate Auto", sla: "1day" },

    { domain: "askmate.com", name: "Ask Mate", sla: "1day" },

    { domain: "splash.com", name: "Splash", sla: "1day" },

    { domain: "yoni.care", name: "Yoni", sla: "1day" },

    { domain: "thedevelopment.com.au", name: "The Development", sla: "1day" },

    { domain: "netsuite.com", name: "NetSuite", sla: "4hr" },

];

const CLASSIFICATION_RULES: ClassificationRules = {

    categories: [

        {

            name: "integration_error",

            priority: 1,

            sla: "4hr",

            bucket: "Urgent",

            keywords: ["error", "failed", "exception", "timeout", "401", "403", "500", "integration", "api", "webhook", "sync"],

            senders: ["cargowise", "eadaptor", "api@", "noreply@"],

        },

        {

            name: "vip_request",

            priority: 1,

            sla: "2hr",

            bucket: "VIP Follow-up",

            keywords: [],

            senders: VIP_DOMAINS.map((d: VipDomain) => d.domain),

        },

        {

            name: "escalation",

            priority: 1,

            sla: "1hr",

            bucket: "VIP Follow-up",

            keywords: ["urgent", "asap", "critical", "emergency", "immediately", "escalate", "ceo", "manager"],

            senders: [],

        },

        {

            name: "order_status",

            priority: 5,

            sla: "1day",

            bucket: "Operations",

            keywords: ["order", "shipment", "tracking", "delivery", "eta", "status", "where is"],

            senders: [],

        },

        {

            name: "client_inquiry",

            priority: 5,

            sla: "2day",

            bucket: "Client Communications",

            keywords: ["question", "inquiry", "help", "support", "issue", "problem"],

            senders: [],

        },

    ],

    defaultCategory: {

        name: "general",

        priority: 9,

        sla: "2day",

        bucket: "Backlog",

    },

};

const RESPONSE_TEMPLATES: ResponseTemplates = {

    integration_error: {

        subject: "RE: {ORIGINAL_SUBJECT}",

        body: `Hi,



Thank you for reporting this integration issue. Our technical team has been notified and is investigating.



Issue Reference: {TICKET_ID}

Reported: {TIMESTAMP}

Expected Resolution: Within {SLA}



We will provide an update as soon as we have more information.



If this is blocking critical operations, please reply with "URGENT" in the subject line.



Best regards,

Navia Freight Technical Support`,

    },

    vip_request: {

        subject: "RE: {ORIGINAL_SUBJECT}",

        body: `Hi {CLIENT_NAME},



Thank you for reaching out. As a valued VIP client, your request has been prioritized.



Our team is reviewing your message and will respond within {SLA}.



If you need immediate assistance, please call our priority support line.



Best regards,

Navia Freight VIP Support`,

    },

    order_status: {

        subject: "RE: {ORIGINAL_SUBJECT}",

        body: `Hi,



Thank you for your inquiry about your order/shipment status.



We are checking the current status with our warehouse and logistics team. You can expect an update within {SLA}.



Order/Reference: {ORDER_NUMBER}



Best regards,

Navia Freight Operations`,

    },

    general: {

        subject: "RE: {ORIGINAL_SUBJECT}",

        body: `Hi,



Thank you for contacting Navia Freight. We have received your message and will respond within {SLA}.



Best regards,

Navia Freight Team`,

    },

};

const BUCKET_MAPPING: BucketMapping = {

    "VIP Follow-up": {

        id: process.env.BUCKET_VIP || "MR_zXOcdqE6RIF6Yd6wPjMgAEaxq",

        categories: ["vip_request", "escalation"],

        defaultPriority: 1,

    },

    "Urgent": {

        id: process.env.BUCKET_URGENT || "r1uZj-Zh7E63PTEPYqPmIsgANefB",

        categories: ["integration_error"],

        defaultPriority: 1,

    },

    "Client Communications": {

        id: process.env.BUCKET_CLIENT || "RGDYN7r6oEWXkIA0L01o5sgAKG4S",

        categories: ["client_inquiry"],

        defaultPriority: 5,

    },

    "Operations": {

        id: process.env.BUCKET_OPERATIONS || "3hPTGvtl7kC5QGaKItSyDMgAD15d",

        categories: ["order_status"],

        defaultPriority: 5,

    },

    "Support": {

        id: process.env.BUCKET_SUPPORT || "UBHJXpRD50qAPuqJODCDm8gAFBEB",

        categories: [],

        defaultPriority: 5,

    },

    "Backlog": {

        id: process.env.BUCKET_BACKLOG || "-ve7obEGl0GRm3nQC7ebN8gAPac6",

        categories: ["general"],

        defaultPriority: 9,

    },

};

const CLIENT_DIRECTORY: ClientDirectoryEntry[] = VIP_DOMAINS.map((d: VipDomain) => ({

    ...d,

    tier: d.sla === "2hr" ? "Platinum" : d.sla === "4hr" ? "Gold" : "Standard",

    contacts: [],

    notes: "",

}));

export async function readEmailResource(uri: string): Promise<EmailResourceReadResult> {

    switch (uri) {

        case "navia://email/vip-domains":

            return {

                contents: [

                    {

                        uri,

                        mimeType: "application/json",

                        text: JSON.stringify({ vipDomains: VIP_DOMAINS }, null, 2),

                    },

                ],

            };

        case "navia://email/classification-rules":

            return {

                contents: [

                    {

                        uri,

                        mimeType: "application/json",

                        text: JSON.stringify(CLASSIFICATION_RULES, null, 2),

                    },

                ],

            };

        case "navia://email/response-templates":

            return {

                contents: [

                    {

                        uri,

                        mimeType: "application/json",

                        text: JSON.stringify(RESPONSE_TEMPLATES, null, 2),

                    },

                ],

            };

        case "navia://planner/bucket-mapping":

            return {

                contents: [

                    {

                        uri,

                        mimeType: "application/json",

                        text: JSON.stringify(BUCKET_MAPPING, null, 2),

                    },

                ],

            };

        case "navia://clients/directory":

            return {

                contents: [

                    {

                        uri,

                        mimeType: "application/json",

                        text: JSON.stringify({ clients: CLIENT_DIRECTORY }, null, 2),

                    },

                ],

            };

        default:

            throw new Error(`Unknown resource: ${uri}`);

    }

}
