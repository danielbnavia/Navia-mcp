import * as fs from "fs";
import * as path from "path";
import { DeviceCodeCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { Tool } from "@modelcontextprotocol/sdk/types.js";
// @ts-ignore
import { PSTFile } from "pst-extractor";
import { getOutlookCategories, getCategoryAction } from "../config/category-actions.js";

interface ClassificationLogEntry {
    timestamp: string;
    sender: string;
    subject: string;
    domain: string;
    category: string;
    bucket: string;
    priority: number;
    isVip: boolean;
    confidence: string;
    triggers: string[];
}

export type SentimentLevel = "very_negative" | "negative" | "neutral" | "positive" | "very_positive";

export type UrgencyLevel = "critical" | "high" | "medium" | "low";

export interface SentimentAnalysis {
    sentiment: SentimentLevel;
    sentimentScore: number;
    urgency: UrgencyLevel;
    urgencyScore: number;
    isAutoReply: boolean;
    indicators: {
        negative: string[];
        positive: string[];
        urgency: string[];
    };
}

interface KeywordWeight {
    keyword: string;
    weight: number;
}

interface PstMessage {
    messageClass?: string;
    subject?: string;
    senderEmailAddress?: string;
    senderName?: string;
    messageDeliveryTime?: Date;
    body?: string;
    importance?: number;
}

interface PstFolder {
    displayName?: string;
    hasSubfolders?: boolean;
    contentCount?: number;
    getSubFolders(): PstFolder[];
    getNextChild(): PstMessage | null;
}

interface PstFileInstance {
    getRootFolder(): PstFolder;
    close(): void;
}

interface GraphMessage {
    id: string;
    subject?: string;
    from?: {
        emailAddress?: {
            address?: string;
            name?: string;
        };
    };
    receivedDateTime: string;
    isRead: boolean;
    importance?: string;
    bodyPreview?: string;
}

interface GraphMessagesResponse {
    value: GraphMessage[];
}
// Store user token for device code auth
let userGraphClient: Client | null = null;
let userEmail: string | null = null;
// PST file path
const PST_PATH = path.join("M:", "powerpages", "Emails", "Danielb@naviafreight.com.pst");
// In-memory classification log (last 500 entries)
const MAX_LOG_ENTRIES = 500;
const classificationLog: ClassificationLogEntry[] = [];
// Log file path
const LOG_FILE = path.join(process.cwd(), "classification-log.json");
// Load existing log on startup
try {
    if (fs.existsSync(LOG_FILE)) {
        const data = fs.readFileSync(LOG_FILE, "utf-8");
        const entries = JSON.parse(data);
        classificationLog.push(...entries.slice(-MAX_LOG_ENTRIES));
    }
}
catch (e) {
    console.log("Starting with empty classification log");
}
// Save log to file periodically
function saveLog(): void {
    try {
        fs.writeFileSync(LOG_FILE, JSON.stringify(classificationLog, null, 2));
    }
    catch (e) {
        console.error("Failed to save classification log:", e);
    }
}
// Export log access functions
export function getClassificationLog(limit: number = 100): ClassificationLogEntry[] {
    return classificationLog.slice(-limit);
}
export function getClassificationStats(): {
    total: number;
    byCategory: Record<string, number>;
    byBucket: Record<string, number>;
    byDomain: Record<string, number>;
    triggerCounts: Record<string, number>;
    potentialFalsePositives: ClassificationLogEntry[];
} {
    const stats: {
        total: number;
        byCategory: Record<string, number>;
        byBucket: Record<string, number>;
        byDomain: Record<string, number>;
        triggerCounts: Record<string, number>;
        potentialFalsePositives: ClassificationLogEntry[];
    } = {
        total: classificationLog.length,
        byCategory: {},
        byBucket: {},
        byDomain: {},
        triggerCounts: {},
        potentialFalsePositives: [],
    };
    for (const entry of classificationLog) {
        // Count by category
        stats.byCategory[entry.category] = (stats.byCategory[entry.category] || 0) + 1;
        // Count by bucket
        stats.byBucket[entry.bucket] = (stats.byBucket[entry.bucket] || 0) + 1;
        // Count by domain
        stats.byDomain[entry.domain] = (stats.byDomain[entry.domain] || 0) + 1;
        // Count triggers
        for (const trigger of entry.triggers) {
            stats.triggerCounts[trigger] = (stats.triggerCounts[trigger] || 0) + 1;
        }
        // Flag potential false positives: low confidence integration errors
        if (entry.category === "integration_error" && entry.confidence === "low") {
            stats.potentialFalsePositives.push(entry);
        }
        // Also flag medium confidence with only generic triggers
        if (entry.category === "integration_error" && entry.confidence === "medium" &&
            entry.triggers.every(t => t.startsWith("integration:error") || t.startsWith("integration:failed"))) {
            stats.potentialFalsePositives.push(entry);
        }
    }
    return stats;
}
export function clearClassificationLog(): void {
    classificationLog.length = 0;
    saveLog();
}
export const emailTools: Tool[] = [
    {
        name: "email_classify",
        description: "Classify an email based on sender, subject, and content. Returns category, priority, SLA, and recommended bucket.",
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
                    description: "Email body content (can be truncated)",
                },
            },
            required: ["sender", "subject"],
        },
    },
    {
        name: "email_extract_entities",
        description: "Extract key entities from email content: order numbers, tracking numbers, client names, dates, amounts",
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
        name: "email_get_client_info",
        description: "Get information about a client based on their email domain",
        inputSchema: {
            type: "object",
            properties: {
                domain: {
                    type: "string",
                    description: "Email domain (e.g., tilesofezra.com)",
                },
            },
            required: ["domain"],
        },
    },
    {
        name: "email_suggest_response",
        description: "Suggest an email response based on the classified email. Takes the original email details and classification category, and returns a suggested reply.",
        inputSchema: {
            type: "object",
            properties: {
                subject: {
                    type: "string",
                    description: "The email subject line from the original email",
                },
                body: {
                    type: "string",
                    description: "The full email body content to generate a response for",
                },
                from: {
                    type: "string",
                    description: "The sender's email address (e.g., sender@example.com)",
                },
                category: {
                    type: "string",
                    description: "The classification category from ClassifyEmail result (e.g., P1.Urgent, P2.Raft, General, integration_error, escalation)",
                },
                tone: {
                    type: "string",
                    description: "The desired tone for the response (e.g., professional, urgent, empathetic, concise)",
                },
            },
            required: ["subject", "body", "from", "category"],
        },
    },
    {
        name: "email_fetch_recent",
        description: "Fetch recent emails from Office 365 mailbox using Microsoft Graph API. Run email_login first to authenticate.",
        inputSchema: {
            type: "object",
            properties: {
                count: {
                    type: "number",
                    description: "Number of emails to fetch (default: 10, max: 50)",
                },
                folder: {
                    type: "string",
                    description: "Folder to read from (default: inbox). Options: inbox, sentitems, drafts",
                },
                unreadOnly: {
                    type: "boolean",
                    description: "Only fetch unread emails (default: false)",
                },
            },
            required: [],
        },
    },
    {
        name: "email_login",
        description: "Authenticate with Office 365 using device code flow. You'll get a code to enter at microsoft.com/devicelogin. No admin consent required.",
        inputSchema: {
            type: "object",
            properties: {},
            required: [],
        },
    },
    {
        name: "email_read_pst",
        description: "Read emails from local Outlook PST file. Returns recent emails from a specified folder.",
        inputSchema: {
            type: "object",
            properties: {
                folder: {
                    type: "string",
                    description: "Folder name to read from (e.g., 'read', 'RAFT emails', 'Warehouse'). Use 'folders' to list all.",
                },
                count: {
                    type: "number",
                    description: "Number of emails to fetch (default: 20, max: 100)",
                },
            },
            required: [],
        },
    },
    {
        name: "email_analyze_sentiment",
        description: "Analyze the sentiment and urgency of an email. Returns sentiment level (very_negative to very_positive), urgency level (critical/high/medium/low), auto-reply detection, and keyword indicators.",
        inputSchema: {
            type: "object",
            properties: {
                subject: {
                    type: "string",
                    description: "Email subject line",
                },
                body: {
                    type: "string",
                    description: "Email body content",
                },
                sender: {
                    type: "string",
                    description: "Sender email address (used for auto-reply detection)",
                },
            },
            required: ["subject", "body"],
        },
    },
    {
        name: "email_summarize_thread",
        description: "Summarize an email thread by extracting key points, action items, and participants from multiple email bodies. Provide the emails in chronological order.",
        inputSchema: {
            type: "object",
            properties: {
                emails: {
                    type: "array",
                    items: {
                        type: "object",
                        properties: {
                            from: { type: "string", description: "Sender email or name" },
                            subject: { type: "string", description: "Email subject" },
                            body: { type: "string", description: "Email body content" },
                            date: { type: "string", description: "Date/time of the email (ISO 8601)" },
                        },
                        required: ["from", "body"],
                    },
                    description: "Array of emails in the thread, in chronological order",
                },
            },
            required: ["emails"],
        },
    },
    {
        name: "email_bulk_triage",
        description: "Classify multiple emails at once. Accepts an array of emails and returns classification results for each. More efficient than calling email_classify repeatedly.",
        inputSchema: {
            type: "object",
            properties: {
                emails: {
                    type: "array",
                    items: {
                        type: "object",
                        properties: {
                            id: { type: "string", description: "Email ID for tracking results" },
                            sender: { type: "string", description: "Sender email address" },
                            subject: { type: "string", description: "Email subject line" },
                            body: { type: "string", description: "Email body (can be truncated)" },
                        },
                        required: ["sender", "subject"],
                    },
                    description: "Array of emails to classify",
                },
            },
            required: ["emails"],
        },
    },
    {
        name: "email_score_priority",
        description: "Calculate a composite priority score (0-100) for an email based on sender VIP status, sentiment, urgency, category, and SLA. Higher score = higher priority. Use this to rank emails for processing order.",
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
            },
            required: ["sender", "subject"],
        },
    },
];
// VIP client domains - use suffix matching for accuracy
const VIP_DOMAINS: string[] = [
    // Existing VIPs
    "tilesofezra.com", "tilesofezra.com.au",
    "medifab.com", "medifab.co.nz",
    "apsystems.com", "apsystems.com.au",
    "raft.ai",
    "vectorai.atlassian.net", // Raft AI Jira support tickets
    "wove.com",
    "lbedesign.com",
    "uskidsgolf.com",
    "bdsanimalhealth.com", "bdsanimalhealth.com.au",
    "mateauto.com", "mateauto.com.au",
    "askmate.com",
    "thedevelopment.com.au",
    "netsuite.com",
    "braumach.com.au",
    "cin7.com",
    // New VIPs from config update (2026-01-27)
    "splashblanket.com", "splash.com", // Splash Blanket
    "yonipleasurepalace.com", "yoni.care", // Yoni Pleasure Palace
    "bfglobalwarehouse.com", "bfglobal.com", // BF Global Warehouse
];
// Internal VIP senders (escalation contacts) - full email addresses only
const INTERNAL_VIP_SENDERS: string[] = [
    "brendan@naviafreight.com",
    "chris@naviafreight.com",
    "bruno@naviafreight.com",
];
// Integration error keywords with weights (higher = more confident)
const INTEGRATION_KEYWORDS: KeywordWeight[] = [
    // High confidence - technical terms
    { keyword: "api error", weight: 3 },
    { keyword: "api failure", weight: 3 },
    { keyword: "api integration failure", weight: 3 },
    { keyword: "api integration", weight: 2 },
    { keyword: "integration failure", weight: 3 },
    { keyword: "webhook failed", weight: 3 },
    { keyword: "sync failed", weight: 3 },
    { keyword: "sync error", weight: 3 },
    { keyword: "connection failed", weight: 3 },
    { keyword: "integration failed", weight: 3 },
    { keyword: "integration error", weight: 3 },
    { keyword: "batch processing failed", weight: 3 },
    // High confidence - Raft AI / JIRA ticket patterns
    { keyword: "automation fails", weight: 3 },
    { keyword: "automation failure", weight: 3 },
    { keyword: "pack processing", weight: 3 },
    { keyword: "pack upload", weight: 3 },
    { keyword: "pushing the pack", weight: 3 },
    { keyword: "extraction issue", weight: 3 },
    { keyword: "extraction failed", weight: 3 },
    { keyword: "incorrect extraction", weight: 3 },
    { keyword: "header field", weight: 2 },
    { keyword: "document classification", weight: 2 },
    { keyword: "invalid file", weight: 3 },
    // High confidence - freight/logistics specific
    { keyword: "edi rejected", weight: 3 },
    { keyword: "edi failure", weight: 3 },
    { keyword: "message rejected", weight: 3 },
    { keyword: "data import failed", weight: 3 },
    { keyword: "export failed", weight: 3 },
    { keyword: "integration halted", weight: 3 },
    // Medium confidence - need context
    { keyword: "eadaptor", weight: 2 }, // Reduced from 3 - sends success notifications too
    { keyword: "cargowise", weight: 2 },
    { keyword: "edi message", weight: 2 }, // Changed from "edi" - too short, matches "immediate"
    { keyword: "edi error", weight: 2 },
    { keyword: "edi issue", weight: 2 },
    { keyword: "magaya", weight: 2 },
    { keyword: "wisetech", weight: 2 },
    { keyword: "customs rejection", weight: 2 },
    { keyword: "shipment exception", weight: 2 },
    { keyword: "consignment error", weight: 2 },
    { keyword: "timeout", weight: 2 },
    // Low confidence - generic terms (need multiple matches)
    { keyword: "exception", weight: 1 }, // Reduced from 2 - too common in general context
    { keyword: "records affected", weight: 1 }, // Reduced from 2 - neutral term
    { keyword: "failed", weight: 1 },
    { keyword: "fails", weight: 1 },
    { keyword: "error", weight: 1 },
    { keyword: "errors", weight: 1 },
    { keyword: "issue with", weight: 2 },
    { keyword: "issue when", weight: 2 },
    { keyword: "issue report", weight: 2 },
    { keyword: "raft issues", weight: 3 },
    { keyword: "transmission issue", weight: 2 },
    { keyword: "transmission issues", weight: 2 },
    { keyword: "transmission error", weight: 2 },
    { keyword: "transmission fail", weight: 2 },
    { keyword: "sync issue", weight: 2 },
    { keyword: "sync fail", weight: 2 },
    // Standalone platform/tool names (for subject-only classification)
    { keyword: "integration", weight: 2 },
    { keyword: "odoo", weight: 2 },
    { keyword: "cin7", weight: 2 },
    { keyword: "shopify", weight: 2 },
    { keyword: "netsuite", weight: 2 },
    { keyword: "machship", weight: 2 },
    { keyword: "naviafill", weight: 2 },
    { keyword: "3pl", weight: 2 },
    { keyword: "codecom", weight: 2 },
    { keyword: "json", weight: 2 },
    { keyword: "xml", weight: 2 },
];
// Exclusion patterns - if found, reduce integration error confidence
const INTEGRATION_EXCLUSIONS: string[] = [
    // Success indicators
    "no error", "no errors", "without error", "error free", "error-free",
    "resolved", "fixed", "corrected", "completed successfully",
    "0 errors", "zero errors", "passed", "successful",
    // Report/summary context (email is ABOUT errors, not reporting a current error)
    "error report", "error summary", "daily error", "weekly error",
    "errors were resolved", "errors have been fixed", "no new errors",
    "all errors addressed", "error monitoring", "error tracking tool",
    "reduce errors", "error handling improved", "previous errors",
];
// HTTP error codes need to be checked with context
const HTTP_ERROR_CODES: string[] = ["401", "403", "404", "500", "502", "503", "504"];
// Escalation keywords with weights
const ESCALATION_KEYWORDS: KeywordWeight[] = [
    { keyword: "escalate", weight: 3 },
    { keyword: "complaint", weight: 3 },
    { keyword: "unacceptable", weight: 3 },
    { keyword: "speak to manager", weight: 3 },
    { keyword: "legal action", weight: 3 },
    { keyword: "customer service ticket", weight: 2 },
    { keyword: "ticket raised", weight: 2 },
    { keyword: "cst0000", weight: 2 }, // Customer Service Ticket ID pattern
    { keyword: "unhappy", weight: 2 },
    { keyword: "dissatisfied", weight: 2 },
    { keyword: "frustrated", weight: 2 },
    { keyword: "disappointed", weight: 2 },
    { keyword: "resolution required", weight: 2 },
    { keyword: "immediate attention", weight: 2 },
    // Repeated contact indicators
    { keyword: "third time", weight: 3 },
    { keyword: "multiple times", weight: 3 },
    { keyword: "again and again", weight: 3 },
    { keyword: "repeatedly", weight: 2 },
    { keyword: "several times", weight: 2 },
    { keyword: "contacted you before", weight: 2 },
    { keyword: "still waiting", weight: 2 },
    { keyword: "no response", weight: 2 },
    { keyword: "been waiting", weight: 2 },
    { keyword: "follow up", weight: 1 },
    { keyword: "following up", weight: 1 },
];
// Urgent keywords with weights
const URGENT_KEYWORDS: KeywordWeight[] = [
    { keyword: "urgent", weight: 2 },
    { keyword: "asap", weight: 2 },
    { keyword: "critical", weight: 2 },
    { keyword: "emergency", weight: 3 },
    { keyword: "immediately", weight: 2 },
    { keyword: "deadline today", weight: 3 },
    { keyword: "blocking", weight: 2 },
    { keyword: "stuck", weight: 1 },
    { keyword: "priority", weight: 1 },
    { keyword: "time sensitive", weight: 2 },
];
// Exclusion patterns for order/shipment false positives
const ORDER_EXCLUSIONS: string[] = [
    "in order to", "order of magnitude", "out of order",
    "tracking pixel", "tracking code", "tracking analytics",
    "shipment of features", "shipment of updates",
];
// Order/shipment status keywords with weights (higher = more confident)
const ORDER_STATUS_KEYWORDS: KeywordWeight[] = [
    // High confidence - explicit status queries
    { keyword: "shipment delayed", weight: 3 },
    { keyword: "delivery status", weight: 3 },
    { keyword: "tracking number", weight: 3 },
    { keyword: "where is my order", weight: 3 },
    { keyword: "order status", weight: 3 },
    { keyword: "shipping update", weight: 3 },
    { keyword: "shipment status", weight: 3 },
    { keyword: "delivery update", weight: 3 },
    { keyword: "freight status", weight: 3 },
    { keyword: "cargo status", weight: 3 },
    { keyword: "shipment update", weight: 3 },
    // Medium confidence - logistics-specific terms
    { keyword: "container", weight: 2 },
    { keyword: "stuck at port", weight: 2 },
    { keyword: "port delay", weight: 2 },
    { keyword: "customs hold", weight: 2 },
    { keyword: "customs clearance", weight: 2 },
    { keyword: "revised eta", weight: 2 },
    { keyword: "eta change", weight: 2 },
    { keyword: "estimated arrival", weight: 2 },
    { keyword: "arrival date", weight: 2 },
    { keyword: "delivery date", weight: 2 },
    { keyword: "expected delivery", weight: 2 },
    { keyword: "in transit", weight: 2 },
    { keyword: "vessel", weight: 2 },
    { keyword: "bill of lading", weight: 2 },
    { keyword: "consignment", weight: 2 },
    // Low confidence - generic terms (need multiple matches)
    { keyword: "shipment", weight: 1 },
    { keyword: "tracking", weight: 1 },
    { keyword: "delivery", weight: 1 },
    { keyword: "order", weight: 1 },
    { keyword: "delayed", weight: 1 },
    { keyword: "arrived", weight: 1 },
    { keyword: "departed", weight: 1 },
];
// ============================================================================
// SENTIMENT ANALYSIS
// ============================================================================
// Negative sentiment keywords (frustration, anger, disappointment)
const NEGATIVE_SENTIMENT: KeywordWeight[] = [
    // High negativity
    { keyword: "unacceptable", weight: 3 },
    { keyword: "terrible", weight: 3 },
    { keyword: "awful", weight: 3 },
    { keyword: "disgusted", weight: 3 },
    { keyword: "furious", weight: 3 },
    { keyword: "outraged", weight: 3 },
    { keyword: "appalling", weight: 3 },
    { keyword: "inexcusable", weight: 3 },
    // Medium negativity
    { keyword: "frustrated", weight: 2 },
    { keyword: "disappointed", weight: 2 },
    { keyword: "unhappy", weight: 2 },
    { keyword: "dissatisfied", weight: 2 },
    { keyword: "annoyed", weight: 2 },
    { keyword: "concerned", weight: 2 },
    { keyword: "worried", weight: 2 },
    { keyword: "upset", weight: 2 },
    { keyword: "confused", weight: 2 },
    { keyword: "ridiculous", weight: 2 },
    { keyword: "unbelievable", weight: 2 },
    // Low negativity
    { keyword: "not happy", weight: 1 },
    { keyword: "not satisfied", weight: 1 },
    { keyword: "still waiting", weight: 1 },
    { keyword: "no response", weight: 1 },
    { keyword: "ignored", weight: 1 },
    { keyword: "again", weight: 1 }, // "this happened again"
    { keyword: "still not", weight: 1 },
    { keyword: "yet another", weight: 1 },
];
// Positive sentiment keywords
const POSITIVE_SENTIMENT: KeywordWeight[] = [
    { keyword: "thank you", weight: 2 },
    { keyword: "thanks", weight: 1 },
    { keyword: "appreciate", weight: 2 },
    { keyword: "grateful", weight: 2 },
    { keyword: "great job", weight: 2 },
    { keyword: "well done", weight: 2 },
    { keyword: "excellent", weight: 2 },
    { keyword: "perfect", weight: 2 },
    { keyword: "amazing", weight: 2 },
    { keyword: "fantastic", weight: 2 },
    { keyword: "pleased", weight: 1 },
    { keyword: "happy with", weight: 1 },
    { keyword: "satisfied", weight: 1 },
    { keyword: "good work", weight: 1 },
];
// Auto-reply patterns in content (reduce priority)
const AUTO_REPLY_CONTENT_PATTERNS: string[] = [
    "automatic reply",
    "auto-reply",
    "autoreply",
    "out of office",
    "ooo:",
    "i am currently out",
    "i'm currently out",
    "away from the office",
    "limited access to email",
    "on annual leave",
    "on vacation",
    "on holiday",
    "will respond when i return",
    "i will be out of the office",
    "undeliverable",
    "delivery status notification",
    "delivery failure",
    "message not delivered",
];
// Auto-reply sender patterns (check sender address)
const AUTO_REPLY_SENDER_PATTERNS: string[] = [
    "mailer-daemon",
    "postmaster",
    "noreply",
    "no-reply",
    "no_reply",
    "donotreply",
    "do-not-reply",
    "do_not_reply",
    "automated",
    "autoresponder",
    "bounce",
    "notifications@",
    "notification@",
    "alert@",
    "alerts@",
];
// Urgency indicators (beyond just keywords)
const HIGH_URGENCY_PATTERNS: RegExp[] = [
    /\bASAP\b/i,
    /\bURGENT\b/i,
    /!!+/, // Multiple exclamation marks
    /\bTODAY\b/i,
    /\bNOW\b/i,
    /\bIMMEDIATELY\b/i,
    /\bCRITICAL\b/i,
    /\bBLOCKING\b/i,
    /\bBLOCKED\b/i,
    /time.?sensitive/i,
    /deadline/i,
];
/**
 * Analyze sentiment and urgency from email content
 */
function analyzeSentiment(subject: string, body: string, sender: string = ""): SentimentAnalysis {
    const content = `${subject} ${body}`.toLowerCase();
    const originalContent = `${subject} ${body}`; // Keep original case for pattern matching
    const senderLower = sender.toLowerCase();
    const indicators: {
        negative: string[];
        positive: string[];
        urgency: string[];
    } = {
        negative: [],
        positive: [],
        urgency: [],
    };
    // Calculate negative sentiment
    let negativeScore = 0;
    for (const { keyword, weight } of NEGATIVE_SENTIMENT) {
        if (content.includes(keyword)) {
            negativeScore += weight;
            indicators.negative.push(keyword);
        }
    }
    // Calculate positive sentiment
    let positiveScore = 0;
    for (const { keyword, weight } of POSITIVE_SENTIMENT) {
        if (content.includes(keyword)) {
            positiveScore += weight;
            indicators.positive.push(keyword);
        }
    }
    // Check for caps lock abuse (more than 3 consecutive caps words)
    const capsMatch = originalContent.match(/\b[A-Z]{3,}\b/g);
    if (capsMatch && capsMatch.length >= 3) {
        negativeScore += 2;
        indicators.negative.push("CAPS_LOCK");
    }
    // Check for multiple punctuation
    if (/[!?]{2,}/.test(originalContent)) {
        negativeScore += 1;
        indicators.negative.push("excessive_punctuation");
    }
    // Calculate urgency
    let urgencyScore = 0;
    for (const pattern of HIGH_URGENCY_PATTERNS) {
        if (pattern.test(originalContent)) {
            urgencyScore += 2;
            indicators.urgency.push(pattern.toString());
        }
    }
    // Check for auto-reply (by content or sender patterns)
    const isAutoReplyContent = AUTO_REPLY_CONTENT_PATTERNS.some(p => content.includes(p));
    const isAutoReplySender = AUTO_REPLY_SENDER_PATTERNS.some(p => senderLower.includes(p));
    const isAutoReply = isAutoReplyContent || isAutoReplySender;
    // Calculate final sentiment score (-10 to +10)
    const sentimentScore = Math.max(-10, Math.min(10, positiveScore - negativeScore));
    // Determine sentiment level
    let sentiment: SentimentLevel;
    if (sentimentScore <= -5)
        sentiment = "very_negative";
    else if (sentimentScore < 0)
        sentiment = "negative";
    else if (sentimentScore === 0)
        sentiment = "neutral";
    else if (sentimentScore <= 3)
        sentiment = "positive";
    else
        sentiment = "very_positive";
    // Determine urgency level
    let urgency: UrgencyLevel;
    urgencyScore = Math.min(10, urgencyScore);
    if (urgencyScore >= 6)
        urgency = "critical";
    else if (urgencyScore >= 4)
        urgency = "high";
    else if (urgencyScore >= 2)
        urgency = "medium";
    else
        urgency = "low";
    return {
        sentiment,
        sentimentScore,
        urgency,
        urgencyScore,
        isAutoReply,
        indicators,
    };
}
// Exclusion patterns for inquiry false positives
const INQUIRY_EXCLUSIONS: string[] = [
    "frequently asked question", "help center", "help desk",
    "helpful tips", "helpful resource", "helpful guide",
];
// Confidence thresholds
const CONFIDENCE_THRESHOLD: { HIGH: number; MEDIUM: number; LOW: number } = {
    HIGH: 3, // Confident classification
    MEDIUM: 2, // Likely correct
    LOW: 1, // Needs review
};
// Helper: Check if domain matches (exact suffix match)
function matchesDomain(senderDomain: string, vipDomain: string): boolean {
    if (!senderDomain || !vipDomain)
        return false;
    // Exact match or subdomain match (e.g., mail.tilesofezra.com matches tilesofezra.com)
    return senderDomain === vipDomain || senderDomain.endsWith(`.${vipDomain}`);
}
// Helper: Calculate weighted keyword score
function calculateKeywordScore(content: string, keywords: KeywordWeight[]): { score: number; matched: string[] } {
    let score = 0;
    const matched: string[] = [];
    for (const { keyword, weight } of keywords) {
        if (content.includes(keyword)) {
            score += weight;
            matched.push(keyword);
        }
    }
    return { score, matched };
}
// Helper: Check for exclusion patterns
function hasExclusionPattern(content: string, exclusions: string[]): boolean {
    return exclusions.some(pattern => content.includes(pattern));
}
// Client info database
const CLIENT_INFO: Record<string, { name: string; tier: string; sla: string; notes: string }> = {
    "tilesofezra.com": { name: "Tiles of Ezra", tier: "VIP", sla: "2hr", notes: "Premium tile supplier" },
    "medifab.com": { name: "Medifab", tier: "VIP", sla: "2hr", notes: "Medical equipment" },
    "raft.ai": { name: "Raft.ai", tier: "VIP", sla: "4hr", notes: "AI/Vector platform - ticket system via Jira" },
    "wove.com": { name: "Wove", tier: "VIP", sla: "4hr", notes: "Integration partner" },
    "apsystems.com": { name: "AP Systems", tier: "Standard", sla: "1day", notes: "Solar equipment" },
    "lbedesign.com": { name: "LBE Design", tier: "Standard", sla: "1day", notes: "Design products" },
    "uskidsgolf.com": { name: "US Kids Golf", tier: "Standard", sla: "1day", notes: "Golf equipment" },
    // New VIPs (2026-01-27)
    "splashblanket.com": { name: "Splash Blanket", tier: "VIP", sla: "2hr", notes: "VIP client - priority handling" },
    "splash.com": { name: "Splash Blanket", tier: "VIP", sla: "2hr", notes: "VIP client - priority handling" },
    "yonipleasurepalace.com": { name: "Yoni Pleasure Palace", tier: "VIP", sla: "2hr", notes: "VIP client - priority handling" },
    "yoni.care": { name: "Yoni Pleasure Palace", tier: "VIP", sla: "2hr", notes: "VIP client - priority handling" },
    "bfglobalwarehouse.com": { name: "BF Global Warehouse", tier: "VIP", sla: "2hr", notes: "VIP client - warehouse operations" },
    "bfglobal.com": { name: "BF Global Warehouse", tier: "VIP", sla: "2hr", notes: "VIP client - warehouse operations" },
    "bdsanimalhealth.com": { name: "BDS Animal Health", tier: "VIP", sla: "2hr", notes: "VIP client - animal health products" },
    // 3PL Warehouse partners
    "dglgroup.com": { name: "DGL Group", tier: "Partner", sla: "4hr", notes: "3PL warehouse - Maddington facility" },
    "vectorai.atlassian.net": { name: "Raft AI Support", tier: "VIP", sla: "4hr", notes: "Raft AI Jira tickets - integration issues" },
};
export async function handleEmailTool(name: string, args: Record<string, unknown>): Promise<{ content: { type: string; text: string }[] }> {
    switch (name) {
        case "email_classify": {
            const sender = ((args.sender as string) || "").toLowerCase();
            const subject = ((args.subject as string) || "").toLowerCase();
            const body = ((args.body as string) || "").toLowerCase();
            const content = `${subject} ${body}`;
            // Extract domain from sender
            const domain = sender.split("@")[1] || "";
            // Track which rules triggered the classification
            const triggers: string[] = [];
            // Determine classification
            let category = "general";
            let priority = 5; // Medium
            let sla = "2day";
            let bucket = "Backlog";
            let isVip = false;
            let isInternalVip = false;
            let confidence = "low";
            // Check for internal VIP senders (highest priority)
            // Matches exact email OR any @naviafreight.com address
            const matchedInternalVip = INTERNAL_VIP_SENDERS.find((v) => sender === v);
            const isNaviaStaff = domain === "naviafreight.com";
            if (matchedInternalVip || isNaviaStaff) {
                isInternalVip = true;
                isVip = true;
                category = "internal_vip"; // Set category for internal staff
                priority = 1;
                sla = "1hr";
                bucket = "Urgent";
                confidence = "high";
                triggers.push(matchedInternalVip ? `internal_vip:${matchedInternalVip}` : `internal_vip:@naviafreight.com`);
            }
            // Check for VIP client domains (using proper suffix matching)
            const matchedVipDomain = VIP_DOMAINS.find((d) => matchesDomain(domain, d));
            if (matchedVipDomain) {
                isVip = true;
                priority = 1;
                sla = "2hr";
                bucket = "VIP Follow-up";
                category = "vip";
                confidence = "high";
                triggers.push(`vip_domain:${matchedVipDomain}`);
            }
            // Check for VIP mentions in content (when not already VIP by domain)
            if (!isVip && /\bvip\b|vip\s*client|important\s*client|priority\s*client|key\s*account/i.test(content)) {
                isVip = true;
                priority = 1;
                sla = "2hr";
                bucket = "VIP Follow-up";
                category = "vip";
                confidence = "medium"; // Medium since it's content-based not domain-based
                triggers.push("vip_content:mentioned");
            }
            // Check for escalation keywords with weighted scoring
            const escalationResult = calculateKeywordScore(content, ESCALATION_KEYWORDS);
            if (escalationResult.score >= CONFIDENCE_THRESHOLD.MEDIUM) {
                category = "escalation";
                priority = 1;
                sla = isVip ? sla : "2hr";
                bucket = "Urgent";
                confidence = escalationResult.score >= CONFIDENCE_THRESHOLD.HIGH ? "high" : "medium";
                triggers.push(...escalationResult.matched.map(k => `escalation:${k}`));
            }
            // Check for integration errors with weighted scoring and exclusions
            const integrationResult = calculateKeywordScore(content, INTEGRATION_KEYWORDS);
            const hasExclusion = hasExclusionPattern(content, INTEGRATION_EXCLUSIONS);
            // HTTP codes need error context (e.g., "error 500" or "500 error" not "500 sqm")
            const matchedHttpErrors = HTTP_ERROR_CODES.filter((code) => content.includes(`error ${code}`) ||
                content.includes(`${code} error`) ||
                content.includes(`code ${code}`) ||
                content.includes(`status ${code}`) ||
                content.includes(`http ${code}`));
            // Add HTTP error weight
            const httpErrorScore = matchedHttpErrors.length * 2;
            const totalIntegrationScore = integrationResult.score + httpErrorScore;
            // Reduce confidence if exclusion patterns found
            const adjustedIntegrationScore = hasExclusion
                ? Math.max(0, totalIntegrationScore - 2)
                : totalIntegrationScore;
            // Only classify as integration error if score meets threshold
            const hasIntegrationError = adjustedIntegrationScore >= CONFIDENCE_THRESHOLD.MEDIUM;
            if (hasIntegrationError) {
                triggers.push(...integrationResult.matched.map(k => `integration:${k}`));
                triggers.push(...matchedHttpErrors.map(c => `http_error:${c}`));
                if (hasExclusion)
                    triggers.push("exclusion_found");
            }
            if (hasIntegrationError && category !== "escalation") {
                const integrationConfidence = adjustedIntegrationScore >= CONFIDENCE_THRESHOLD.HIGH ? "high" : "medium";
                if (!isVip) {
                    category = "integration_error";
                    priority = 1;
                    sla = "4hr";
                    bucket = "Urgent";
                    confidence = integrationConfidence;
                }
                else {
                    category = "vip_integration_error";
                    priority = 1;
                    sla = "2hr";
                    bucket = "Urgent";
                    confidence = integrationConfidence;
                }
            }
            // Check for urgent keywords with weighted scoring
            const urgentResult = calculateKeywordScore(content, URGENT_KEYWORDS);
            if (urgentResult.score >= CONFIDENCE_THRESHOLD.LOW) {
                priority = 1;
                if (!isVip)
                    sla = "2hr";
                if (bucket === "Backlog")
                    bucket = "Urgent";
                triggers.push(...urgentResult.matched.map(k => `urgent:${k}`));
                // If still general category but urgent, treat as client inquiry requiring action
                if (category === "general" && urgentResult.score >= CONFIDENCE_THRESHOLD.MEDIUM) {
                    category = "client_inquiry";
                    bucket = "Urgent";
                    confidence = "medium";
                    triggers.push("urgent_inquiry:detected");
                }
            }
            // Check for order/shipment related with weighted scoring
            const hasOrderExclusion = hasExclusionPattern(content, ORDER_EXCLUSIONS);
            if (!hasOrderExclusion) {
                const orderResult = calculateKeywordScore(content, ORDER_STATUS_KEYWORDS);
                if (orderResult.score >= CONFIDENCE_THRESHOLD.LOW && category === "general") {
                    category = "order_status";
                    bucket = "Operations";
                    sla = "1day";
                    confidence = orderResult.score >= CONFIDENCE_THRESHOLD.HIGH ? "high" : "medium";
                    triggers.push(...orderResult.matched.map(k => `order:${k}`));
                }
            }
            // Check for client inquiry (but not if it's clearly an integration/technical issue or warehouse related)
            const hasInquiryExclusion = hasExclusionPattern(content, INQUIRY_EXCLUSIONS);
            const isTechnicalContext = /api|integration|sync|error|failure|system|server|webhook|connection/i.test(content);
            const isWarehouseContext = /warehouse|inventory|stock|sku|putaway|goods handling/i.test(content);
            if (!hasInquiryExclusion && !isTechnicalContext && !isWarehouseContext && (content.includes("question") || content.includes("inquiry") || content.includes("help"))) {
                if (category === "general") {
                    category = "client_inquiry";
                    bucket = "Client Communications";
                    sla = "1day";
                    confidence = "medium";
                    if (content.includes("question"))
                        triggers.push("inquiry:question");
                    if (content.includes("inquiry"))
                        triggers.push("inquiry:inquiry");
                    if (content.includes("help"))
                        triggers.push("inquiry:help");
                }
            }
            // Check for AP Posting / Accounts Payable (P1)
            if (/unable to post|atp trigger|ap posting|posting failure|payment reminder|invoice due|payment due|accounts payable|ap\s*post|overdue invoice|outstanding payment|remittance/i.test(content)) {
                category = "ap_posting";
                priority = 1;
                sla = "2hr";
                bucket = "P1.AP-Posting-Failure";
                confidence = "high";
                triggers.push("ap_posting:detected");
            }
            // Check for RAFT tickets/support
            if (/raft[-\s]?\d+|raft.*ticket|ticket.*raft|raft\s*support|raft\s*issue|#raft/i.test(content) ||
                (domain.includes("raft") && /ticket|support|issue/i.test(content))) {
                category = "raft_ticket";
                priority = 2;
                sla = "4hr";
                bucket = "P2.Raft-Support";
                confidence = "high";
                triggers.push("raft_ticket:detected");
            }
            // Check for Raft/CW1 sync issues (P2)
            if (/pushed but not posting|awaiting response from cargowise|raft.*not.*posting|cw1.*sync/i.test(content)) {
                if (category !== "ap_posting") {
                    category = "raft_cw1";
                    priority = 2;
                    sla = "4hr";
                    bucket = "P2.Raft→CW1-NotPosting";
                    confidence = "high";
                    triggers.push("raft_cw1:detected");
                }
            }
            // Check for Wove-specific issues (P2)
            if (/\bwove\b/i.test(content) && (category === "general" || category === "internal_vip")) {
                category = "wove";
                priority = 2;
                sla = "4hr";
                bucket = "P2.Integration-Wove";
                confidence = "medium";
                triggers.push("wove:detected");
            }
            // Check for Warehouse/GHI (P3)
            if (/warehouse order|goods handling|ghi|W\d{7}|putaway|put away|inward.*confirmation|outward.*confirmation|stock receipt|stock transfer|warehouse.*inventory|inventory.*warehouse|warehouse.*stock|stock.*level|sku[-\s]?\d+|warehouse\s*[a-z]/i.test(content)) {
                if (category === "general" || category === "order_status" || category === "internal_vip") {
                    category = "warehouse";
                    priority = 3;
                    sla = "1day";
                    bucket = "P3.Warehouse-Orders-GHI";
                    confidence = "medium";
                    triggers.push("warehouse:detected");
                }
            }
            // Log the classification
            const logEntry = {
                timestamp: new Date().toISOString(),
                sender,
                subject: subject.substring(0, 100),
                domain,
                category,
                bucket,
                priority,
                isVip,
                confidence,
                triggers,
            };
            classificationLog.push(logEntry);
            if (classificationLog.length > MAX_LOG_ENTRIES) {
                classificationLog.shift();
            }
            // Save every 10 classifications
            if (classificationLog.length % 10 === 0) {
                saveLog();
            }
            // Analyze sentiment
            const originalSubject = (args.subject as string) || "";
            const originalBody = (args.body as string) || "";
            const sentiment = analyzeSentiment(originalSubject, originalBody, sender);
            // Adjust priority based on sentiment
            let adjustedPriority = priority;
            if (sentiment.isAutoReply) {
                // Auto-replies get lower priority
                adjustedPriority = Math.min(5, priority + 2);
            }
            else if (sentiment.sentiment === "very_negative" && priority > 1) {
                // Very negative sentiment bumps priority
                adjustedPriority = Math.max(1, priority - 1);
            }
            // Get Outlook categories based on classification + content
            const outlookCategories = getOutlookCategories(category, sender, originalSubject, originalBody);
            // Get action config for additional context
            const action = getCategoryAction(category);
            console.log(`[CLASSIFY] ${domain} → ${category} (${bucket}) [${confidence}] sentiment: ${sentiment.sentiment} (${sentiment.sentimentScore}) urgency: ${sentiment.urgency} outlook: [${outlookCategories.join(", ")}] triggers: ${triggers.join(", ") || "none"}`);
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            classification: {
                                category,
                                priority: adjustedPriority,
                                priorityLabel: adjustedPriority === 1 ? "Urgent" : adjustedPriority === 3 ? "High" : adjustedPriority === 5 ? "Medium" : "Low",
                                sla,
                                recommendedBucket: bucket,
                                isVip,
                                isInternalVip,
                                domain,
                                confidence,
                                triggers,
                            },
                            sentiment: {
                                level: sentiment.sentiment,
                                score: sentiment.sentimentScore,
                                urgency: sentiment.urgency,
                                urgencyScore: sentiment.urgencyScore,
                                isAutoReply: sentiment.isAutoReply,
                                indicators: sentiment.indicators,
                            },
                            outlook: {
                                categories: outlookCategories,
                                createTask: action.createTask,
                                sendNotification: action.sendNotification,
                                teamsChannel: action.teamsChannel || "",
                            },
                        }, null, 2),
                    },
                ],
            };
        }
        case "email_extract_entities": {
            // Support both 'content' and 'body' parameters, plus 'subject'
            const body = ((args.body as string) || (args.content as string) || "");
            const subject = ((args.subject as string) || "");
            const content = `${subject} ${body}`;
            // Extract patterns - invoice/order numbers like INV-2025-9876, ORD-2026-001, PO147198
            const orderNumbers = content.match(/(?:ORD|SO|PO|INV)[-#]?[\d]+(?:[-][\d]+)*/gi) || [];
            const trackingNumbers = content.match(/\b1Z[A-Z0-9]{16}\b|\b\d{12,22}\b/gi) || [];
            const dates = content.match(/\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/g) || [];
            const amounts = content.match(/\$[\d,]+\.?\d*/g) || [];
            const emails = content.match(/[\w.-]+@[\w.-]+\.\w+/gi) || [];
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            entities: {
                                orderNumbers: [...new Set(orderNumbers)],
                                trackingNumbers: [...new Set(trackingNumbers)],
                                dates: [...new Set(dates)],
                                amounts: [...new Set(amounts)],
                                emails: [...new Set(emails)],
                            },
                        }, null, 2),
                    },
                ],
            };
        }
        case "email_get_client_info": {
            // Support both 'domain' and 'from' parameters
            let domain = "";
            if (args.domain) {
                domain = (args.domain as string).toLowerCase();
            }
            else if (args.from) {
                // Extract domain from email address
                const fromAddr = args.from as string;
                const match = fromAddr.match(/@([^>]+)/);
                domain = match ? match[1].toLowerCase() : fromAddr.toLowerCase();
            }
            const clientInfo = CLIENT_INFO[domain];
            if (clientInfo) {
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                found: true,
                                domain,
                                ...clientInfo,
                                isVip: clientInfo.tier === "VIP",
                            }, null, 2),
                        },
                    ],
                };
            }
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            found: false,
                            domain,
                            message: "Client not found in database. Treat as standard client with 2-day SLA.",
                            defaultSla: "2day",
                            defaultBucket: "Client Communications",
                        }, null, 2),
                    },
                ],
            };
        }
        case "email_suggest_response": {
            const subject = (args.subject as string) || "";
            const body = (args.body as string) || "";
            const from = (args.from as string) || "";
            const category = (args.category as string) || "general";
            const tone = (args.tone as string) || "professional";

            const analysisText = `${subject}\n${body}`.toLowerCase();
            const hasAny = (needles: string[]): boolean => needles.some(n => analysisText.includes(n));

            // Extract sender's first name from email or use generic greeting
            const senderName = from.split("@")[0]?.replace(/[._-]/g, " ")?.split(" ")[0] || "";
            const greeting = senderName ? `Hi ${senderName.charAt(0).toUpperCase() + senderName.slice(1)}` : "Hi";

            // Map categories to SLA and response context
            const categoryConfig: Record<string, { sla: string; team: string; signoff: string }> = {
                "P0.Brendan": { sla: "today", team: "priority response team", signoff: "Navia Freight Management" },
                "P1.Urgent": { sla: "1 hour", team: "priority response team", signoff: "Navia Freight Management" },
                "P1.AP-Posting": { sla: "2 hours", team: "accounts team", signoff: "Navia Freight Accounts" },
                "P2.Raft": { sla: "4 hours", team: "Raft support team", signoff: "Navia Freight Operations" },
                "P2.Wove": { sla: "4 hours", team: "Wove integration team", signoff: "Navia Freight Operations" },
                "P2.Integration": { sla: "4 hours", team: "integration team", signoff: "Navia Freight Operations" },
                "P3.Warehouse": { sla: "1 business day", team: "warehouse team", signoff: "Navia Freight Warehouse" },
                "General": { sla: "2 business days", team: "support team", signoff: "Navia Freight Team" },
                // Legacy category names
                "integration_error": { sla: "4 hours", team: "integration team", signoff: "Navia Freight Operations" },
                "client_inquiry": { sla: "2 business days", team: "customer service team", signoff: "Navia Freight Customer Service" },
                "order_status": { sla: "1 business day", team: "operations team", signoff: "Navia Freight Operations" },
                "escalation": { sla: "1 hour", team: "priority response team", signoff: "Navia Freight Management" },
                "vip": { sla: "2 hours", team: "VIP account manager", signoff: "Navia Freight Management" },
                "ap_posting": { sla: "2 hours", team: "accounts team", signoff: "Navia Freight Accounts" },
                "raft_cw1": { sla: "4 hours", team: "Raft support team", signoff: "Navia Freight Operations" },
                "raft_ticket": { sla: "4 hours", team: "Raft support team", signoff: "Navia Freight Operations" },
                "wove": { sla: "4 hours", team: "Wove integration team", signoff: "Navia Freight Operations" },
                "warehouse": { sla: "1 business day", team: "warehouse team", signoff: "Navia Freight Warehouse" },
                "general": { sla: "2 business days", team: "support team", signoff: "Navia Freight Team" },
            };

            const config = categoryConfig[category] || categoryConfig["general"];

            // Adjust language based on tone
            const isUrgent = tone.toLowerCase().includes("urgent") || category.includes("P1") || category.includes("escalation");
            const isEmpathetic = tone.toLowerCase().includes("empathetic") || tone.toLowerCase().includes("apologetic");

            // Lightweight content analysis to improve acknowledgement quality.
            // This intentionally does NOT attempt to fully answer; it only acknowledges receipt and sets expectations.
            const intent = (() => {
                if (hasAny(["system down", "outage", "cannot access", "can't access", "unable to", "blocked", "urgent", "asap", "critical"])) {
                    return "urgent";
                }
                if (hasAny(["approve", "approval", "sign off", "sign-off", "signoff", "authorise", "authorize"])) {
                    return "approval";
                }
                if (hasAny(["invoice", "payment", "remittance", "ap ", "a/p", "posting", "credit note", "refund"])) {
                    return "finance";
                }
                if (hasAny(["meeting", "call", "teams", "zoom", "calendar", "schedule", "reschedule"])) {
                    return "meeting";
                }
                if (hasAny(["eta", "timeline", "when can", "by when", "due date", "deadline"])) {
                    return "timing";
                }
                if (hasAny(["integration", "api", "webhook", "odoo", "cargowise", "cw1", "edi", "sync", "mapping"])) {
                    return "integration";
                }
                return "general";
            })();

            const acknowledgementLine = (() => {
                switch (intent) {
                    case "urgent":
                        return "Received - I have flagged this as urgent and I am on it.";
                    case "approval":
                        return "Received - I will review what you need approved and confirm shortly.";
                    case "finance":
                        return "Received - I will review the finance/AP item and revert with next steps.";
                    case "meeting":
                        return "Received - I will check timing and come back with a proposed time / confirmation.";
                    case "timing":
                        return "Received - I will review the timing/deadline and confirm the next update.";
                    case "integration":
                        return "Received - I will review the integration item and revert with an update.";
                    default:
                        return "Received - I will review and revert shortly.";
                }
            })();

            let responseBody = "";
            if (isUrgent) {
                responseBody = `${greeting},

${acknowledgementLine}

This has been flagged as a priority item and our ${config.team} is actively working on it. We aim to provide an update within ${config.sla}.

If you have any additional details that may help us resolve this faster, please reply directly to this email.

Best regards,
${config.signoff}`;
            } else if (isEmpathetic) {
                responseBody = `${greeting},

${acknowledgementLine}

Our ${config.team} has been notified and will respond within ${config.sla}. We are committed to resolving this promptly.

Please don't hesitate to reach out if you need anything further in the meantime.

Kind regards,
${config.signoff}`;
            } else {
                responseBody = `${greeting},

${acknowledgementLine}

We have received your message and our ${config.team} will review and respond within ${config.sla}.

If you have any additional information to share, please reply to this email.

Best regards,
${config.signoff}`;
            }

            console.log(`[SUGGEST_RESPONSE] category=${category} tone=${tone} from=${from} subject=${subject.substring(0, 50)}`);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            category,
                            tone,
                            sla: config.sla,
                            suggestedResponse: responseBody,
                        }, null, 2),
                    },
                ],
            };
        }
        case "email_analyze_sentiment": {
            const subject = (args.subject as string) || "";
            const body = (args.body as string) || "";
            const sender = (args.sender as string) || "";

            const sentiment = analyzeSentiment(subject, body, sender);

            console.log(`[ANALYZE_SENTIMENT] sender=${sender.substring(0, 30)} sentiment=${sentiment.sentiment} urgency=${sentiment.urgency}`);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            sentiment: {
                                level: sentiment.sentiment,
                                score: sentiment.sentimentScore,
                            },
                            urgency: {
                                level: sentiment.urgency,
                                score: sentiment.urgencyScore,
                            },
                            isAutoReply: sentiment.isAutoReply,
                            indicators: sentiment.indicators,
                        }, null, 2),
                    },
                ],
            };
        }

        case "email_summarize_thread": {
            const emails = args.emails as Array<{ from: string; subject?: string; body: string; date?: string }>;

            if (!emails || emails.length === 0) {
                return {
                    content: [{ type: "text", text: JSON.stringify({ error: "No emails provided" }, null, 2) }],
                };
            }

            // Extract participants
            const participants = [...new Set(emails.map(e => e.from))];

            // Extract action items (lines starting with action verbs or containing "please", "need to", etc.)
            const actionItemPatterns = /(?:^|\n)\s*[-•*]?\s*(?:please|need to|action required|todo|to do|follow up|can you|could you|would you|must|should|will you|kindly|ensure|make sure|don't forget|remember to)\b[^.\n]*/gi;
            const allBodies = emails.map(e => e.body).join("\n");
            const actionMatches = allBodies.match(actionItemPatterns) || [];
            const actionItems = actionMatches.map(a => a.trim().replace(/^[-•*]\s*/, "")).slice(0, 10);

            // Extract key points (sentences with important keywords)
            const keyPointPatterns = /(?:decided|agreed|confirmed|approved|rejected|deadline|due date|budget|cost|timeline|milestone|update|status|completed|resolved|issue|problem|solution|proposal|recommend)/gi;
            const sentences = allBodies.split(/[.!?\n]+/).filter(s => s.trim().length > 20);
            const keyPoints = sentences
                .filter(s => keyPointPatterns.test(s))
                .map(s => s.trim())
                .slice(0, 8);

            // Thread metadata
            const threadSubject = emails.find(e => e.subject)?.subject || "(no subject)";
            const latestSentiment = analyzeSentiment(
                emails[emails.length - 1].subject || "",
                emails[emails.length - 1].body,
                emails[emails.length - 1].from,
            );

            console.log(`[SUMMARIZE_THREAD] subject="${threadSubject.substring(0, 50)}" emails=${emails.length} participants=${participants.length}`);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            summary: {
                                subject: threadSubject,
                                emailCount: emails.length,
                                participants,
                                dateRange: {
                                    first: emails[0].date || null,
                                    last: emails[emails.length - 1].date || null,
                                },
                            },
                            keyPoints: keyPoints.length > 0 ? keyPoints : ["No key decision points detected — manual review recommended"],
                            actionItems: actionItems.length > 0 ? actionItems : ["No explicit action items detected"],
                            latestSentiment: {
                                level: latestSentiment.sentiment,
                                urgency: latestSentiment.urgency,
                            },
                        }, null, 2),
                    },
                ],
            };
        }

        case "email_bulk_triage": {
            const emails = args.emails as Array<{ id?: string; sender: string; subject: string; body?: string }>;

            if (!emails || emails.length === 0) {
                return {
                    content: [{ type: "text", text: JSON.stringify({ error: "No emails provided" }, null, 2) }],
                };
            }

            // Process each email through the classify logic
            const results = [];
            for (const email of emails) {
                // Call handleEmailTool recursively for each email
                const classifyResult = await handleEmailTool("email_classify", {
                    sender: email.sender,
                    subject: email.subject,
                    body: email.body || "",
                });

                const parsed = JSON.parse(classifyResult.content[0].text);
                results.push({
                    id: email.id || `${email.sender}:${email.subject.substring(0, 30)}`,
                    sender: email.sender,
                    subject: email.subject.substring(0, 80),
                    classification: parsed.classification,
                    sentiment: parsed.sentiment,
                    outlook: parsed.outlook,
                });
            }

            // Sort by priority (lower number = higher priority)
            results.sort((a, b) => (a.classification?.priority || 5) - (b.classification?.priority || 5));

            console.log(`[BULK_TRIAGE] Processed ${results.length} emails`);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            processed: results.length,
                            summary: {
                                urgent: results.filter(r => r.classification?.priority === 1).length,
                                high: results.filter(r => r.classification?.priority === 2 || r.classification?.priority === 3).length,
                                medium: results.filter(r => r.classification?.priority === 5).length,
                                low: results.filter(r => r.classification?.priority > 5).length,
                            },
                            results,
                        }, null, 2),
                    },
                ],
            };
        }

        case "email_score_priority": {
            const sender = ((args.sender as string) || "").toLowerCase();
            const subject = ((args.subject as string) || "");
            const body = ((args.body as string) || "");

            // Get classification
            const classifyResult = await handleEmailTool("email_classify", { sender, subject, body });
            const parsed = JSON.parse(classifyResult.content[0].text);
            const classification = parsed.classification;
            const sentimentData = parsed.sentiment;

            // Calculate composite score (0-100, higher = more important)
            let score = 50; // Base score

            // VIP bonus (+20)
            if (classification.isVip) score += 20;
            if (classification.isInternalVip) score += 10; // Extra for internal

            // Priority adjustment (P1=+30, P2=+15, P3=+5)
            if (classification.priority === 1) score += 30;
            else if (classification.priority <= 3) score += 15;
            else if (classification.priority <= 5) score += 5;

            // Sentiment adjustment
            if (sentimentData.level === "very_negative") score += 15;
            else if (sentimentData.level === "negative") score += 8;

            // Urgency adjustment
            if (sentimentData.urgency === "critical") score += 15;
            else if (sentimentData.urgency === "high") score += 10;

            // Auto-reply penalty
            if (sentimentData.isAutoReply) score -= 40;

            // Category bonuses
            if (classification.category === "escalation") score += 10;
            if (classification.category === "ap_posting") score += 5;

            // Clamp to 0-100
            score = Math.max(0, Math.min(100, score));

            const priorityLabel = score >= 80 ? "critical" : score >= 60 ? "high" : score >= 40 ? "medium" : "low";

            console.log(`[SCORE_PRIORITY] sender=${sender.substring(0, 30)} score=${score} label=${priorityLabel}`);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            score,
                            priorityLabel,
                            breakdown: {
                                base: 50,
                                vipBonus: classification.isVip ? (classification.isInternalVip ? 30 : 20) : 0,
                                priorityBonus: classification.priority === 1 ? 30 : classification.priority <= 3 ? 15 : classification.priority <= 5 ? 5 : 0,
                                sentimentAdjustment: sentimentData.level === "very_negative" ? 15 : sentimentData.level === "negative" ? 8 : 0,
                                urgencyAdjustment: sentimentData.urgency === "critical" ? 15 : sentimentData.urgency === "high" ? 10 : 0,
                                autoReplyPenalty: sentimentData.isAutoReply ? -40 : 0,
                                categoryBonus: classification.category === "escalation" ? 10 : classification.category === "ap_posting" ? 5 : 0,
                            },
                            classification: {
                                category: classification.category,
                                priority: classification.priority,
                                isVip: classification.isVip,
                                confidence: classification.confidence,
                            },
                            sentiment: {
                                level: sentimentData.level,
                                urgency: sentimentData.urgency,
                                isAutoReply: sentimentData.isAutoReply,
                            },
                        }, null, 2),
                    },
                ],
            };
        }

        case "email_login": {
            try {
                const tenantId = process.env.AZURE_TENANT_ID;
                const clientId = process.env.AZURE_CLIENT_ID;
                if (!tenantId || !clientId) {
                    throw new Error("Missing AZURE_TENANT_ID or AZURE_CLIENT_ID in .env");
                }
                console.log("[EMAIL_LOGIN] Starting device code authentication...");
                // Create device code credential
                const credential = new DeviceCodeCredential({
                    tenantId,
                    clientId,
                    userPromptCallback: (info) => {
                        console.log("\n========================================");
                        console.log("TO SIGN IN:");
                        console.log(`1. Go to: ${info.verificationUri}`);
                        console.log(`2. Enter code: ${info.userCode}`);
                        console.log("========================================\n");
                    },
                });
                // Get token (this will trigger the device code prompt)
                const tokenResponse = await credential.getToken([
                    "https://graph.microsoft.com/Mail.Read",
                    "https://graph.microsoft.com/User.Read",
                ]);
                // Create Graph client with user token
                userGraphClient = Client.init({
                    authProvider: (done) => {
                        done(null, tokenResponse.token);
                    },
                });
                // Get user info to confirm login
                const me = await userGraphClient.api("/me").get();
                userEmail = me.mail || me.userPrincipalName;
                console.log(`[EMAIL_LOGIN] Successfully logged in as: ${userEmail}`);
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                success: true,
                                message: "Successfully authenticated!",
                                user: userEmail,
                                displayName: me.displayName,
                                hint: "You can now use email_fetch_recent to read your emails.",
                            }, null, 2),
                        },
                    ],
                };
            }
            catch (error) {
                const errorMessage = (error as { message?: string }).message || "Unknown error";
                console.error("[EMAIL_LOGIN] Error:", errorMessage);
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                success: false,
                                error: errorMessage,
                                hint: "Make sure to complete the sign-in at microsoft.com/devicelogin with the code shown in the server console.",
                            }, null, 2),
                        },
                    ],
                };
            }
        }
        case "email_fetch_recent": {
            const count = Math.min(Math.max((args.count as number) || 10, 1), 50);
            const folder = (args.folder as string) || "inbox";
            const unreadOnly = (args.unreadOnly as boolean) || false;
            // Check if user is logged in
            if (!userGraphClient) {
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                success: false,
                                error: "Not logged in. Run email_login first to authenticate.",
                                hint: "Call the email_login tool, then check the server console for the device code.",
                            }, null, 2),
                        },
                    ],
                };
            }
            try {
                const folderMap: Record<string, string> = {
                    inbox: "inbox",
                    sentitems: "sentItems",
                    drafts: "drafts",
                };
                const folderName = folderMap[folder.toLowerCase()] || "inbox";
                // Build filter for unread if requested
                const filter = unreadOnly ? "&$filter=isRead eq false" : "";
                // Fetch emails using user's token
                const response = await (userGraphClient as Client)
                    .api(`/me/mailFolders/${folderName}/messages?$top=${count}&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,importance${filter}`)
                    .get();
                const emails = (response as GraphMessagesResponse).value.map((email: GraphMessage) => ({
                    id: email.id,
                    subject: email.subject || "(no subject)",
                    from: email.from?.emailAddress?.address || "unknown",
                    fromName: email.from?.emailAddress?.name || "",
                    receivedAt: email.receivedDateTime,
                    isRead: email.isRead,
                    importance: email.importance,
                    preview: email.bodyPreview?.substring(0, 200) || "",
                }));
                console.log(`[EMAIL_FETCH] Retrieved ${emails.length} emails for ${userEmail}`);
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                success: true,
                                user: userEmail,
                                folder: folderName,
                                count: emails.length,
                                emails,
                            }, null, 2),
                        },
                    ],
                };
            }
            catch (error) {
                const errorMessage = (error as { message?: string }).message || "Unknown error";
                console.error("[EMAIL_FETCH] Error:", errorMessage);
                // Token might have expired
                if (errorMessage.includes("401") || errorMessage.includes("token")) {
                    userGraphClient = null;
                    userEmail = null;
                }
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                success: false,
                                error: errorMessage,
                                hint: errorMessage.includes("401")
                                    ? "Session expired. Run email_login again."
                                    : "Check the folder name and try again.",
                            }, null, 2),
                        },
                    ],
                };
            }
        }
        case "email_read_pst": {
            const folderName = (args.folder as string) || "read";
            const count = Math.min(Math.max((args.count as number) || 20, 1), 100);
            try {
                if (!fs.existsSync(PST_PATH)) {
                    return {
                        content: [
                            {
                                type: "text",
                                text: JSON.stringify({
                                    success: false,
                                    error: `PST file not found at ${PST_PATH}`,
                                }, null, 2),
                            },
                        ],
                    };
                }
                const pstFile = new (PSTFile as unknown as { new(path: string): PstFileInstance })(PST_PATH);
                const rootFolder = pstFile.getRootFolder();
                // Helper to find folder
                function findFolder(folder: PstFolder, target: string): PstFolder | null {
                    if (folder.displayName?.toLowerCase() === target.toLowerCase()) {
                        return folder;
                    }
                    if (folder.hasSubfolders) {
                        for (const sub of folder.getSubFolders()) {
                            const found = findFolder(sub, target);
                            if (found)
                                return found;
                        }
                    }
                    return null;
                }
                // Helper to list folders
                function listFolders(folder: PstFolder, indent: string = ""): string[] {
                    const result: string[] = [];
                    const name = folder.displayName || "(root)";
                    const itemCount = folder.contentCount || 0;
                    if (itemCount > 0) {
                        result.push(`${indent}${name} (${itemCount} items)`);
                    }
                    if (folder.hasSubfolders) {
                        for (const sub of folder.getSubFolders()) {
                            result.push(...listFolders(sub, indent + "  "));
                        }
                    }
                    return result;
                }
                // List folders mode
                if (folderName.toLowerCase() === "folders") {
                    const folders = listFolders(rootFolder);
                    pstFile.close();
                    return {
                        content: [
                            {
                                type: "text",
                                text: JSON.stringify({
                                    success: true,
                                    mode: "list_folders",
                                    folders,
                                }, null, 2),
                            },
                        ],
                    };
                }
                // Find and read folder
                const targetFolder = findFolder(rootFolder, folderName);
                if (!targetFolder) {
                    const folders = listFolders(rootFolder);
                    pstFile.close();
                    return {
                        content: [
                            {
                                type: "text",
                                text: JSON.stringify({
                                    success: false,
                                    error: `Folder "${folderName}" not found`,
                                    availableFolders: folders.slice(0, 30),
                                }, null, 2),
                            },
                        ],
                    };
                }
                // Read emails
                const emails: Array<{ subject: string; from: string; fromName: string; receivedDate: string | null; body: string; importance: string; }> = [];
                if ((targetFolder.contentCount || 0) > 0) {
                    let email: PstMessage | null = targetFolder.getNextChild();
                    while (email && emails.length < count) {
                        if (email.messageClass === "IPM.Note") {
                            emails.push({
                                subject: email.subject || "(no subject)",
                                from: email.senderEmailAddress || email.senderName || "unknown",
                                fromName: email.senderName || "",
                                receivedDate: email.messageDeliveryTime?.toISOString() || null,
                                body: (email.body || "").substring(0, 500),
                                importance: email.importance === 2 ? "high" : email.importance === 0 ? "low" : "normal",
                            });
                        }
                        email = targetFolder.getNextChild();
                    }
                }
                // Sort by date
                emails.sort((a, b) => {
                    const dateA = a.receivedDate ? new Date(a.receivedDate).getTime() : 0;
                    const dateB = b.receivedDate ? new Date(b.receivedDate).getTime() : 0;
                    return dateB - dateA;
                });
                pstFile.close();
                console.log(`[EMAIL_PST] Read ${emails.length} emails from folder "${folderName}"`);
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                success: true,
                                folder: folderName,
                                totalInFolder: targetFolder.contentCount || 0,
                                returned: emails.length,
                                emails,
                            }, null, 2),
                        },
                    ],
                };
            }
            catch (error) {
                const errorMessage = (error as { message?: string }).message || "Unknown error";
                console.error("[EMAIL_PST] Error:", errorMessage);
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                success: false,
                                error: errorMessage,
                            }, null, 2),
                        },
                    ],
                };
            }
        }
        default:
            throw new Error(`Unknown email tool: ${name}`);
    }
}
