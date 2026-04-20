import type { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient } from "../lib/graph.js";

type ToolResult = Promise<{ content: { type: string; text: string }[] }>;

interface GraphEmailAddress {
    address?: string;
    name?: string;
}

interface GraphRecipient {
    emailAddress?: GraphEmailAddress;
}

interface GraphMessage {
    id: string;
    subject?: string;
    from?: GraphRecipient;
    toRecipients?: GraphRecipient[];
    ccRecipients?: GraphRecipient[];
    receivedDateTime?: string;
    sentDateTime?: string;
    isRead?: boolean;
    importance?: string;
    flag?: unknown;
    categories?: string[];
    hasAttachments?: boolean;
    bodyPreview?: string;
    body?: unknown;
}

interface GraphListResponse<T> {
    value: T[];
}

interface GraphBatchSubResponse {
    id: string;
    status: number;
    body?: {
        error?: {
            message?: string;
        };
    };
}

interface GraphBatchResponse {
    responses: GraphBatchSubResponse[];
}

interface MasterCategory {
    id: string;
    displayName: string;
    color: string;
}

export const outlookTools: Tool[] = [
    {
        name: "outlook_get_emails",
        description: "Get emails from a mailbox. Requires user ID or UPN for app-only auth.",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN (e.g., user@domain.com). Required for app-only auth.",
                },
                folder: {
                    type: "string",
                    enum: ["inbox", "sentitems", "drafts", "deleteditems", "archive"],
                    description: "Mail folder to query. Default: inbox",
                },
                top: {
                    type: "number",
                    description: "Number of emails to return. Default: 10, Max: 50",
                },
                filter: {
                    type: "string",
                    description: "OData filter (e.g., \"isRead eq false\", \"from/emailAddress/address eq 'sender@domain.com'\")",
                },
                search: {
                    type: "string",
                    description: "Search query (e.g., \"subject:urgent\")",
                },
                orderBy: {
                    type: "string",
                    description: "Sort order (e.g., \"receivedDateTime desc\")",
                },
            },
            required: ["userId"],
        },
    },
    {
        name: "outlook_get_email",
        description: "Get a specific email by ID with full body content",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN",
                },
                messageId: {
                    type: "string",
                    description: "The email message ID",
                },
            },
            required: ["userId", "messageId"],
        },
    },
    {
        name: "outlook_flag_email",
        description: "Set or clear a flag on an email",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN",
                },
                messageId: {
                    type: "string",
                    description: "The email message ID",
                },
                flagStatus: {
                    type: "string",
                    enum: ["notFlagged", "flagged", "complete"],
                    description: "Flag status to set",
                },
                dueDateTime: {
                    type: "string",
                    description: "Due date for the flag (ISO 8601 format)",
                },
            },
            required: ["userId", "messageId", "flagStatus"],
        },
    },
    {
        name: "outlook_assign_category",
        description: "Assign one or more categories to an email",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN",
                },
                messageId: {
                    type: "string",
                    description: "The email message ID",
                },
                categories: {
                    type: "array",
                    items: { type: "string" },
                    description: "Categories to assign (e.g., [\"P1 - Urgent\", \"VIP Client\"])",
                },
            },
            required: ["userId", "messageId", "categories"],
        },
    },
    {
        name: "outlook_assign_categories_bulk",
        description: "Assign categories to multiple emails at once",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN",
                },
                messageIds: {
                    type: "array",
                    items: { type: "string" },
                    description: "Array of email message IDs",
                },
                categories: {
                    type: "array",
                    items: { type: "string" },
                    description: "Categories to assign to all emails",
                },
            },
            required: ["userId", "messageIds", "categories"],
        },
    },
    {
        name: "outlook_list_categories",
        description: "List available Outlook categories for a user",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN",
                },
            },
            required: ["userId"],
        },
    },
    {
        name: "outlook_mark_read",
        description: "Mark an email as read or unread",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN",
                },
                messageId: {
                    type: "string",
                    description: "The email message ID",
                },
                isRead: {
                    type: "boolean",
                    description: "true to mark as read, false to mark as unread",
                },
            },
            required: ["userId", "messageId", "isRead"],
        },
    },
    {
        name: "outlook_move_email",
        description: "Move an email to a different folder",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN",
                },
                messageId: {
                    type: "string",
                    description: "The email message ID",
                },
                destinationFolder: {
                    type: "string",
                    description: "Destination folder ID or well-known name (inbox, archive, deleteditems)",
                },
            },
            required: ["userId", "messageId", "destinationFolder"],
        },
    },
    {
        name: "outlook_send_email",
        description: "Send an email",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN (sender)",
                },
                to: {
                    type: "array",
                    items: { type: "string" },
                    description: "Recipient email addresses",
                },
                cc: {
                    type: "array",
                    items: { type: "string" },
                    description: "CC recipient email addresses",
                },
                subject: {
                    type: "string",
                    description: "Email subject",
                },
                body: {
                    type: "string",
                    description: "Email body (HTML supported)",
                },
                bodyType: {
                    type: "string",
                    enum: ["text", "html"],
                    description: "Body content type. Default: html",
                },
                importance: {
                    type: "string",
                    enum: ["low", "normal", "high"],
                    description: "Email importance",
                },
            },
            required: ["userId", "to", "subject", "body"],
        },
    },
    {
        name: "outlook_get_calendar_events",
        description: "Get calendar events for a user within a date range. Returns event subject, start/end times, attendees, and location.",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN (e.g., user@domain.com)",
                },
                startDateTime: {
                    type: "string",
                    description: "Start of date range (ISO 8601, e.g., '2026-03-06T00:00:00Z')",
                },
                endDateTime: {
                    type: "string",
                    description: "End of date range (ISO 8601, e.g., '2026-03-07T23:59:59Z')",
                },
                top: {
                    type: "number",
                    description: "Maximum events to return. Default: 20, Max: 50",
                },
            },
            required: ["userId", "startDateTime", "endDateTime"],
        },
    },
    {
        name: "outlook_check_availability",
        description: "Check the availability (free/busy) of one or more users for a given time range. Useful for scheduling meetings.",
        inputSchema: {
            type: "object",
            properties: {
                userId: {
                    type: "string",
                    description: "User ID or UPN of the requesting user",
                },
                schedules: {
                    type: "array",
                    items: { type: "string" },
                    description: "Email addresses to check availability for",
                },
                startDateTime: {
                    type: "string",
                    description: "Start of availability window (ISO 8601)",
                },
                endDateTime: {
                    type: "string",
                    description: "End of availability window (ISO 8601)",
                },
                availabilityViewInterval: {
                    type: "number",
                    description: "Duration of each time slot in minutes (default: 30)",
                },
            },
            required: ["userId", "schedules", "startDateTime", "endDateTime"],
        },
    },
];

// Default user for testing
const DEFAULT_USER_ID = process.env.OUTLOOK_USER_ID || "";

export async function handleOutlookTool(name: string, args: Record<string, unknown>): ToolResult {
    const client = await getGraphClient();
    const userId = ((args.userId as string | undefined) || DEFAULT_USER_ID) as string;

    if (!userId) {
        throw new Error("userId is required for Outlook operations with app-only auth");
    }

    switch (name) {
        case "outlook_get_emails": {
            const folder = (args.folder as string | undefined) || "inbox";
            const top = Math.min((args.top as number | undefined) || 10, 50);
            const filter = args.filter as string | undefined;
            const search = args.search as string | undefined;
            const orderBy = (args.orderBy as string | undefined) || "receivedDateTime desc";

            let request = client
                .api(`/users/${userId}/mailFolders/${folder}/messages`)
                .top(top)
                .orderby(orderBy)
                .select("id,subject,from,toRecipients,receivedDateTime,isRead,importance,flag,categories,hasAttachments,bodyPreview");

            if (filter) {
                request = request.filter(filter);
            }

            if (search) {
                request = request.search(search);
            }

            const result = (await request.get()) as GraphListResponse<GraphMessage>;
            const emails = result.value.map((email: GraphMessage) => ({
                id: email.id,
                subject: email.subject,
                from: email.from?.emailAddress,
                to: email.toRecipients?.map((r: GraphRecipient) => r.emailAddress),
                receivedDateTime: email.receivedDateTime,
                isRead: email.isRead,
                importance: email.importance,
                flag: email.flag,
                categories: email.categories,
                hasAttachments: email.hasAttachments,
                bodyPreview: email.bodyPreview,
            }));

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({ count: emails.length, emails }, null, 2),
                    },
                ],
            };
        }

        case "outlook_get_email": {
            const messageId = args.messageId as string;
            const email = (await client
                .api(`/users/${userId}/messages/${messageId}`)
                .select("id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,importance,flag,categories,hasAttachments,body,internetMessageHeaders")
                .get()) as GraphMessage;

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                id: email.id,
                                subject: email.subject,
                                from: email.from?.emailAddress,
                                to: email.toRecipients?.map((r: GraphRecipient) => r.emailAddress),
                                cc: email.ccRecipients?.map((r: GraphRecipient) => r.emailAddress),
                                receivedDateTime: email.receivedDateTime,
                                sentDateTime: email.sentDateTime,
                                isRead: email.isRead,
                                importance: email.importance,
                                flag: email.flag,
                                categories: email.categories,
                                hasAttachments: email.hasAttachments,
                                body: email.body,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "outlook_flag_email": {
            const messageId = args.messageId as string;
            const flagStatus = args.flagStatus as string;
            const dueDateTime = args.dueDateTime as string | undefined;
            const flagPayload: {
                flag: {
                    flagStatus: string;
                    dueDateTime?: {
                        dateTime: string;
                        timeZone: string;
                    };
                };
            } = {
                flag: {
                    flagStatus,
                },
            };

            if (dueDateTime && flagStatus === "flagged") {
                flagPayload.flag.dueDateTime = {
                    dateTime: dueDateTime,
                    timeZone: "UTC",
                };
            }

            await client.api(`/users/${userId}/messages/${messageId}`).patch(flagPayload);
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                messageId,
                                flagStatus,
                                dueDateTime,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "outlook_assign_category": {
            const messageId = args.messageId as string;
            const categories = args.categories as string[];
            await client.api(`/users/${userId}/messages/${messageId}`).patch({
                categories,
            });

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                messageId,
                                categories,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "outlook_assign_categories_bulk": {
            const messageIds = args.messageIds as string[];
            const categories = args.categories as string[];
            const results: Array<{ messageId: string; success: boolean; error?: string }> = [];

            // Use batch API for efficiency
            const BATCH_SIZE = 20;

            for (let i = 0; i < messageIds.length; i += BATCH_SIZE) {
                const batch = messageIds.slice(i, i + BATCH_SIZE);
                const batchRequests = batch.map((messageId: string, index: number) => ({
                    id: `${index}`,
                    method: "PATCH",
                    url: `/users/${userId}/messages/${messageId}`,
                    headers: { "Content-Type": "application/json" },
                    body: { categories },
                }));

                try {
                    const batchResponse = (await client.api("/$batch").post({
                        requests: batchRequests,
                    })) as GraphBatchResponse;

                    for (const response of batchResponse.responses) {
                        const messageId = batch[parseInt(response.id, 10)];
                        if (response.status >= 200 && response.status < 300) {
                            results.push({ messageId, success: true });
                        } else {
                            results.push({
                                messageId,
                                success: false,
                                error: response.body?.error?.message || `HTTP ${response.status}`,
                            });
                        }
                    }
                } catch (error: unknown) {
                    const errorMessage = error instanceof Error ? error.message : "Unknown error";
                    for (const messageId of batch) {
                        results.push({ messageId, success: false, error: errorMessage });
                    }
                }
            }

            const successCount = results.filter((r) => r.success).length;
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: successCount === messageIds.length,
                                summary: {
                                    total: messageIds.length,
                                    succeeded: successCount,
                                    failed: messageIds.length - successCount,
                                },
                                categories,
                                results,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "outlook_list_categories": {
            const result = (await client.api(`/users/${userId}/outlook/masterCategories`).get()) as GraphListResponse<MasterCategory>;
            const categories = result.value.map((cat: MasterCategory) => ({
                id: cat.id,
                displayName: cat.displayName,
                color: cat.color,
            }));

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({ categories }, null, 2),
                    },
                ],
            };
        }

        case "outlook_mark_read": {
            const messageId = args.messageId as string;
            const isRead = args.isRead as boolean;
            await client.api(`/users/${userId}/messages/${messageId}`).patch({ isRead });

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                messageId,
                                isRead,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "outlook_move_email": {
            const messageId = args.messageId as string;
            const destinationFolder = args.destinationFolder as string;

            // Well-known folder names map to IDs
            const wellKnownFolders = ["inbox", "archive", "deleteditems", "drafts", "sentitems", "junkemail"];
            const destinationId = wellKnownFolders.includes(destinationFolder.toLowerCase())
                ? destinationFolder
                : destinationFolder;

            const result = (await client.api(`/users/${userId}/messages/${messageId}/move`).post({
                destinationId,
            })) as { id: string };

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                messageId: result.id,
                                newLocation: destinationFolder,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "outlook_send_email": {
            const to = args.to as string[];
            const cc = args.cc as string[] | undefined;
            const subject = args.subject as string;
            const body = args.body as string;
            const bodyType = (args.bodyType as string | undefined) || "html";
            const importance = (args.importance as string | undefined) || "normal";

            const message = {
                message: {
                    subject,
                    body: {
                        contentType: bodyType === "html" ? "HTML" : "Text",
                        content: body,
                    },
                    toRecipients: to.map((email: string) => ({
                        emailAddress: { address: email },
                    })),
                    ccRecipients: cc?.map((email: string) => ({
                        emailAddress: { address: email },
                    })),
                    importance,
                },
                saveToSentItems: true,
            };

            await client.api(`/users/${userId}/sendMail`).post(message);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                to,
                                cc,
                                subject,
                                importance,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "outlook_get_calendar_events": {
            const startDateTime = args.startDateTime as string;
            const endDateTime = args.endDateTime as string;
            const top = Math.min((args.top as number | undefined) || 20, 50);

            const events = (await client
                .api(`/users/${userId}/calendarView`)
                .query({
                    startDateTime,
                    endDateTime,
                    $top: top,
                    $orderby: "start/dateTime",
                    $select: "id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,showAs,importance",
                })
                .header("Prefer", 'outlook.timezone="Australia/Perth"')
                .get()) as GraphListResponse<Record<string, unknown>>;

            const eventList = events.value.map((event: Record<string, unknown>) => ({
                id: event.id,
                subject: event.subject,
                start: event.start,
                end: event.end,
                location: event.location,
                organizer: event.organizer,
                attendees: event.attendees,
                isAllDay: event.isAllDay,
                isCancelled: event.isCancelled,
                showAs: event.showAs,
                importance: event.importance,
            }));

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            count: eventList.length,
                            dateRange: { start: startDateTime, end: endDateTime },
                            events: eventList,
                        }, null, 2),
                    },
                ],
            };
        }

        case "outlook_check_availability": {
            const schedules = args.schedules as string[];
            const startDateTime = args.startDateTime as string;
            const endDateTime = args.endDateTime as string;
            const interval = (args.availabilityViewInterval as number | undefined) || 30;

            const result = (await client
                .api(`/users/${userId}/calendar/getSchedule`)
                .post({
                    schedules,
                    startTime: {
                        dateTime: startDateTime,
                        timeZone: "Australia/Perth",
                    },
                    endTime: {
                        dateTime: endDateTime,
                        timeZone: "Australia/Perth",
                    },
                    availabilityViewInterval: interval,
                })) as { value: Array<Record<string, unknown>> };

            const scheduleResults = result.value.map((schedule: Record<string, unknown>) => ({
                email: schedule.scheduleId,
                availabilityView: schedule.availabilityView,
                scheduleItems: schedule.scheduleItems,
                workingHours: schedule.workingHours,
            }));

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            count: scheduleResults.length,
                            timeRange: { start: startDateTime, end: endDateTime },
                            intervalMinutes: interval,
                            schedules: scheduleResults,
                        }, null, 2),
                    },
                ],
            };
        }

        default:
            throw new Error(`Unknown outlook tool: ${name}`);
    }
}
