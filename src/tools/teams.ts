import type { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient } from "../lib/graph.js";

export const teamsTools: Tool[] = [
    {
        name: "teams_post_message",
        description: "Post a message to a Teams channel",
        inputSchema: {
            type: "object",
            properties: {
                teamId: {
                    type: "string",
                    description: "The Teams team ID",
                },
                channelId: {
                    type: "string",
                    description: "The channel ID within the team",
                },
                message: {
                    type: "string",
                    description: "The message content (supports HTML)",
                },
                importance: {
                    type: "string",
                    enum: ["normal", "high", "urgent"],
                    description: "Message importance level",
                },
            },
            required: ["teamId", "channelId", "message"],
        },
    },
    {
        name: "teams_post_adaptive_card",
        description: "Post an adaptive card to a Teams channel and optionally wait for response",
        inputSchema: {
            type: "object",
            properties: {
                teamId: {
                    type: "string",
                    description: "The Teams team ID",
                },
                channelId: {
                    type: "string",
                    description: "The channel ID within the team",
                },
                card: {
                    type: "object",
                    description: "The adaptive card JSON payload",
                },
                waitForResponse: {
                    type: "boolean",
                    description: "Whether to wait for user response (not supported via Graph - use Power Automate)",
                },
            },
            required: ["teamId", "channelId", "card"],
        },
    },
    {
        name: "teams_list_channels",
        description: "List channels in a team",
        inputSchema: {
            type: "object",
            properties: {
                teamId: {
                    type: "string",
                    description: "The Teams team ID",
                },
            },
            required: ["teamId"],
        },
    },
    {
        name: "teams_list_teams",
        description: "List teams the app has access to",
        inputSchema: {
            type: "object",
            properties: {},
        },
    },
];

// Default team/channel from env
const DEFAULT_TEAM_ID = process.env.TEAMS_GROUP_ID || "";
const DEFAULT_CHANNEL_ID = process.env.TEAMS_CHANNEL_ID || "";

export async function handleTeamsTool(
    name: string,
    args: Record<string, unknown>,
): Promise<{ content: { type: string; text: string }[] }> {
    const client = await getGraphClient();

    switch (name) {
        case "teams_post_message": {
            const teamIdArg = args.teamId as string | undefined;
            const channelIdArg = args.channelId as string | undefined;
            const messageArg = args.message as string;
            const importanceArg = args.importance as string | undefined;

            const teamId = teamIdArg || DEFAULT_TEAM_ID;
            const channelId = channelIdArg || DEFAULT_CHANNEL_ID;
            const message = messageArg;
            const importance = importanceArg || "normal";

            if (!teamId || !channelId) {
                throw new Error("teamId and channelId are required");
            }

            const messagePayload = {
                body: {
                    contentType: "html",
                    content: message,
                },
                importance,
            };

            const result = await client
                .api(`/teams/${teamId}/channels/${channelId}/messages`)
                .post(messagePayload);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                messageId: result.id,
                                teamId,
                                channelId,
                                webUrl: result.webUrl,
                                createdDateTime: result.createdDateTime,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "teams_post_adaptive_card": {
            const teamIdArg = args.teamId as string | undefined;
            const channelIdArg = args.channelId as string | undefined;
            const cardArg = args.card as Record<string, unknown>;
            const waitForResponseArg = args.waitForResponse as boolean | undefined;

            const teamId = teamIdArg || DEFAULT_TEAM_ID;
            const channelId = channelIdArg || DEFAULT_CHANNEL_ID;
            const card = cardArg;

            if (!teamId || !channelId) {
                throw new Error("teamId and channelId are required");
            }

            // Wrap the card in the required attachment format
            const messagePayload = {
                body: {
                    contentType: "html",
                    content: "", // Empty for adaptive card
                },
                attachments: [
                    {
                        contentType: "application/vnd.microsoft.card.adaptive",
                        content: JSON.stringify(card),
                    },
                ],
            };

            const result = await client
                .api(`/teams/${teamId}/channels/${channelId}/messages`)
                .post(messagePayload);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                messageId: result.id,
                                teamId,
                                channelId,
                                webUrl: result.webUrl,
                                note: waitForResponseArg
                                    ? "waitForResponse requires Power Automate - Graph API does not support waiting"
                                    : undefined,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "teams_list_channels": {
            const teamIdArg = args.teamId as string | undefined;
            const teamId = teamIdArg || DEFAULT_TEAM_ID;

            if (!teamId) {
                throw new Error("teamId is required");
            }

            const result = await client.api(`/teams/${teamId}/channels`).get();
            const channels = result.value.map(
                (ch: {
                    id: string;
                    displayName: string;
                    description: string;
                    membershipType: string;
                    webUrl: string;
                }) => ({
                    id: ch.id,
                    displayName: ch.displayName,
                    description: ch.description,
                    membershipType: ch.membershipType,
                    webUrl: ch.webUrl,
                }),
            );

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({ channels }, null, 2),
                    },
                ],
            };
        }

        case "teams_list_teams": {
            // List groups that are teams
            const result = await client
                .api("/groups")
                .filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
                .select("id,displayName,description")
                .get();

            const teams = result.value.map(
                (team: { id: string; displayName: string; description: string }) => ({
                    id: team.id,
                    displayName: team.displayName,
                    description: team.description,
                }),
            );

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({ teams }, null, 2),
                    },
                ],
            };
        }

        default:
            throw new Error(`Unknown teams tool: ${name}`);
    }
}
