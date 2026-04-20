import type { Tool } from "@modelcontextprotocol/sdk/types.js";
import { deleteMemory, listMemoryKeys, recallMemory, storeMemory } from "../services/memory.js";

type ToolResult = Promise<{ content: { type: string; text: string }[] }>;

const USER_ID = "danielb@naviafreight.com";

export const memoryTools: Tool[] = [
    {
        name: "memory_store",
        description: "Store a memory for the current user",
        inputSchema: {
            type: "object",
            properties: {
                key: {
                    type: "string",
                    description: "Memory key",
                },
                value: {
                    type: "string",
                    description: "Memory value",
                },
                category: {
                    type: "string",
                    enum: ["preference", "context", "history", "decision"],
                    description: "Memory category",
                },
            },
            required: ["key", "value", "category"],
        },
    },
    {
        name: "memory_recall",
        description: "Recall memory by key or category",
        inputSchema: {
            type: "object",
            properties: {
                key: {
                    type: "string",
                    description: "Optional memory key",
                },
                category: {
                    type: "string",
                    enum: ["preference", "context", "history", "decision"],
                    description: "Optional memory category",
                },
            },
        },
    },
    {
        name: "memory_delete",
        description: "Delete a memory by key",
        inputSchema: {
            type: "object",
            properties: {
                key: {
                    type: "string",
                    description: "Memory key",
                },
            },
            required: ["key"],
        },
    },
    {
        name: "memory_list_keys",
        description: "List all memory keys for the current user",
        inputSchema: {
            type: "object",
            properties: {},
        },
    },
];

export async function handleMemoryTool(toolName: string, args: Record<string, unknown>): ToolResult {
    switch (toolName) {
        case "memory_store": {
            const key = args.key as string;
            const value = args.value as string;
            const category = args.category as string;
            await storeMemory(USER_ID, key, value, category);
            const result = {
                success: true,
                key,
                category,
            };

            return {
                content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
            };
        }

        case "memory_recall": {
            const key = args.key as string | undefined;
            const category = args.category as string | undefined;
            const memories = await recallMemory(USER_ID, key, category);
            const result = {
                count: memories.length,
                memories,
            };

            return {
                content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
            };
        }

        case "memory_delete": {
            const key = args.key as string;
            await deleteMemory(USER_ID, key);
            const result = {
                success: true,
                key,
            };

            return {
                content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
            };
        }

        case "memory_list_keys": {
            const keys = await listMemoryKeys(USER_ID);
            const result = {
                count: keys.length,
                keys,
            };

            return {
                content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
            };
        }

        default:
            throw new Error(`Unknown memory tool: ${toolName}`);
    }
}
