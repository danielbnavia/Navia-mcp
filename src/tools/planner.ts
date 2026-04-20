import type { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient } from "../lib/graph.js";

type ToolResult = Promise<{ content: { type: string; text: string }[] }>;

interface PlannerTask {
    id: string;
    title: string;
    bucketId: string;
    dueDateTime?: string;
    priority?: number;
    percentComplete: number;
    createdDateTime?: string;
    "@odata.etag"?: string;
}

interface PlannerTaskDetails {
    description?: string;
    checklist?: Record<string, unknown>;
}

interface PlannerBucket {
    id: string;
    name: string;
}

interface PlannerListResponse<T> {
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

// Batch processing constants
const BATCH_SIZE = 20; // Microsoft Graph batch limit
const BATCH_DELAY_MS = 1000; // Delay between batches to avoid rate limiting
const CONCURRENT_FETCHES = 10; // Parallel etag fetches

// Helper to delay execution
const delay = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

// Helper to chunk array into smaller arrays
function chunkArray<T>(array: T[], size: number): T[][] {
    const chunks: T[][] = [];
    for (let i = 0; i < array.length; i += size) {
        chunks.push(array.slice(i, i + size));
    }
    return chunks;
}

export const plannerTools: Tool[] = [
    {
        name: "planner_list_tasks",
        description: "List all tasks from the Email Tasks planner plan. Returns task titles, due dates, assignments, and completion status.",
        inputSchema: {
            type: "object",
            properties: {
                bucketId: {
                    type: "string",
                    description: "Optional: Filter by bucket ID. Leave empty to get all tasks.",
                },
                includeCompleted: {
                    type: "boolean",
                    description: "Include completed tasks. Default: false",
                },
            },
        },
    },
    {
        name: "planner_get_task",
        description: "Get details of a specific Planner task by ID, including checklist items and description.",
        inputSchema: {
            type: "object",
            properties: {
                taskId: {
                    type: "string",
                    description: "The Planner task ID",
                },
            },
            required: ["taskId"],
        },
    },
    {
        name: "planner_create_task",
        description: "Create a new task in the Email Tasks planner. Automatically assigns to the specified bucket.",
        inputSchema: {
            type: "object",
            properties: {
                title: {
                    type: "string",
                    description: "Task title",
                },
                bucketName: {
                    type: "string",
                    enum: ["VIP Follow-up", "Urgent", "Client Communications", "Operations", "Support", "Backlog"],
                    description: "Which bucket to put the task in",
                },
                dueDateTime: {
                    type: "string",
                    description: "Due date in ISO format (e.g., 2026-01-22T17:00:00Z)",
                },
                priority: {
                    type: "number",
                    enum: [1, 3, 5, 9],
                    description: "Priority: 1=Urgent, 3=Important, 5=Medium, 9=Low",
                },
                notes: {
                    type: "string",
                    description: "Task description/notes",
                },
            },
            required: ["title", "bucketName"],
        },
    },
    {
        name: "planner_update_task",
        description: "Update an existing Planner task (mark complete, change priority, update due date)",
        inputSchema: {
            type: "object",
            properties: {
                taskId: {
                    type: "string",
                    description: "The Planner task ID to update",
                },
                percentComplete: {
                    type: "number",
                    enum: [0, 50, 100],
                    description: "Completion: 0=Not started, 50=In progress, 100=Complete",
                },
                priority: {
                    type: "number",
                    enum: [1, 3, 5, 9],
                    description: "Priority: 1=Urgent, 3=Important, 5=Medium, 9=Low",
                },
                dueDateTime: {
                    type: "string",
                    description: "New due date in ISO format",
                },
            },
            required: ["taskId"],
        },
    },
    {
        name: "planner_list_buckets",
        description: "List all buckets in the Email Tasks plan with their IDs and task counts",
        inputSchema: {
            type: "object",
            properties: {},
        },
    },
    {
        name: "planner_delete_task",
        description: "Delete a single Planner task by ID",
        inputSchema: {
            type: "object",
            properties: {
                taskId: {
                    type: "string",
                    description: "The Planner task ID to delete",
                },
            },
            required: ["taskId"],
        },
    },
    {
        name: "planner_batch_delete_tasks",
        description: "Delete multiple Planner tasks in batches. Handles large numbers of tasks (100+) with proper batching and rate limiting to avoid timeouts.",
        inputSchema: {
            type: "object",
            properties: {
                taskIds: {
                    type: "array",
                    items: { type: "string" },
                    description: "Array of task IDs to delete",
                },
            },
            required: ["taskIds"],
        },
    },
    {
        name: "planner_batch_update_tasks",
        description: "Update multiple Planner tasks in batches (e.g., mark many as complete). Handles large numbers of tasks with proper batching.",
        inputSchema: {
            type: "object",
            properties: {
                taskIds: {
                    type: "array",
                    items: { type: "string" },
                    description: "Array of task IDs to update",
                },
                percentComplete: {
                    type: "number",
                    enum: [0, 50, 100],
                    description: "Completion: 0=Not started, 50=In progress, 100=Complete",
                },
                priority: {
                    type: "number",
                    enum: [1, 3, 5, 9],
                    description: "Priority: 1=Urgent, 3=Important, 5=Medium, 9=Low",
                },
            },
            required: ["taskIds"],
        },
    },
];

const BUCKET_MAP: Record<string, string> = {
    "VIP Follow-up": process.env.BUCKET_VIP || "MR_zXOcdqE6RIF6Yd6wPjMgAEaxq",
    Urgent: process.env.BUCKET_URGENT || "r1uZj-Zh7E63PTEPYqPmIsgANefB",
    "Client Communications": process.env.BUCKET_CLIENT || "RGDYN7r6oEWXkIA0L01o5sgAKG4S",
    Operations: process.env.BUCKET_OPERATIONS || "3hPTGvtl7kC5QGaKItSyDMgAD15d",
    Support: process.env.BUCKET_SUPPORT || "UBHJXpRD50qAPuqJODCDm8gAFBEB",
    Backlog: process.env.BUCKET_BACKLOG || "-ve7obEGl0GRm3nQC7ebN8gAPac6",
};

export async function handlePlannerTool(name: string, args: Record<string, unknown>): ToolResult {
    const client = await getGraphClient();
    const planId = process.env.PLANNER_PLAN_ID || "isN5lApe9UKwJzUcBXpiWsgABOwz";

    switch (name) {
        case "planner_list_tasks": {
            const tasks = (await client.api(`/planner/plans/${planId}/tasks`).get()) as PlannerListResponse<PlannerTask>;
            let filteredTasks = tasks.value;

            const bucketId = args.bucketId as string | undefined;
            if (bucketId) {
                filteredTasks = filteredTasks.filter((t: PlannerTask) => t.bucketId === bucketId);
            }

            if (!args.includeCompleted) {
                filteredTasks = filteredTasks.filter((t: PlannerTask) => t.percentComplete < 100);
            }

            const taskList = filteredTasks.map((t: PlannerTask) => ({
                id: t.id,
                title: t.title,
                bucketId: t.bucketId,
                dueDateTime: t.dueDateTime,
                priority: t.priority,
                percentComplete: t.percentComplete,
                createdDateTime: t.createdDateTime,
            }));

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({ count: taskList.length, tasks: taskList }, null, 2),
                    },
                ],
            };
        }

        case "planner_get_task": {
            const taskId = args.taskId as string;
            const task = (await client.api(`/planner/tasks/${taskId}`).get()) as PlannerTask;
            const details = (await client.api(`/planner/tasks/${taskId}/details`).get()) as PlannerTaskDetails;
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                id: task.id,
                                title: task.title,
                                dueDateTime: task.dueDateTime,
                                priority: task.priority,
                                percentComplete: task.percentComplete,
                                description: details.description,
                                checklist: details.checklist,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "planner_create_task": {
            const bucketName = args.bucketName as string;
            const bucketId = BUCKET_MAP[bucketName];
            if (!bucketId) {
                throw new Error(`Unknown bucket: ${bucketName}`);
            }

            const newTask: {
                planId: string;
                bucketId: string;
                title: string;
                dueDateTime?: string;
                priority?: number;
            } = {
                planId,
                bucketId,
                title: args.title as string,
            };

            if (args.dueDateTime) {
                newTask.dueDateTime = args.dueDateTime as string;
            }
            if (args.priority) {
                newTask.priority = args.priority as number;
            }

            const task = (await client.api("/planner/tasks").post(newTask)) as PlannerTask;

            // Add notes if provided
            if (args.notes) {
                const etag = task["@odata.etag"] as string;
                await client
                    .api(`/planner/tasks/${task.id}/details`)
                    .header("If-Match", etag)
                    .patch({ description: args.notes as string });
            }

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                taskId: task.id,
                                title: task.title,
                                bucket: bucketName,
                                message: `Task created successfully in ${bucketName}`,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "planner_update_task": {
            const taskId = args.taskId as string;

            // Get current task to get etag
            const currentTask = (await client.api(`/planner/tasks/${taskId}`).get()) as PlannerTask;
            const etag = currentTask["@odata.etag"] as string;
            const updates: { percentComplete?: number; priority?: number; dueDateTime?: string } = {};

            if (args.percentComplete !== undefined) updates.percentComplete = args.percentComplete as number;
            if (args.priority !== undefined) updates.priority = args.priority as number;
            if (args.dueDateTime) updates.dueDateTime = args.dueDateTime as string;

            await client
                .api(`/planner/tasks/${taskId}`)
                .header("If-Match", etag)
                .patch(updates);

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                taskId,
                                updates,
                                message: "Task updated successfully",
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "planner_list_buckets": {
            const buckets = (await client.api(`/planner/plans/${planId}/buckets`).get()) as PlannerListResponse<PlannerBucket>;
            const tasks = (await client.api(`/planner/plans/${planId}/tasks`).get()) as PlannerListResponse<PlannerTask>;
            const bucketList = buckets.value.map((b: PlannerBucket) => {
                const taskCount = tasks.value.filter((t: PlannerTask) => t.bucketId === b.id).length;
                const incompleteTasks = tasks.value.filter((t: PlannerTask) => t.bucketId === b.id && t.percentComplete < 100).length;
                return {
                    id: b.id,
                    name: b.name,
                    totalTasks: taskCount,
                    incompleteTasks,
                };
            });

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({ buckets: bucketList }, null, 2),
                    },
                ],
            };
        }

        case "planner_delete_task": {
            const taskId = args.taskId as string;

            // Get current task to get etag
            const currentTask = (await client.api(`/planner/tasks/${taskId}`).get()) as PlannerTask;
            const etag = currentTask["@odata.etag"] as string;

            await client
                .api(`/planner/tasks/${taskId}`)
                .header("If-Match", etag)
                .delete();

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: true,
                                taskId,
                                message: "Task deleted successfully",
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "planner_batch_delete_tasks": {
            const taskIds = args.taskIds as string[];
            const results: Array<{ taskId: string; success: boolean; error?: string }> = [];

            // Step 1: Fetch all etags in parallel batches
            console.log(`Fetching etags for ${taskIds.length} tasks...`);
            const taskEtags = new Map<string, string>();
            const fetchChunks = chunkArray(taskIds, CONCURRENT_FETCHES);

            for (let i = 0; i < fetchChunks.length; i++) {
                const chunk = fetchChunks[i];
                const fetchPromises = chunk.map(async (taskId: string) => {
                    try {
                        const task = (await client.api(`/planner/tasks/${taskId}`).get()) as PlannerTask;
                        return { taskId, etag: task["@odata.etag"] || null, success: true };
                    } catch (error: unknown) {
                        return { taskId, etag: null, success: false, error: error instanceof Error ? error.message : "Unknown error" };
                    }
                });

                const fetchResults = await Promise.all(fetchPromises);
                for (const result of fetchResults) {
                    if (result.success && result.etag) {
                        taskEtags.set(result.taskId, result.etag);
                    } else {
                        results.push({ taskId: result.taskId, success: false, error: result.error || "Failed to fetch etag" });
                    }
                }

                // Delay between fetch batches
                if (i < fetchChunks.length - 1) {
                    await delay(500);
                }
            }

            // Step 2: Delete tasks using Graph batch API (max 20 per batch)
            const tasksToDelete = Array.from(taskEtags.entries());
            const deleteChunks = chunkArray(tasksToDelete, BATCH_SIZE);
            console.log(`Deleting ${tasksToDelete.length} tasks in ${deleteChunks.length} batches...`);

            for (let i = 0; i < deleteChunks.length; i++) {
                const chunk = deleteChunks[i];

                // Build batch request
                const batchRequests = chunk.map(([taskId, etag], index: number) => ({
                    id: `${index}`,
                    method: "DELETE",
                    url: `/planner/tasks/${taskId}`,
                    headers: {
                        "If-Match": etag,
                    },
                }));

                try {
                    const batchResponse = (await client.api("/$batch").post({
                        requests: batchRequests,
                    })) as GraphBatchResponse;

                    // Process batch response
                    for (const response of batchResponse.responses) {
                        const taskId = chunk[parseInt(response.id, 10)][0];
                        if (response.status >= 200 && response.status < 300) {
                            results.push({ taskId, success: true });
                        } else {
                            results.push({
                                taskId,
                                success: false,
                                error: `HTTP ${response.status}: ${response.body?.error?.message || "Unknown error"}`,
                            });
                        }
                    }
                } catch (error: unknown) {
                    // If batch fails, mark all tasks in chunk as failed
                    for (const [taskId] of chunk) {
                        results.push({ taskId, success: false, error: error instanceof Error ? error.message : "Unknown error" });
                    }
                }

                // Progress log
                const completed = results.filter((r) => r.success).length;
                console.log(`Progress: ${completed}/${taskIds.length} tasks deleted (batch ${i + 1}/${deleteChunks.length})`);

                // Delay between delete batches
                if (i < deleteChunks.length - 1) {
                    await delay(BATCH_DELAY_MS);
                }
            }

            const successCount = results.filter((r) => r.success).length;
            const failCount = results.filter((r) => !r.success).length;

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: failCount === 0,
                                summary: {
                                    total: taskIds.length,
                                    deleted: successCount,
                                    failed: failCount,
                                },
                                results,
                                message: `Batch delete completed: ${successCount} deleted, ${failCount} failed`,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "planner_batch_update_tasks": {
            const taskIds = args.taskIds as string[];
            const updates: { percentComplete?: number; priority?: number } = {};

            if (args.percentComplete !== undefined) updates.percentComplete = args.percentComplete as number;
            if (args.priority !== undefined) updates.priority = args.priority as number;
            if (Object.keys(updates).length === 0) {
                throw new Error("No updates specified. Provide percentComplete or priority.");
            }

            const results: Array<{ taskId: string; success: boolean; error?: string }> = [];

            // Step 1: Fetch all etags in parallel batches
            console.log(`Fetching etags for ${taskIds.length} tasks...`);
            const taskEtags = new Map<string, string>();
            const fetchChunks = chunkArray(taskIds, CONCURRENT_FETCHES);

            for (let i = 0; i < fetchChunks.length; i++) {
                const chunk = fetchChunks[i];
                const fetchPromises = chunk.map(async (taskId: string) => {
                    try {
                        const task = (await client.api(`/planner/tasks/${taskId}`).get()) as PlannerTask;
                        return { taskId, etag: task["@odata.etag"] || null, success: true };
                    } catch (error: unknown) {
                        return { taskId, etag: null, success: false, error: error instanceof Error ? error.message : "Unknown error" };
                    }
                });

                const fetchResults = await Promise.all(fetchPromises);
                for (const result of fetchResults) {
                    if (result.success && result.etag) {
                        taskEtags.set(result.taskId, result.etag);
                    } else {
                        results.push({ taskId: result.taskId, success: false, error: result.error || "Failed to fetch etag" });
                    }
                }

                if (i < fetchChunks.length - 1) {
                    await delay(500);
                }
            }

            // Step 2: Update tasks using Graph batch API (max 20 per batch)
            const tasksToUpdate = Array.from(taskEtags.entries());
            const updateChunks = chunkArray(tasksToUpdate, BATCH_SIZE);
            console.log(`Updating ${tasksToUpdate.length} tasks in ${updateChunks.length} batches...`);

            for (let i = 0; i < updateChunks.length; i++) {
                const chunk = updateChunks[i];

                // Build batch request
                const batchRequests = chunk.map(([taskId, etag], index: number) => ({
                    id: `${index}`,
                    method: "PATCH",
                    url: `/planner/tasks/${taskId}`,
                    headers: {
                        "If-Match": etag,
                        "Content-Type": "application/json",
                    },
                    body: updates,
                }));

                try {
                    const batchResponse = (await client.api("/$batch").post({
                        requests: batchRequests,
                    })) as GraphBatchResponse;

                    // Process batch response
                    for (const response of batchResponse.responses) {
                        const taskId = chunk[parseInt(response.id, 10)][0];
                        if (response.status >= 200 && response.status < 300) {
                            results.push({ taskId, success: true });
                        } else {
                            results.push({
                                taskId,
                                success: false,
                                error: `HTTP ${response.status}: ${response.body?.error?.message || "Unknown error"}`,
                            });
                        }
                    }
                } catch (error: unknown) {
                    for (const [taskId] of chunk) {
                        results.push({ taskId, success: false, error: error instanceof Error ? error.message : "Unknown error" });
                    }
                }

                const completed = results.filter((r) => r.success).length;
                console.log(`Progress: ${completed}/${taskIds.length} tasks updated (batch ${i + 1}/${updateChunks.length})`);

                if (i < updateChunks.length - 1) {
                    await delay(BATCH_DELAY_MS);
                }
            }

            const successCount = results.filter((r) => r.success).length;
            const failCount = results.filter((r) => !r.success).length;

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                success: failCount === 0,
                                summary: {
                                    total: taskIds.length,
                                    updated: successCount,
                                    failed: failCount,
                                },
                                updates,
                                results,
                                message: `Batch update completed: ${successCount} updated, ${failCount} failed`,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        default:
            throw new Error(`Unknown planner tool: ${name}`);
    }
}
