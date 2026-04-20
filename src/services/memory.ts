import { DefaultAzureCredential } from "@azure/identity";
import { TableClient } from "@azure/data-tables";

export interface MemoryRecord {
    partitionKey: string;
    rowKey: string;
    value: string;
    category: "preference" | "context" | "history" | "decision";
    updatedAt: string;
}

const TABLE_NAME = "agentmemory";
const STORAGE_ACCOUNT = process.env.AZURE_STORAGE_ACCOUNT || "naviamcpstorage";
const USER_CATEGORIES = new Set(["preference", "context", "history", "decision"]);
const MAX_VALUE_LENGTH = 10000;

const inMemoryStore = new Map<string, Map<string, MemoryRecord>>();

let tableClient: TableClient | null = null;
let tableReadyPromise: Promise<void> | null = null;
let warnedFallback = false;
let useInMemoryFallback = false;

function isNotFoundError(error: unknown): boolean {
    if (!error || typeof error !== "object") {
        return false;
    }

    const maybeStatusCode = (error as { statusCode?: number }).statusCode;
    return maybeStatusCode === 404;
}

function isAlreadyExistsError(error: unknown): boolean {
    if (!error || typeof error !== "object") {
        return false;
    }

    const maybeStatusCode = (error as { statusCode?: number }).statusCode;
    return maybeStatusCode === 409;
}

function warnFallbackOnce(reason: unknown): void {
    if (warnedFallback) {
        return;
    }

    warnedFallback = true;
    const message = reason instanceof Error ? reason.message : String(reason);
    console.warn(`[memory] Azure Table Storage unavailable, using in-memory fallback: ${message}`);
}

function normalizeCategory(category: string): MemoryRecord["category"] {
    const normalized = category.toLowerCase();
    if (USER_CATEGORIES.has(normalized)) {
        return normalized as MemoryRecord["category"];
    }
    return "context";
}

function normalizeValue(value: string): string {
    if (value.length <= MAX_VALUE_LENGTH) {
        return value;
    }
    return value.slice(0, MAX_VALUE_LENGTH);
}

function getUserMemories(userId: string): Map<string, MemoryRecord> {
    const existing = inMemoryStore.get(userId);
    if (existing) {
        return existing;
    }

    const created = new Map<string, MemoryRecord>();
    inMemoryStore.set(userId, created);
    return created;
}

async function ensureTableClient(): Promise<TableClient | null> {
    if (useInMemoryFallback) {
        return null;
    }

    if (tableClient) {
        return tableClient;
    }

    if (!tableReadyPromise) {
        tableReadyPromise = (async () => {
            try {
                const endpoint = `https://${STORAGE_ACCOUNT}.table.core.windows.net`;
                const credential = new DefaultAzureCredential();
                const client = new TableClient(endpoint, TABLE_NAME, credential);
                try {
                    await client.createTable();
                }
                catch (error) {
                    if (!isAlreadyExistsError(error)) {
                        throw error;
                    }
                }
                tableClient = client;
            }
            catch (error) {
                useInMemoryFallback = true;
                warnFallbackOnce(error);
            }
        })();
    }

    await tableReadyPromise;
    return tableClient;
}

function toMemoryRecord(entity: Record<string, unknown>): MemoryRecord {
    return {
        partitionKey: String(entity.partitionKey || ""),
        rowKey: String(entity.rowKey || ""),
        value: String(entity.value || ""),
        category: normalizeCategory(String(entity.category || "context")),
        updatedAt: String(entity.updatedAt || new Date().toISOString()),
    };
}

async function storeMemoryInFallback(userId: string, key: string, value: string, category: string): Promise<void> {
    const userMemories = getUserMemories(userId);
    userMemories.set(key, {
        partitionKey: userId,
        rowKey: key,
        value: normalizeValue(value),
        category: normalizeCategory(category),
        updatedAt: new Date().toISOString(),
    });
}

async function recallMemoryFromFallback(userId: string, key?: string, category?: string): Promise<MemoryRecord[]> {
    const userMemories = getUserMemories(userId);
    const normalizedCategory = category ? normalizeCategory(category) : undefined;

    if (key) {
        const found = userMemories.get(key);
        if (!found) {
            return [];
        }
        if (normalizedCategory && found.category !== normalizedCategory) {
            return [];
        }
        return [found];
    }

    const values = Array.from(userMemories.values());
    if (!normalizedCategory) {
        return values;
    }

    return values.filter((memory) => memory.category === normalizedCategory);
}

async function deleteMemoryFromFallback(userId: string, key: string): Promise<void> {
    const userMemories = getUserMemories(userId);
    userMemories.delete(key);
}

async function listMemoryKeysFromFallback(userId: string): Promise<string[]> {
    const userMemories = getUserMemories(userId);
    return Array.from(userMemories.keys());
}

export async function storeMemory(userId: string, key: string, value: string, category: string): Promise<void> {
    const client = await ensureTableClient();

    if (!client) {
        await storeMemoryInFallback(userId, key, value, category);
        return;
    }

    const now = new Date().toISOString();
    // TableClient expects a generic record-like entity; our MemoryRecord type is
    // stricter and intentionally does not have an index signature.
    const entity: Record<string, unknown> = {
        partitionKey: userId,
        rowKey: key,
        value: normalizeValue(value),
        category: normalizeCategory(category),
        updatedAt: now,
    };

    try {
        await client.upsertEntity(entity, "Merge");
    }
    catch (error) {
        useInMemoryFallback = true;
        warnFallbackOnce(error);
        await storeMemoryInFallback(userId, key, value, category);
    }
}

export async function recallMemory(userId: string, key?: string, category?: string): Promise<MemoryRecord[]> {
    const client = await ensureTableClient();

    if (!client) {
        return recallMemoryFromFallback(userId, key, category);
    }

    const normalizedCategory = category ? normalizeCategory(category) : undefined;

    try {
        if (key) {
            const entity = await client.getEntity<Record<string, unknown>>(userId, key);
            const memory = toMemoryRecord(entity);
            if (normalizedCategory && memory.category !== normalizedCategory) {
                return [];
            }
            return [memory];
        }

        const filter = normalizedCategory
            ? `partitionKey eq '${userId.replace(/'/g, "''")}' and category eq '${normalizedCategory}'`
            : `partitionKey eq '${userId.replace(/'/g, "''")}'`;

        const memories: MemoryRecord[] = [];
        const entities = client.listEntities<Record<string, unknown>>({
            queryOptions: {
                filter,
            },
        });

        for await (const entity of entities) {
            memories.push(toMemoryRecord(entity));
        }

        return memories;
    }
    catch (error) {
        if (key && isNotFoundError(error)) {
            return [];
        }

        useInMemoryFallback = true;
        warnFallbackOnce(error);
        return recallMemoryFromFallback(userId, key, category);
    }
}

export async function deleteMemory(userId: string, key: string): Promise<void> {
    const client = await ensureTableClient();

    if (!client) {
        await deleteMemoryFromFallback(userId, key);
        return;
    }

    try {
        await client.deleteEntity(userId, key);
    }
    catch (error) {
        if (isNotFoundError(error)) {
            return;
        }

        useInMemoryFallback = true;
        warnFallbackOnce(error);
        await deleteMemoryFromFallback(userId, key);
    }
}

export async function listMemoryKeys(userId: string): Promise<string[]> {
    const client = await ensureTableClient();

    if (!client) {
        return listMemoryKeysFromFallback(userId);
    }

    try {
        const entities = client.listEntities<Record<string, unknown>>({
            queryOptions: {
                filter: `partitionKey eq '${userId.replace(/'/g, "''")}'`,
                select: ["rowKey"],
            },
        });

        const keys: string[] = [];
        for await (const entity of entities) {
            keys.push(String(entity.rowKey || ""));
        }

        return keys;
    }
    catch (error) {
        useInMemoryFallback = true;
        warnFallbackOnce(error);
        return listMemoryKeysFromFallback(userId);
    }
}
