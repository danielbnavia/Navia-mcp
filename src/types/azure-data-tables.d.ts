declare module "@azure/data-tables" {
    export interface ListTableEntitiesOptions {
        queryOptions?: {
            filter?: string;
            select?: string[];
        };
    }

    export class TableClient {
        constructor(url: string, tableName: string, credential: unknown);
        createTable(): Promise<void>;
        upsertEntity<T extends Record<string, unknown>>(entity: T, mode?: "Merge" | "Replace"): Promise<void>;
        getEntity<T extends Record<string, unknown>>(partitionKey: string, rowKey: string): Promise<T>;
        deleteEntity(partitionKey: string, rowKey: string): Promise<void>;
        listEntities<T extends Record<string, unknown>>(options?: ListTableEntitiesOptions): AsyncIterable<T>;
    }
}
