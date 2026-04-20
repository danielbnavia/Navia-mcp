import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";

import "isomorphic-fetch";

let graphClient: Client | null = null;

export async function getGraphClient(): Promise<Client> {

    if (graphClient) {

        return graphClient;

    }

    const tenantId = process.env.AZURE_TENANT_ID;

    const clientId = process.env.AZURE_CLIENT_ID;

    const clientSecret = process.env.AZURE_CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {

        throw new Error("Missing Azure credentials. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET");

    }

    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

    graphClient = Client.initWithMiddleware({

        authProvider: {

            getAccessToken: async () => {

                const token = await credential.getToken("https://graph.microsoft.com/.default") as { token: string };

                return token.token;

            },

        },

    });

    return graphClient;

}

// For user-delegated auth (alternative approach)

export async function getGraphClientWithToken(accessToken: string): Promise<Client> {

    return Client.init({

        authProvider: (done) => {

            done(null, accessToken);

        },

    });

}
