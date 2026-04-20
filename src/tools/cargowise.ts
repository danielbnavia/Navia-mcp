import type { Tool } from "@modelcontextprotocol/sdk/types.js";

export const cargowiseTools: Tool[] = [
    {
        name: "cargowise_get_shipment",
        description: "Get shipment details from CargoWise by shipment number or reference",
        inputSchema: {
            type: "object",
            properties: {
                shipmentNumber: {
                    type: "string",
                    description: "CargoWise shipment number (e.g., SHP-12345)",
                },
                reference: {
                    type: "string",
                    description: "Alternative: Customer reference number",
                },
            },
        },
    },
    {
        name: "cargowise_get_order",
        description: "Get warehouse order details from CargoWise",
        inputSchema: {
            type: "object",
            properties: {
                orderNumber: {
                    type: "string",
                    description: "Order number (e.g., ORD-12345, SO-12345)",
                },
            },
            required: ["orderNumber"],
        },
    },
    {
        name: "cargowise_track_container",
        description: "Get container tracking information",
        inputSchema: {
            type: "object",
            properties: {
                containerNumber: {
                    type: "string",
                    description: "Container number (e.g., MSCU1234567)",
                },
            },
            required: ["containerNumber"],
        },
    },
    {
        name: "cargowise_get_inventory",
        description: "Get current inventory levels for a product/SKU",
        inputSchema: {
            type: "object",
            properties: {
                sku: {
                    type: "string",
                    description: "Product SKU",
                },
                clientCode: {
                    type: "string",
                    description: "Client/company code in CargoWise",
                },
            },
            required: ["sku"],
        },
    },
    {
        name: "cargowise_list_pending_orders",
        description: "List pending warehouse orders for a client",
        inputSchema: {
            type: "object",
            properties: {
                clientCode: {
                    type: "string",
                    description: "Client/company code in CargoWise",
                },
                status: {
                    type: "string",
                    enum: ["pending", "processing", "shipped", "all"],
                    description: "Filter by order status",
                },
                limit: {
                    type: "number",
                    description: "Maximum number of orders to return (default: 10)",
                },
            },
        },
    },
    {
        name: "cargowise_check_integration_status",
        description: "Check the status of CargoWise eAdaptor integration and recent API calls",
        inputSchema: {
            type: "object",
            properties: {
                includeErrors: {
                    type: "boolean",
                    description: "Include recent error details",
                },
            },
        },
    },
];

// CargoWise eAdaptor client
class CargoWiseClient {
    endpoint: string;
    username: string;
    password: string;

    constructor() {
        this.endpoint = process.env.CARGOWISE_ENDPOINT || "";
        this.username = process.env.CARGOWISE_USERNAME || "";
        this.password = process.env.CARGOWISE_PASSWORD || "";
    }

    isConfigured(): boolean {
        return !!(this.endpoint && this.username && this.password);
    }

    async query(xmlRequest: string): Promise<string> {
        if (!this.isConfigured()) {
            throw new Error("CargoWise eAdaptor not configured. Set CARGOWISE_ENDPOINT, CARGOWISE_USERNAME, CARGOWISE_PASSWORD");
        }

        const response = await fetch(this.endpoint, {
            method: "POST",
            headers: {
                "Content-Type": "application/xml",
                "Authorization": `Basic ${Buffer.from(`${this.username}:${this.password}`).toString("base64")}`,
            },
            body: xmlRequest,
        });

        if (!response.ok) {
            throw new Error(`CargoWise API error: ${response.status} ${response.statusText}`);
        }

        return response.text();
    }
}

const cwClient = new CargoWiseClient();

export async function handleCargowiseTool(
    name: string,
    args: Record<string, unknown>,
): Promise<{ content: { type: string; text: string }[] }> {
    // Check if CargoWise is configured
    if (!cwClient.isConfigured()) {
        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify(
                        {
                            error: "CargoWise eAdaptor not configured",
                            message: "Set CARGOWISE_ENDPOINT, CARGOWISE_USERNAME, and CARGOWISE_PASSWORD environment variables",
                            configured: false,
                        },
                        null,
                        2,
                    ),
                },
            ],
        };
    }

    switch (name) {
        case "cargowise_get_shipment": {
            const shipmentNumberArg = args.shipmentNumber as string | undefined;
            const referenceArg = args.reference as string | undefined;
            const shipmentNumber = shipmentNumberArg;
            const reference = referenceArg;

            if (!shipmentNumber && !reference) {
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({ error: "Provide either shipmentNumber or reference" }, null, 2),
                        },
                    ],
                };
            }

            // Build Universal XML request for shipment query
            const xmlRequest = `<?xml version="1.0" encoding="utf-8"?>

<UniversalShipment xmlns="http://www.cargowise.com/Schemas/Universal/2011/11">

  <ShipmentHeader>

    ${shipmentNumber ? `<ShipmentNumber>${shipmentNumber}</ShipmentNumber>` : ""}

    ${reference ? `<ShipmentReference>${reference}</ShipmentReference>` : ""}

  </ShipmentHeader>

</UniversalShipment>`;

            try {
                const response = await cwClient.query(xmlRequest);

                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify(
                                {
                                    success: true,
                                    shipmentNumber: shipmentNumber || reference,
                                    rawResponse: response.substring(0, 2000), // Truncate for readability
                                    note: "Full XML response available - parse as needed",
                                },
                                null,
                                2,
                            ),
                        },
                    ],
                };
            }
            catch (error) {
                const typedError = error as Error;
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify(
                                {
                                    success: false,
                                    error: typedError.message,
                                    shipmentNumber: shipmentNumber || reference,
                                },
                                null,
                                2,
                            ),
                        },
                    ],
                };
            }
        }

        case "cargowise_get_order": {
            const orderNumberArg = args.orderNumber as string;
            const orderNumber = orderNumberArg;

            const xmlRequest = `<?xml version="1.0" encoding="utf-8"?>

<UniversalEvent xmlns="http://www.cargowise.com/Schemas/Universal/2011/11">

  <Event>

    <EventType>ORD</EventType>

    <OrderNumber>${orderNumber}</OrderNumber>

  </Event>

</UniversalEvent>`;

            try {
                const response = await cwClient.query(xmlRequest);

                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify(
                                {
                                    success: true,
                                    orderNumber,
                                    rawResponse: response.substring(0, 2000),
                                },
                                null,
                                2,
                            ),
                        },
                    ],
                };
            }
            catch (error) {
                const typedError = error as Error;
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({ success: false, error: typedError.message, orderNumber }, null, 2),
                        },
                    ],
                };
            }
        }

        case "cargowise_track_container": {
            const containerNumberArg = args.containerNumber as string;
            const containerNumber = containerNumberArg;

            // Container tracking - would typically call shipping line APIs
            // For now, return a structured response indicating the lookup
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                containerNumber,
                                status: "lookup_pending",
                                message: "Container tracking requires integration with shipping line APIs (MSC, Maersk, etc.)",
                                suggestion: "Check CargoWise shipment linked to this container for latest status",
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "cargowise_get_inventory": {
            const skuArg = args.sku as string;
            const clientCodeArg = args.clientCode as string | undefined;
            const sku = skuArg;
            const clientCode = clientCodeArg;

            const xmlRequest = `<?xml version="1.0" encoding="utf-8"?>

<UniversalEvent xmlns="http://www.cargowise.com/Schemas/Universal/2011/11">

  <Event>

    <EventType>INV</EventType>

    <ProductCode>${sku}</ProductCode>

    ${clientCode ? `<ClientCode>${clientCode}</ClientCode>` : ""}

  </Event>

</UniversalEvent>`;

            try {
                const response = await cwClient.query(xmlRequest);

                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify(
                                {
                                    success: true,
                                    sku,
                                    clientCode: clientCode || "all",
                                    rawResponse: response.substring(0, 2000),
                                },
                                null,
                                2,
                            ),
                        },
                    ],
                };
            }
            catch (error) {
                const typedError = error as Error;
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({ success: false, error: typedError.message, sku }, null, 2),
                        },
                    ],
                };
            }
        }

        case "cargowise_list_pending_orders": {
            const clientCodeArg = args.clientCode as string | undefined;
            const statusArg = args.status as string | undefined;
            const limitArg = args.limit as number | undefined;

            const clientCode = clientCodeArg;
            const status = statusArg || "pending";
            const limit = limitArg || 10;

            // This would query CargoWise for pending orders
            // Returning structured placeholder
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                query: {
                                    clientCode: clientCode || "all",
                                    status,
                                    limit,
                                },
                                message: "Order list query requires specific CargoWise report configuration",
                                suggestion: "Use CargoWise reports or specific order number lookup",
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        case "cargowise_check_integration_status": {
            const includeErrorsArg = args.includeErrors as boolean | undefined;
            const includeErrors = includeErrorsArg;

            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(
                            {
                                integration: {
                                    name: "CargoWise eAdaptor",
                                    endpoint: process.env.CARGOWISE_ENDPOINT ? "Configured" : "Not configured",
                                    authenticated: cwClient.isConfigured(),
                                    status: cwClient.isConfigured() ? "ready" : "not_configured",
                                },
                                recentActivity: {
                                    note: "Activity logging not yet implemented",
                                    suggestion: "Check Power Automate flow run history for eAdaptor calls",
                                },
                                errors: includeErrors
                                    ? {
                                        note: "Error history not yet implemented",
                                        suggestion: "Check Integration Failure alerts in Teams",
                                    }
                                    : undefined,
                            },
                            null,
                            2,
                        ),
                    },
                ],
            };
        }

        default:
            throw new Error(`Unknown cargowise tool: ${name}`);
    }
}
