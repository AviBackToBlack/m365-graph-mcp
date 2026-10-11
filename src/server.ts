import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import type { EndpointCatalog } from "./catalog.js";
import { errorResult } from "./envelope.js";
import {
  READ_ONLY_ANNOTATIONS,
  assertNoReservedInputNames,
  buildDescription,
  derivePathParamSchema,
  passThroughInputSchema,
  validateToolArgs,
  notImplementedHandler,
  type HandlerMap,
} from "./registry.js";

export const SERVER_NAME = "m365-graph-mcp";
export const SERVER_VERSION = "0.0.0-private";

/** Registers exactly the catalog tools; tools without a handler get a not_implemented stub. */
export function createServer(
  catalog: EndpointCatalog,
  handlers: HandlerMap = new Map(),
): McpServer {
  const server = new McpServer({ name: SERVER_NAME, version: SERVER_VERSION });

  for (const entry of catalog) {
    const impl = handlers.get(entry.toolName);
    const handler = impl?.handler ?? notImplementedHandler(entry.toolName);
    // Catalog-derived path params always win over a handler-supplied shape, so a domain
    // handler can add fields (query options) but can never drop or loosen a required ID.
    const inputSchema = { ...impl?.inputSchema, ...derivePathParamSchema(entry.pathPattern) };

    assertNoReservedInputNames(entry.toolName, inputSchema);

    server.registerTool(
      entry.toolName,
      {
        description: buildDescription(entry),
        inputSchema: passThroughInputSchema(inputSchema),
        annotations: { ...READ_ONLY_ANNOTATIONS },
      },
      async (rawArgs: unknown): Promise<CallToolResult> => {
        const validated = validateToolArgs(entry.toolName, inputSchema, rawArgs);
        if (!validated.ok) return validated.result;
        try {
          return await handler(validated.args);
        } catch {
          // Never surface raw error text: it could embed request details or tokens.
          return errorResult("graph_error", `Tool ${entry.toolName} failed unexpectedly.`);
        }
      },
    );
  }

  return server;
}
