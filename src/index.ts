#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { selectTokenProvider } from "./auth/select-token-provider.js";
import { loadEndpointCatalog } from "./catalog.js";
import { GraphClient } from "./graph/client.js";
import { createServer } from "./server.js";
import { createHandlers } from "./tools/index.js";

// stdout is the MCP protocol channel: never write to it; diagnostics go to stderr.
async function main(): Promise<void> {
  const catalog = await loadEndpointCatalog();
  // Explicit token source (D45): M365_GRAPH_TOKEN_SOURCE = manual (default; the manually supplied
  // short-lived token in M365_GRAPH_ACCESS_TOKEN) or bridge (the local Graph Explorer token
  // bridge). Invalid selection/config refuses to start; a missing/expired token or a stopped
  // bridge is read lazily per request and surfaces as reauth_required in the tool result.
  const client = new GraphClient({ tokenProvider: selectTokenProvider(), catalog });
  const server = createServer(catalog, createHandlers(client, catalog));
  try {
    await server.connect(new StdioServerTransport());
  } catch (error) {
    // Release any stdin listeners so the failed process can exit on its own.
    await server.close().catch(() => undefined);
    throw error;
  }
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : "unknown error";
  console.error(`m365-graph-mcp failed to start: ${message}`);
  // Not process.exit(): let stderr flush and the event loop drain naturally.
  process.exitCode = 1;
});
