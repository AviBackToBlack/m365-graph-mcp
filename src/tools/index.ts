import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { HandlerMap, ToolImplementation } from "../registry.js";
import { batchTargets } from "../graph/batch.js";
import { createBatchTool } from "./batch.js";
import { registerCalendarAvailabilityTools } from "./calendar-availability.js";
import { registerCalendarTools } from "./calendar.js";
import { registerCalendarDeltaTool } from "./calendar-delta.js";
import { registerCalendarEventsDeltaTool } from "./calendar-events-delta.js";
import { registerChatMessageTools } from "./chat-messages.js";
import { registerChannelMessageTools } from "./channel-messages.js";
import { registerChannelTabTools } from "./channel-tabs.js";
import { registerChatTools } from "./chats.js";
import { registerDirectoryTools } from "./directory.js";
import { registerGroupTools } from "./groups.js";
import { registerIdentityTools } from "./identity.js";
import { registerMailTools } from "./mail.js";
import { registerMailDeltaTool } from "./mail-delta.js";
import { registerPeopleTools } from "./people.js";
import { registerPresenceInsightTools } from "./presence-insights.js";
import { registerCopilotRetrieveTool } from "./copilot-retrieve.js";
import { registerSearchQueryTool } from "./search-query.js";
import { registerSharePointDriveTools } from "./sharepoint-drives.js";
import { registerSharePointListTools } from "./sharepoint-lists.js";
import { registerSharePointSiteTools } from "./sharepoint-sites.js";
import { registerTeamTools } from "./teams.js";

/** Handlers keyed by catalog `toolName`. Domain units add their tools here. */
export function createHandlers(client: GraphClient, catalog: EndpointCatalog): HandlerMap {
  const handlers = new Map<string, ToolImplementation>();
  registerMailTools(handlers, client, catalog);
  registerMailDeltaTool(handlers, client, catalog);
  registerCalendarTools(handlers, client, catalog);
  registerCalendarDeltaTool(handlers, client, catalog);
  registerCalendarEventsDeltaTool(handlers, client, catalog);
  registerCalendarAvailabilityTools(handlers, client, catalog);
  registerIdentityTools(handlers, client, catalog);
  registerDirectoryTools(handlers, client, catalog);
  registerPeopleTools(handlers, client, catalog);
  registerGroupTools(handlers, client, catalog);
  registerChatTools(handlers, client, catalog);
  registerChatMessageTools(handlers, client, catalog);
  registerTeamTools(handlers, client, catalog);
  registerChannelMessageTools(handlers, client, catalog);
  registerChannelTabTools(handlers, client, catalog);
  registerSharePointSiteTools(handlers, client, catalog);
  registerSharePointDriveTools(handlers, client, catalog);
  registerSharePointListTools(handlers, client, catalog);
  registerPresenceInsightTools(handlers, client, catalog);
  registerSearchQueryTool(handlers, client, catalog);
  registerCopilotRetrieveTool(handlers, client, catalog);
  // Domain handlers (GET and retrieval-only POST) are registered above this line. Batch is registered last and may
  // only target implemented GET tools that declare a batchPolicy, so it never widens
  // the surface or bypasses a tool's query limits.
  handlers.set("graph-batch", createBatchTool(client, catalog, batchTargets(catalog, handlers)));
  return handlers;
}
