import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import {
  CHANNEL_MESSAGE_KEYS,
  GET_BODY_MAX_CHARS,
  LIST_BODY_MAX_CHARS,
  hostedContentsContract,
  projectChannelMessage,
  withProjection,
} from "./chat-messages.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";

/**
 * Teams channel message tools (SM-12 part B): `list-channel-messages`, `get-channel-message`,
 * `list-channel-message-replies` and `list-channel-message-hosted-contents` (all v1.0 GET).
 *
 * Delegated permissions (current Microsoft docs, October 2026; see PERMISSIONS.md): all four pages
 * document `ChannelMessage.Read.All` as the least privileged delegated permission (work/school
 * only; `Group.Read.All`/`Group.ReadWrite.All` are higher), matching the catalog. The server never
 * pre-checks scopes; Graph decides.
 *
 * Offered options are typed only. The channel-message list documents `$top` (default 20, at most
 * 50) and `$expand` (replies); the replies list documents `$top` (at most 50) only; get message and
 * hosted contents document no OData at all. So: lists take a local `top` (1-100, default 20) fetched
 * in Graph pages of at most 50; there is NO `$filter`/`$orderby`/`$search` (the page says the other
 * OData parameters "aren't currently supported", so no typed time range exists, unlike chats) and
 * `$expand=replies` is deliberately not offered: expanded replies would bypass the per-message
 * projection and need their own paging (up to 200 per message by default, then
 * `replies@odata.nextLink`); callers use `list-channel-message-replies`.
 * Graph orders channel messages by the last modification of the whole reply chain.
 *
 * Every message is projected like a chat message (plain-text body capped at 2000 characters in
 * lists/replies and 20000 in get, allowlisted scalars, `from`, `mentions`, `attachments`,
 * `reactions`, `eventDetail` type only) with two deltas: `chatId` (always null for channels) is
 * dropped and `channelIdentity` (`teamId`, `channelId`) is kept; `summary` (a notification string
 * that duplicates the body) stays dropped. Hosted-content markers in a channel body are scoped to
 * the message's own team, channel and message (see `html-text.ts`). Hosted contents list references
 * only: there is no byte download (no `/$value` path is in the catalog).
 */

const listContract = (): QueryContract =>
  withProjection(
    defineQueryContract({
      // Graph: "up to 50 channel messages per page" (replies: "Maximum allowed $top value is 50").
      top: { default: 20, max: 100, pageMax: 50 },
      projectEntity: { keep: CHANNEL_MESSAGE_KEYS },
    }),
    true,
    (item) => projectChannelMessage(item, LIST_BODY_MAX_CHARS),
  );

const getContract = (): QueryContract =>
  withProjection(
    defineQueryContract({ projectEntity: { keep: CHANNEL_MESSAGE_KEYS } }),
    false,
    (item) => projectChannelMessage(item, GET_BODY_MAX_CHARS),
  );

export function registerChannelMessageTools(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const patternOf = (toolName: string): string => {
    const entry = catalog.find((e) => e.toolName === toolName);
    if (entry === undefined || entry.method !== "get" || entry["apiVersion"] === "beta") {
      throw new Error(`Catalog has no v1.0 GET entry for ${toolName}.`);
    }
    return entry.pathPattern;
  };
  for (const name of ["list-channel-messages", "list-channel-message-replies"]) {
    handlers.set(name, listTool(client, name, patternOf(name), listContract()));
  }
  handlers.set(
    "get-channel-message",
    getTool(client, "get-channel-message", patternOf("get-channel-message"), getContract()),
  );
  handlers.set(
    "list-channel-message-hosted-contents",
    listTool(
      client,
      "list-channel-message-hosted-contents",
      patternOf("list-channel-message-hosted-contents"),
      hostedContentsContract(),
    ),
  );
}
