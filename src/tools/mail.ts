import type { EndpointCatalog } from "../catalog.js";
import { errorResult, successResult } from "../envelope.js";
import { GraphClientError } from "../graph/errors.js";
import type { GraphClient, GraphRequest } from "../graph/client.js";
import { pathParamNames } from "../graph/url.js";
import type { ToolArgs, ToolImplementation } from "../registry.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";

/**
 * Mail core tools (SM-5). Query options, defaults and limits live in the per-tool contracts
 * below; the same contract is each tool's `batchPolicy`.
 *
 * Collection tools return a bounded list plus `meta.hasMore`; they never hand the caller a
 * Graph `@odata.nextLink` (a link could carry query options outside the contract). More
 * results are reached by narrowing the query or, for non-search listings, by `skip`. Message
 * listings follow pages internally up to `top` (max 1000, so every `$search` match is
 * reachable) while each request asks for at most 250.
 */

// Properties documented on the Graph v1.0 `message` resource that are safe to return.
// Bodies (`body`, `uniqueBody`) are deliberately absent from the list-tool allowlist; use
// get-mail-message. `internetMessageHeaders` is excluded as bulky and rarely needed.
export const MESSAGE_METADATA = [
  "id",
  "subject",
  "bodyPreview",
  "from",
  "sender",
  "toRecipients",
  "ccRecipients",
  "bccRecipients",
  "replyTo",
  "receivedDateTime",
  "sentDateTime",
  "createdDateTime",
  "lastModifiedDateTime",
  "hasAttachments",
  "importance",
  "isRead",
  "isDraft",
  "conversationId",
  "parentFolderId",
  "internetMessageId",
  "webLink",
  "categories",
  "flag",
  "inferenceClassification",
] as const;

const MESSAGE_ORDERBY = [
  "receivedDateTime",
  "sentDateTime",
  "createdDateTime",
  "lastModifiedDateTime",
  "subject",
  "importance",
  "from/emailAddress/address",
  "sender/emailAddress/address",
] as const;

const FOLDER_PROPERTIES = [
  "id",
  "displayName",
  "parentFolderId",
  "childFolderCount",
  "unreadItemCount",
  "totalItemCount",
  "isHidden",
] as const;

/** Attachment metadata only. `contentBytes`/`contentLocation`-style content is never allowed. */
const ATTACHMENT_PROPERTIES = [
  "id",
  "name",
  "contentType",
  "size",
  "isInline",
  "lastModifiedDateTime",
] as const;

export const LIST_DEFAULT_SELECT = [
  "id",
  "subject",
  "from",
  "toRecipients",
  "receivedDateTime",
  "bodyPreview",
  "isRead",
  "hasAttachments",
] as const;

const GET_DEFAULT_SELECT = [
  "id",
  "subject",
  "from",
  "sender",
  "toRecipients",
  "ccRecipients",
  "bccRecipients",
  "replyTo",
  "receivedDateTime",
  "sentDateTime",
  "hasAttachments",
  "importance",
  "isRead",
  "conversationId",
  "parentFolderId",
  "internetMessageId",
  "webLink",
  "categories",
  "body",
] as const;

// $top: Graph allows 1..1000 for messages but warns large pages can 504, and responses are
// capped locally (5 MB); 250 compact items per page stays well inside both. `top` may reach
// 1000 (Graph's $search result cap) by following pages in the single-request path; batch is
// bound to the 250 page size. Folders and attachments document no maximum, so the limits are
// project choices.
function messageListContract(): QueryContract {
  return defineQueryContract({
    top: { default: 10, max: 1000, pageMax: 250 },
    select: { allowed: MESSAGE_METADATA, default: LIST_DEFAULT_SELECT, requiredInBatch: true },
    filter: { maxLength: 1000 },
    search: { maxLength: 500 },
    orderby: { allowed: MESSAGE_ORDERBY, maxClauses: 3 },
    skip: { max: 10_000 },
  });
}

const folderContract = (): QueryContract =>
  defineQueryContract({
    top: { default: 25, max: 250 },
    select: {
      allowed: FOLDER_PROPERTIES,
      default: FOLDER_PROPERTIES,
      requiredInBatch: false,
    },
    filter: { maxLength: 500 },
    skip: { max: 10_000 },
    includeHiddenFolders: true,
  });

const getMessageContract = (): QueryContract =>
  defineQueryContract({
    select: {
      allowed: [...MESSAGE_METADATA, "body", "uniqueBody"],
      default: GET_DEFAULT_SELECT,
      requiredInBatch: false,
    },
    bodyType: { default: "text" },
  });

const attachmentContract = (): QueryContract =>
  defineQueryContract({
    top: { default: 25, max: 100 },
    select: {
      allowed: ATTACHMENT_PROPERTIES,
      default: ATTACHMENT_PROPERTIES,
      // An unselected attachment list embeds contentBytes, and batch cannot reshape bodies.
      requiredInBatch: true,
    },
  });

export function pathParams(pathPattern: string, args: ToolArgs): Record<string, string> {
  const out: Record<string, string> = {};
  for (const name of pathParamNames(pathPattern)) {
    const value = args[name];
    if (typeof value === "string") out[name] = value;
  }
  return out;
}

export function failure(toolName: string, err: unknown): ReturnType<typeof errorResult> {
  if (err instanceof GraphClientError) return err.toResult();
  return errorResult("graph_error", `Tool ${toolName} failed unexpectedly.`);
}

export function listTool(
  client: GraphClient,
  toolName: string,
  pathPattern: string,
  contract: QueryContract,
): ToolImplementation {
  return {
    inputSchema: contract.shape,
    // The contract's projector also runs on batch sub-responses (collection level), so
    // unselected content (bodies, contentBytes) never leaves either path.
    batchPolicy: contract.batchPolicy,
    handler: async (args) => {
      try {
        const { query, headers, limit } = contract.toRequest(args);
        const base: GraphRequest = {
          pathPattern,
          pathParams: pathParams(pathPattern, args),
          query,
          headers,
        };
        // Follow-up pages go through GraphClient.fetchLink bound to `base`; `maxItems` is the
        // caller's (bounded) limit, and the client's page ceiling applies.
        const result = await client.paginate(base, { maxItems: limit });
        const project = contract.itemProjector(query);
        const items = result.items.map(project).filter((item) => item !== undefined);
        return successResult(items, {
          count: items.length,
          hasMore: result.truncated,
          pages: result.pages,
        });
      } catch (err) {
        return failure(toolName, err);
      }
    },
  };
}

export function getTool(
  client: GraphClient,
  toolName: string,
  pathPattern: string,
  contract: QueryContract,
): ToolImplementation {
  return {
    inputSchema: contract.shape,
    batchPolicy: contract.batchPolicy,
    handler: async (args) => {
      try {
        const { query, headers } = contract.toRequest(args);
        const res = await client.request({
          pathPattern,
          pathParams: pathParams(pathPattern, args),
          query,
          headers,
        });
        const meta: Record<string, unknown> = { clientRequestId: res.clientRequestId };
        if (res.requestId !== undefined) meta["requestId"] = res.requestId;
        // Identity unless the contract reshapes the entity (calendar extended properties).
        const projected = contract.projectResponse(res.data, query);
        if (projected === undefined && contract.projectsEntity) {
          // A strict entity contract never passes a non-object (array, string, null) through.
          throw new GraphClientError("graph_error", "Graph returned an unexpected response shape.");
        }
        return successResult(projected ?? null, meta);
      } catch (err) {
        return failure(toolName, err);
      }
    },
  };
}

/**
 * Adds the SM-5 mail tools and the SM-6 shared-mailbox tools. Each is batch-reachable only
 * through its own query contract. The shared tools reuse the personal-mailbox contracts
 * unchanged (same limits, `$select` allowlists, search/body rules; attachment-free); the only
 * difference is the catalog path, whose `{user-id}` is validated as a GUID or UPN.
 */
export function registerMailTools(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const patternOf = (toolName: string): string => {
    const entry = catalog.find((e) => e.toolName === toolName);
    if (entry === undefined || entry.method !== "get") {
      throw new Error(`Catalog has no GET entry for ${toolName}.`);
    }
    return entry.pathPattern;
  };
  const list = (toolName: string, contract: QueryContract) =>
    handlers.set(toolName, listTool(client, toolName, patternOf(toolName), contract));

  list("list-mail-messages", messageListContract());
  list("list-mail-folders", folderContract());
  list("list-mail-child-folders", folderContract());
  list("list-mail-folder-messages", messageListContract());
  handlers.set(
    "get-mail-message",
    getTool(client, "get-mail-message", patternOf("get-mail-message"), getMessageContract()),
  );
  list("list-mail-attachments", attachmentContract());
  list("list-shared-mailbox-messages", messageListContract());
  list("list-shared-mailbox-folder-messages", messageListContract());
  handlers.set(
    "get-shared-mailbox-message",
    getTool(
      client,
      "get-shared-mailbox-message",
      patternOf("get-shared-mailbox-message"),
      getMessageContract(),
    ),
  );
}
