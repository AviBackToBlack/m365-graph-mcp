import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { directoryCollectionTool } from "./directory.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";

/**
 * Teams chat tools (SM-11 part A): `list-chats`, `get-chat`, `list-chat-members` (all v1.0 GET).
 * Messages, replies, pinned messages and hosted contents are later parts.
 *
 * Delegated permissions (current Microsoft docs, October 2026; see PERMISSIONS.md): list and get
 * need `Chat.ReadBasic`; list members uses `Chat.ReadBasic` too, its documented least privilege
 * (D18). None supports personal Microsoft accounts.
 *
 * Graph documents NO `$select` for any of the three (list: only `$expand`, `$top` max 50,
 * `$filter`, `$orderby`; get and members: "doesn't support OData query parameters"), so none is
 * ever sent and every response is projected locally to an allowlist.
 *
 * `list-chats` offers `top` (50 per request is Graph's documented cap; more is reached by
 * following validated pages), a typed `expand` of exactly `lastMessagePreview`, and an optional
 * `orderby` restricted, when provided, to `lastMessagePreview/createdDateTime` descending (the only
 * documented order; omitted, no `$orderby` is sent). Not
 * offered: `$filter` (v1.0 documents only an installedApps filter, no chat property),
 * `$expand=members` (Graph silently returns at most 25 members per chat; use
 * `list-chat-members`), `$search`, `$skip`, `$count`.
 *
 * `lastMessagePreview` is message CONTENT, so only metadata is exposed, and only when the caller
 * asked for the expansion: id, createdDateTime, isDeleted, messageType and the sender's id and
 * display name. Its HTML `body` and `eventDetail` are dropped (body handling is part B).
 * `onlineMeetingInfo` (carries the meeting join URL) is dropped too.
 *
 * `list-chat-members` offers a local `top` only (no `$top` is sent) and projects per
 * `@odata.type`; only `aadUserConversationMember` keeps its details.
 */

const AAD_USER_MEMBER = "#microsoft.graph.aadUserConversationMember";

/** Top-level chat properties that survive projection (scalars only; nested ones have projectors). */
const CHAT_SCALARS = [
  "id",
  "topic",
  "createdDateTime",
  "lastUpdatedDateTime",
  "chatType",
  "tenantId",
  "webUrl",
  "isHiddenForAllMembers",
] as const;
const CHAT_KEYS = [...CHAT_SCALARS, "viewpoint"] as const;
const PREVIEW_KEY = "lastMessagePreview";

type Rec = Record<string, unknown>;

export const isRecord = (v: unknown): v is Rec =>
  typeof v === "object" && v !== null && !Array.isArray(v);
const isScalar = (v: unknown): v is string | number | boolean | null =>
  v === null || ["string", "number", "boolean"].includes(typeof v);

/** Copies only the named keys whose value is a scalar (never an object or array). */
export function pickScalars(obj: Rec, keys: readonly string[]): Rec {
  const out: Rec = {};
  for (const k of keys) {
    const v = obj[k];
    if (Object.hasOwn(obj, k) && isScalar(v)) out[k] = v;
  }
  return out;
}

export function projectSender(from: unknown): unknown {
  if (from === null) return null;
  if (!isRecord(from)) return undefined;
  const out: Rec = {};
  for (const k of ["user", "application"]) {
    const v = from[k];
    if (v === null) out[k] = null;
    else if (isRecord(v)) out[k] = pickScalars(v, ["id", "displayName"]);
  }
  return out;
}

/** Metadata-only message preview: no body, no eventDetail, no attachments. */
function projectPreview(preview: unknown): unknown {
  if (preview === null) return null;
  if (!isRecord(preview)) return undefined;
  const out = pickScalars(preview, ["id", "createdDateTime", "isDeleted", "messageType"]);
  const from = projectSender(preview["from"]);
  if (from !== undefined) out["from"] = from;
  return out;
}

/**
 * Keeps the chat allowlist; `lastMessagePreview` only when `withPreview` (the request asked for
 * the expansion). Graph-supplied data is untrusted: nothing else survives. Non-objects yield
 * undefined.
 */
export function projectChat(item: unknown, withPreview: boolean): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, CHAT_SCALARS);
  const viewpoint = item["viewpoint"];
  if (viewpoint === null) out["viewpoint"] = null;
  else if (isRecord(viewpoint)) {
    out["viewpoint"] = pickScalars(viewpoint, ["isHidden", "lastMessageReadDateTime"]);
  }
  if (withPreview && Object.hasOwn(item, PREVIEW_KEY)) {
    const preview = projectPreview(item[PREVIEW_KEY]);
    if (preview !== undefined) out[PREVIEW_KEY] = preview;
  }
  return out;
}

const stringArray = (v: unknown): string[] | undefined =>
  Array.isArray(v) ? v.filter((x): x is string => typeof x === "string") : undefined;

/**
 * Per-`@odata.type` member projection (fail closed). `aadUserConversationMember` keeps id,
 * `@odata.type`, displayName, email, userId, tenantId, roles and visibleHistoryStartDateTime;
 * any other type (anonymous guest, Skype, Microsoft account, ACS, ...) keeps `id` and
 * `@odata.type` only; a missing type keeps `id` only. Non-objects yield undefined.
 */
export function projectChatMember(item: unknown): unknown {
  if (!isRecord(item)) return undefined;
  const type = item["@odata.type"];
  if (typeof type !== "string") return pickScalars(item, ["id"]);
  if (type !== AAD_USER_MEMBER) return pickScalars(item, ["id", "@odata.type"]);
  const out = pickScalars(item, [
    "id",
    "@odata.type",
    "displayName",
    "email",
    "userId",
    "tenantId",
    "visibleHistoryStartDateTime",
  ]);
  const roles = stringArray(item["roles"]);
  if (roles !== undefined) out["roles"] = roles;
  return out;
}

const wantsPreview = (query: Readonly<Record<string, unknown>>): boolean => {
  const expand = query["$expand"];
  return typeof expand === "string" && expand.split(",").includes(PREVIEW_KEY);
};

/**
 * Routes every projection path of `base` (single entity, collection item, batch body) through
 * `projectChat`, with `lastMessagePreview` kept only when THAT request's `$expand` asked for it.
 * `base`'s own projection still runs first in a batch (it clips to the subrequest's `$top`).
 */
function withChatProjection(base: QueryContract, collection: boolean): QueryContract {
  const inner = base.batchPolicy.project;
  return {
    ...base,
    projectItem: (item) => projectChat(item, false),
    projectResponse: (item, query) => projectChat(item, wantsPreview(query)),
    itemProjector: (query) => {
      const preview = wantsPreview(query);
      return (item) => projectChat(item, preview);
    },
    batchPolicy: {
      check: base.batchPolicy.check,
      project: (body, pairs) => {
        const preview = wantsPreview(Object.fromEntries(pairs));
        const out = inner === undefined ? body : inner(body, pairs);
        if (!collection) return projectChat(out, preview) ?? {};
        if (!isRecord(out) || !Array.isArray(out["value"])) return {};
        const value = out["value"]
          .map((i) => projectChat(i, preview))
          .filter((i) => i !== undefined);
        return { ...out, value };
      },
    },
  };
}

const listContract = (): QueryContract =>
  withChatProjection(
    defineQueryContract({
      // Graph: "Maximum allowed $top value is 50" per request; more is reached by paging.
      top: { default: 25, max: 100, pageMax: 50 },
      expand: { allowed: [PREVIEW_KEY] },
      orderby: {
        allowed: ["lastMessagePreview/createdDateTime"],
        maxClauses: 1,
        requireInFilter: false,
        descOnly: true,
      },
      projectEntity: { keep: CHAT_KEYS },
    }),
    true,
  );

const getContract = (): QueryContract =>
  withChatProjection(defineQueryContract({ projectEntity: { keep: CHAT_KEYS } }), false);

const membersContract = (): QueryContract =>
  // No `$top` on the wire (the page documents no OData options): `top` is a local limit.
  defineQueryContract({ top: { default: 100, max: 250, wire: false } });

export function registerChatTools(
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
  handlers.set(
    "list-chats",
    listTool(client, "list-chats", patternOf("list-chats"), listContract()),
  );
  handlers.set("get-chat", getTool(client, "get-chat", patternOf("get-chat"), getContract()));
  handlers.set(
    "list-chat-members",
    directoryCollectionTool(
      client,
      "list-chat-members",
      patternOf("list-chat-members"),
      membersContract(),
      projectChatMember,
    ),
  );
}
