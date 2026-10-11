import { z } from "zod";
import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import { GraphClientError } from "../graph/errors.js";
import type { ToolImplementation } from "../registry.js";
import { isRecord, pickScalars, projectSender } from "./chats.js";
import {
  TRUNCATION_MARKER,
  capPlainText,
  htmlToText,
  type HtmlTextContext,
  type HtmlTextResult,
} from "./html-text.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";
import { instantTicks } from "./timestamps.js";

/**
 * Teams chat message tools (SM-11 parts B and C): `list-chat-messages`, `get-chat-message`,
 * `list-pinned-chat-messages` and `list-chat-message-hosted-contents` (all v1.0 GET). Chat-message replies are not a capability: Microsoft
 * documents replies only for CHANNEL messages, so the catalog entry was removed (D21).
 *
 * Delegated permissions (current Microsoft docs, October 2026): list messages, get message and
 * list pinned messages all document `Chat.Read` as least privileged, which the catalog uses for
 * all three (D20). None supports
 * personal accounts. The server never pre-checks scopes; Graph decides.
 *
 * Message content is UNTRUSTED data. Every message is projected to an allowlist:
 * - scalars: id, replyToId, messageType, createdDateTime, lastModifiedDateTime,
 *   lastEditedDateTime, deletedDateTime, subject, importance, chatId, webUrl;
 * - `from` (user/application id and display name), `mentions` (id, text and the mentioned
 *   identity's id and display name), `attachments` (id, contentType, name only: no `content`,
 *   `contentUrl`, `thumbnailUrl` or app id), `reactions` (aggregated to type, name and count: no
 *   reacting users, no `reactionContentUrl`), `eventDetail` (its `@odata.type` only);
 * - `body`: ALWAYS plain text (`contentType: "text"`), with `sourceContentType` and, when a
 *   length cap cut it, `truncated: true` plus a visible marker. HTML bodies go through the
 *   dependency-free `htmlToText`; text bodies are only control-character scrubbed and capped.
 *   No raw HTML is ever returned and no `Prefer: outlook.body-content-type` header is sent (that
 *   is an Outlook header, undocumented for Teams).
 * Dropped: etag, summary (channel-only), locale, onBehalfOf, policyViolation, messageHistory,
 * channelIdentity and every unknown property or annotation.
 *
 * Body caps: 2000 characters per message in lists and pinned messages, 20000 in
 * `get-chat-message`. The same caps apply to graph-batch sub-responses.
 */

const MESSAGE_SCALARS = [
  "id",
  "replyToId",
  "messageType",
  "createdDateTime",
  "lastModifiedDateTime",
  "lastEditedDateTime",
  "deletedDateTime",
  "subject",
  "importance",
  "chatId",
  "webUrl",
] as const;
const MESSAGE_KEYS = [
  ...MESSAGE_SCALARS,
  "from",
  "body",
  "mentions",
  "attachments",
  "reactions",
  "eventDetail",
] as const;

export const LIST_BODY_MAX_CHARS = 2000;
export const GET_BODY_MAX_CHARS = 20_000;
/** Longest nested array (mentions, attachments, reaction types) returned per message. */
const MAX_NESTED = 50;

type Rec = Record<string, unknown>;

const asArray = (v: unknown): unknown[] | undefined => (Array.isArray(v) ? v : undefined);

function projectBody(body: unknown, ctx: HtmlTextContext, maxChars: number): unknown {
  if (body === null) return null;
  if (!isRecord(body)) return undefined;
  const content = typeof body["content"] === "string" ? body["content"] : "";
  // Fail closed: only an exact `text` body skips the HTML converter. A missing, re-cased or
  // future contentType is treated as HTML, so markup can never pass through verbatim.
  const isText = body["contentType"] === "text";
  const result: HtmlTextResult = isText
    ? capPlainText(content, maxChars)
    : htmlToText(content, maxChars, ctx);
  const out: Rec = {
    contentType: "text",
    // What Graph declared, as a fixed value (never the raw, untrusted string).
    sourceContentType: isText ? "text" : body["contentType"] === "html" ? "html" : "unknown",
    content: result.truncated ? `${result.text}${TRUNCATION_MARKER}` : result.text,
  };
  if (result.truncated) out["truncated"] = true;
  return out;
}

function projectMentioned(mentioned: unknown): unknown {
  if (mentioned === null) return null;
  if (!isRecord(mentioned)) return undefined;
  const out: Rec = {};
  for (const k of ["user", "application", "conversation", "tag"]) {
    const v = mentioned[k];
    if (v === null) out[k] = null;
    else if (isRecord(v)) {
      out[k] = pickScalars(
        v,
        k === "conversation"
          ? ["id", "displayName", "conversationIdentityType"]
          : ["id", "displayName"],
      );
    }
  }
  return out;
}

function projectMentions(mentions: unknown): Rec[] | undefined {
  const list = asArray(mentions);
  if (list === undefined) return undefined;
  const out: Rec[] = [];
  for (const m of list) {
    if (out.length >= MAX_NESTED) break;
    if (!isRecord(m)) continue;
    const item = pickScalars(m, ["id", "mentionText"]);
    const mentioned = projectMentioned(m["mentioned"]);
    if (mentioned !== undefined) item["mentioned"] = mentioned;
    out.push(item);
  }
  return out;
}

function projectAttachments(attachments: unknown): Rec[] | undefined {
  const list = asArray(attachments);
  if (list === undefined) return undefined;
  // Bounded work: stop scanning once MAX_NESTED attachments are kept.
  const out: Rec[] = [];
  for (const a of list) {
    if (out.length >= MAX_NESTED) break;
    if (isRecord(a)) out.push(pickScalars(a, ["id", "contentType", "name"]));
  }
  return out;
}

/** Reactions collapse to `{ reactionType, displayName?, count }`: who reacted is not returned. */
function projectReactions(reactions: unknown): Rec[] | undefined {
  const list = asArray(reactions);
  if (list === undefined) return undefined;
  const groups = new Map<string, Rec>();
  for (const r of list) {
    if (!isRecord(r) || typeof r["reactionType"] !== "string") continue;
    const name = typeof r["displayName"] === "string" ? r["displayName"] : undefined;
    // Unambiguous key: no separator a crafted reactionType could contain; "" and absent differ.
    const key = JSON.stringify([r["reactionType"], name ?? null]);
    const group = groups.get(key);
    if (group !== undefined) {
      group["count"] = (group["count"] as number) + 1;
    } else if (groups.size < MAX_NESTED) {
      const g: Rec = { reactionType: r["reactionType"] };
      if (name !== undefined) g["displayName"] = name;
      g["count"] = 1;
      groups.set(key, g);
    }
  }
  return [...groups.values()];
}

/**
 * Projects one `chatMessage` to the allowlist above, normalizing its body to text of at most
 * `bodyMax` characters. Non-objects yield undefined.
 */
export function projectChatMessage(item: unknown, bodyMax: number): unknown {
  return projectMessage(item, bodyMax, "chat");
}

/**
 * Channel counterpart (SM-12 part B): the same allowlist and body normalization, except that the
 * always-null `chatId` is replaced by `channelIdentity` (`teamId`, `channelId`), which also scopes
 * the hosted-content markers in the body to this message's own team, channel and message.
 */
export function projectChannelMessage(item: unknown, bodyMax: number): unknown {
  return projectMessage(item, bodyMax, "channel");
}

const CHANNEL_SCALARS = MESSAGE_SCALARS.filter((k) => k !== "chatId");
export const CHANNEL_MESSAGE_KEYS = [
  ...MESSAGE_KEYS.filter((k) => k !== "chatId"),
  "channelIdentity",
] as const;

function projectMessage(item: unknown, bodyMax: number, kind: "chat" | "channel"): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, kind === "chat" ? MESSAGE_SCALARS : CHANNEL_SCALARS);
  const from = projectSender(item["from"]);
  if (from !== undefined) out["from"] = from;
  let identity: Rec | undefined;
  if (kind === "channel" && isRecord(item["channelIdentity"])) {
    identity = pickScalars(item["channelIdentity"], ["teamId", "channelId"]);
    out["channelIdentity"] = identity;
  }
  if (Object.hasOwn(item, "body")) {
    const ctx: HtmlTextContext = {};
    if (kind === "chat") {
      if (typeof item["chatId"] === "string") ctx.chatId = item["chatId"];
    } else if (
      typeof identity?.["teamId"] === "string" &&
      typeof identity["channelId"] === "string"
    ) {
      ctx.teamId = identity["teamId"];
      ctx.channelId = identity["channelId"];
      if (typeof item["replyToId"] === "string") ctx.replyToId = item["replyToId"];
    }
    if (typeof item["id"] === "string") ctx.messageId = item["id"];
    const body = projectBody(item["body"], ctx, bodyMax);
    if (body !== undefined) out["body"] = body;
  }
  const mentions = projectMentions(item["mentions"]);
  if (mentions !== undefined) out["mentions"] = mentions;
  const attachments = projectAttachments(item["attachments"]);
  if (attachments !== undefined) out["attachments"] = attachments;
  const reactions = projectReactions(item["reactions"]);
  if (reactions !== undefined) out["reactions"] = reactions;
  const detail = item["eventDetail"];
  if (isRecord(detail)) out["eventDetail"] = pickScalars(detail, ["@odata.type"]);
  return out;
}

/** Pinned message info: `id`, `@odata.type`, and the message only when `$expand=message` was asked. */
export function projectPinnedMessage(item: unknown, withMessage: boolean): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, ["id", "@odata.type"]);
  if (withMessage && Object.hasOwn(item, "message")) {
    const message = item["message"];
    if (message === null) out["message"] = null;
    else {
      const projected = projectChatMessage(message, LIST_BODY_MAX_CHARS);
      if (projected !== undefined) out["message"] = projected;
    }
  }
  return out;
}

const wantsMessage = (query: Readonly<Record<string, unknown>>): boolean => {
  const expand = query["$expand"];
  return typeof expand === "string" && expand.split(",").includes("message");
};

/**
 * Routes every projection path of `base` (single entity, collection item, batch body) through
 * `project`, given THAT request's wire query. `base`'s own projection still runs first in a batch
 * (it clips to the subrequest's `$top`).
 */
export function withProjection(
  base: QueryContract,
  collection: boolean,
  project: (item: unknown, query: Readonly<Record<string, unknown>>) => unknown,
): QueryContract {
  const inner = base.batchPolicy.project;
  return {
    ...base,
    projectItem: (item) => project(item, {}),
    projectResponse: (item, query) => project(item, query),
    itemProjector: (query) => (item) => project(item, query),
    batchPolicy: {
      check: base.batchPolicy.check,
      project: (body, pairs) => {
        const query = Object.fromEntries(pairs);
        const out = inner === undefined ? body : inner(body, pairs);
        if (!collection) return project(out, query) ?? {};
        if (!isRecord(out) || !Array.isArray(out["value"])) return {};
        const value = out["value"].map((i) => project(i, query)).filter((i) => i !== undefined);
        return { ...out, value };
      },
    },
  };
}

// --- Typed time range for list-chat-messages ---------------------------------------------------
//
// Graph's `$filter` on chat messages is only honoured with a matching `$orderby` on the same
// property (otherwise it is silently ignored), `lastModifiedDateTime` supports `gt`/`lt` and
// `createdDateTime` only `lt`. Callers therefore never send raw OData: they pass typed instants and
// the server sets the coupled `$orderby <property> desc` itself.

type RangeProperty = "lastModifiedDateTime" | "createdDateTime";
interface Range {
  property: RangeProperty;
  after?: string;
  before?: string;
}

const RANGE_KEYS = ["modifiedAfter", "modifiedBefore", "createdBefore"] as const;
const WIRE_STAMP_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/;

function invalid(message: string): GraphClientError {
  return new GraphClientError("invalid_input", message);
}

function stampInput(what: string) {
  return (
    z
      .string()
      .refine((v) => instantTicks(v) !== undefined)
      // Graph filters at millisecond precision; a sub-millisecond cutoff would have to be rounded,
      // silently moving the boundary, so it is rejected instead.
      .refine((v) => {
        const ticks = instantTicks(v);
        return ticks === undefined || ticks % 10_000n === 0n;
      })
      .optional()
      .describe(
        `${what} as an ISO-8601 timestamp with a Z or an explicit UTC offset (e.g. 2024-01-31T00:00:00Z), at most millisecond precision (a non-zero sub-millisecond fraction is rejected). Sent to Graph as UTC.`,
      )
  );
}

const RANGE_SHAPE = {
  modifiedAfter: stampInput(
    "Only messages whose lastModifiedDateTime is after this instant. Results are then ordered by lastModifiedDateTime descending (set by the server); cannot be combined with createdBefore or an orderby on createdDateTime.",
  ),
  modifiedBefore: stampInput(
    "Only messages whose lastModifiedDateTime is before this instant (use the oldest lastModifiedDateTime of the previous result to read older messages). Ordered by lastModifiedDateTime descending; cannot be combined with createdBefore or an orderby on createdDateTime.",
  ),
  createdBefore: stampInput(
    "Only messages whose createdDateTime is before this instant. Ordered by createdDateTime descending (set by the server); cannot be combined with modifiedAfter/modifiedBefore or an orderby on lastModifiedDateTime.",
  ),
} satisfies z.ZodRawShape;
const RANGE_SCHEMA = z.strictObject(RANGE_SHAPE);

/** Canonical wire instant: UTC, exactly three fraction digits (the form Graph's examples use). */
function wireStamp(value: string): string {
  const ticks = instantTicks(value) as bigint;
  let ms = ticks / 10_000n;
  if (ticks % 10_000n < 0n) ms -= 1n;
  return new Date(Number(ms)).toISOString();
}

function filterWire(range: Range): string {
  const p = range.property;
  const clauses: string[] = [];
  if (range.after !== undefined) clauses.push(`${p} gt ${range.after}`);
  if (range.before !== undefined) clauses.push(`${p} lt ${range.before}`);
  return clauses.join(" and ");
}

const TS = "(\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}\\.\\d{3}Z)";
const FILTER_FORMS: ReadonlyArray<{ re: RegExp; property: RangeProperty }> = [
  {
    re: new RegExp(`^lastModifiedDateTime gt ${TS}(?: and lastModifiedDateTime lt ${TS})?$`),
    property: "lastModifiedDateTime",
  },
  { re: new RegExp(`^lastModifiedDateTime lt ${TS}$`), property: "lastModifiedDateTime" },
  { re: new RegExp(`^createdDateTime lt ${TS}$`), property: "createdDateTime" },
];

const rangeProblem = (range: Range): string | undefined => {
  for (const v of [range.after, range.before]) {
    if (v !== undefined && (!WIRE_STAMP_RE.test(v) || instantTicks(v) === undefined)) {
      return "Invalid timestamp.";
    }
    if (v !== undefined && wireStamp(v) !== v) return "Invalid timestamp.";
  }
  if (
    range.after !== undefined &&
    range.before !== undefined &&
    (instantTicks(range.after) as bigint) >= (instantTicks(range.before) as bigint)
  ) {
    return "modifiedAfter must be before modifiedBefore.";
  }
  return undefined;
};

/** The typed range of tool args (undefined when none was given); throws `invalid_input`. */
function rangeOfArgs(args: Readonly<Record<string, unknown>>): Range | undefined {
  const raw: Record<string, unknown> = {};
  for (const k of RANGE_KEYS) if (args[k] !== undefined) raw[k] = args[k];
  if (Object.keys(raw).length === 0) return undefined;
  const parsed = RANGE_SCHEMA.safeParse(raw);
  if (!parsed.success) {
    const bad = new Set(parsed.error.issues.map((i) => String(i.path[0])));
    throw invalid(`Invalid query options: ${[...bad].join(", ")}.`);
  }
  const { modifiedAfter, modifiedBefore, createdBefore } = parsed.data;
  if (
    createdBefore !== undefined &&
    (modifiedAfter !== undefined || modifiedBefore !== undefined)
  ) {
    throw invalid("createdBefore cannot be combined with modifiedAfter or modifiedBefore.");
  }
  const range: Range =
    createdBefore !== undefined
      ? { property: "createdDateTime", before: wireStamp(createdBefore) }
      : { property: "lastModifiedDateTime" };
  if (modifiedAfter !== undefined) range.after = wireStamp(modifiedAfter);
  if (modifiedBefore !== undefined) range.before = wireStamp(modifiedBefore);
  const problem = rangeProblem(range);
  if (problem !== undefined) throw invalid(problem);
  return range;
}

function parseFilterWire(value: string): Range {
  for (const form of FILTER_FORMS) {
    const m = form.re.exec(value);
    if (m === null) continue;
    const range: Range = { property: form.property };
    if (m[2] !== undefined) {
      range.after = m[1] as string;
      range.before = m[2];
    } else if (value.includes(" gt ")) range.after = m[1] as string;
    else range.before = m[1] as string;
    const problem = rangeProblem(range);
    if (problem !== undefined) throw invalid(problem);
    return range;
  }
  throw invalid("$filter must be one of the typed message time ranges.");
}

/**
 * Adds the typed range to a list contract. `$filter` is not part of the base contract, so the base
 * batch policy would reject it; this wrapper takes it out of the batch pairs, parses it back into
 * exactly the forms the single tool emits and requires the coupled `$orderby` (without which Graph
 * would silently ignore the filter).
 */
function withRange(base: QueryContract): QueryContract {
  const coupled = (property: RangeProperty): string => `${property} desc`;
  return {
    ...base,
    shape: { ...base.shape, ...RANGE_SHAPE },
    toRequest: (args) => {
      const range = rangeOfArgs(args);
      const req = base.toRequest(args);
      if (range === undefined) return req;
      const given = req.query["$orderby"];
      if (given !== undefined && given !== coupled(range.property)) {
        throw invalid(
          `With this time range, orderby must be ${coupled(range.property)} (or omitted).`,
        );
      }
      return {
        ...req,
        query: {
          ...req.query,
          $orderby: coupled(range.property),
          $filter: filterWire(range),
        },
      };
    },
    batchPolicy: {
      ...base.batchPolicy,
      check: (pairs, headers) => {
        const filters = pairs.filter(([name]) => name === "$filter");
        if (filters.length > 1) throw invalid("Duplicate query parameter.");
        const rest = pairs.filter(([name]) => name !== "$filter");
        base.batchPolicy.check(rest, headers);
        const filter = filters[0];
        if (filter === undefined) return;
        const range = parseFilterWire(filter[1]);
        const orderby = rest.find(([name]) => name === "$orderby")?.[1];
        if (orderby !== coupled(range.property)) {
          throw invalid(`$filter requires $orderby=${coupled(range.property)}.`);
        }
      },
    },
  };
}

const listContract = (): QueryContract =>
  withRange(
    withProjection(
      defineQueryContract({
        // Graph: "Maximum allowed $top value is 50" per request; more is reached by paging.
        top: { default: 20, max: 100, pageMax: 50 },
        orderby: {
          allowed: ["createdDateTime", "lastModifiedDateTime"],
          maxClauses: 1,
          requireInFilter: false,
          descOnly: true,
        },
        projectEntity: { keep: MESSAGE_KEYS },
      }),
      true,
      (item) => projectChatMessage(item, LIST_BODY_MAX_CHARS),
    ),
  );

const getContract = (): QueryContract =>
  withProjection(defineQueryContract({ projectEntity: { keep: MESSAGE_KEYS } }), false, (item) =>
    projectChatMessage(item, GET_BODY_MAX_CHARS),
  );

const pinnedContract = (): QueryContract =>
  withProjection(
    defineQueryContract({
      // The page documents only `$expand`: `top` is a local limit and no `$top` is sent.
      top: { default: 25, max: 100, wire: false },
      expand: { allowed: ["message"] },
      projectEntity: { keep: ["message"] },
    }),
    true,
    (item, query) => projectPinnedMessage(item, wantsMessage(query)),
  );

// --- Hosted contents (SM-11 part C) ---------------------------------------------------------------
//
// `list-chat-message-hosted-contents` lists REFERENCES only. Microsoft documents `Chat.Read` as the
// least privileged delegated permission (the catalog uses it since D23), no OData
// support, and that `contentBytes`/`contentType` are always null in the list. There is deliberately
// no byte download (no `/$value` path is in the catalog); each item is projected to `id` and, if a
// string, `contentType`. `contentBytes`, `@odata.mediaContentType`/`@odata.mediaReadLink` and any
// unknown property are dropped, so even a future Graph that fills `contentBytes` never leaks it.
// The `id` is the same value `htmlToText` surfaces as `[image hostedContentId=<id>]`.

/**
 * Projects one `chatMessageHostedContent` to `{ id, contentType? }`. Non-objects and items without a
 * string `id` yield undefined (dropped): a reference without an id is useless.
 */
export function projectHostedContent(item: unknown): unknown {
  if (!isRecord(item) || typeof item["id"] !== "string") return undefined;
  const out: Rec = { id: item["id"] };
  if (typeof item["contentType"] === "string") out["contentType"] = item["contentType"];
  return out;
}

export const hostedContentsContract = (): QueryContract =>
  withProjection(
    defineQueryContract({
      // The page documents no OData: `top` is a local limit and no `$top` is sent.
      top: { default: 25, max: 100, wire: false },
      // Load-bearing on a collection contract: it adds `contentType` to the inner batch keep-set
      // (`projectCollection`), so the outer `projectHostedContent` sees it in batch as on the
      // single path. The batch-parity test pins this.
      projectEntity: { keep: ["contentType"] },
    }),
    true,
    (item) => projectHostedContent(item),
  );

export function registerChatMessageTools(
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
    "list-chat-messages",
    listTool(client, "list-chat-messages", patternOf("list-chat-messages"), listContract()),
  );
  handlers.set(
    "get-chat-message",
    getTool(client, "get-chat-message", patternOf("get-chat-message"), getContract()),
  );
  handlers.set(
    "list-pinned-chat-messages",
    listTool(
      client,
      "list-pinned-chat-messages",
      patternOf("list-pinned-chat-messages"),
      pinnedContract(),
    ),
  );
  handlers.set(
    "list-chat-message-hosted-contents",
    listTool(
      client,
      "list-chat-message-hosted-contents",
      patternOf("list-chat-message-hosted-contents"),
      hostedContentsContract(),
    ),
  );
}
