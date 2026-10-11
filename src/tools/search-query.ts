import { z } from "zod";
import type { EndpointCatalog } from "../catalog.js";
import { errorResult, successResult } from "../envelope.js";
import { GraphClientError } from "../graph/errors.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { postPath } from "./calendar-availability.js";
import { isRecord, pickScalars } from "./chats.js";

/**
 * SM-14 part B: `search-query` (`POST /search/query`, v1.0), decision D38.
 *
 * Retrieval-only POST. The Graph body is CONSTRUCTED here from validated typed inputs and is always
 * exactly `{ requests: [ { entityTypes, query: { queryString }, from, size } ] }`: one logical
 * search request, no other request property is ever sent (no fields, aggregations,
 * aggregationFilters, collapseProperties, sortProperties, queryTemplate, contentSources, region,
 * enableTopResults, queryAlterationOptions, resultTemplateOptions, sharePointOneDriveOptions) and
 * no caller headers. The registry validates the input with a strict object, so every other top
 * level property is rejected before the handler runs. There is no `batchPolicy`, so `graph-batch`
 * (GET-only) can never reach this tool.
 *
 * Microsoft documents (Learn, October 2026): the service supports a single searchRequest at a time;
 * only specific entity type combinations; `size` max 25 for message and event (the overview also
 * says 1000 and searchRequest says 500 generally); `from` must start at 0 on the first page for
 * message/event (we always send an explicit `from`, default 0). The `queryString` length limit,
 * the `from` maximum, the 50 cap for non-mail/event types, and every response bound below are
 * PROJECT limits, not Microsoft ones.
 *
 * Responses are untrusted. Hits are projected: `hitId`, `rank`, a bounded `summary` (Search
 * highlight markup stripped) and a per-entity-type allowlisted `resource`. A resource whose
 * `@odata.type` is not a supported type of the requested family is dropped (envelope kept). A
 * malformed body fails closed.
 */

export const SEARCH_ENTITY_TYPES = [
  "message",
  "event",
  "chatMessage",
  "drive",
  "driveItem",
  "list",
  "listItem",
  "site",
] as const;
export type SearchEntityType = (typeof SEARCH_ENTITY_TYPES)[number];

const SHAREPOINT_TYPES: readonly SearchEntityType[] = [
  "drive",
  "driveItem",
  "list",
  "listItem",
  "site",
];
/** message, event and chatMessage each must be searched alone. */
const SOLO_TYPES: readonly SearchEntityType[] = ["message", "event", "chatMessage"];

/** Project limits (not Microsoft limits). */
export const MAX_QUERY_STRING = 1000;
export const MAX_FROM = 1000;
export const DEFAULT_SIZE = 25;
/** Microsoft-documented cap for message and event. */
export const MAX_SIZE_MAIL_EVENT = 25;
/** Project cap for every other supported type. */
export const MAX_SIZE_OTHER = 50;
/** Response bounds (project): exceeding one means Graph broke its contract, so fail closed. */
export const MAX_SUMMARY = 500;
const MAX_SEARCH_TERMS = 50;
const MAX_TERM = 500;
const MAX_CONTAINERS = 5;
const MAX_HIT_ID = 1024;
const MAX_RESOURCE_TEXT = 4096;

function malformed(what: string): never {
  throw new GraphClientError("graph_error", `Graph returned a malformed ${what}.`);
}

function failure(err: unknown): ReturnType<typeof errorResult> {
  if (err instanceof GraphClientError) return err.toResult();
  return errorResult("graph_error", "Tool search-query failed unexpectedly.");
}

/** Local combination rule (D38). Returns an error message, or undefined when allowed. */
export function entityTypeCombinationIssue(types: readonly SearchEntityType[]): string | undefined {
  if (types.length === 1) return undefined;
  for (const solo of SOLO_TYPES) {
    if (types.includes(solo)) {
      return `${solo} must be searched alone; it cannot be combined with other entity types.`;
    }
  }
  return undefined; // only SharePoint/OneDrive types remain: any combination is allowed
}

/** Maximum page size for the selected entity types. */
export function maxSizeFor(types: readonly SearchEntityType[]): number {
  return types.includes("message") || types.includes("event")
    ? MAX_SIZE_MAIL_EVENT
    : MAX_SIZE_OTHER;
}

const SEARCH_QUERY_SHAPE = {
  entityTypes: z
    .array(z.enum(SEARCH_ENTITY_TYPES))
    .min(1)
    .max(SEARCH_ENTITY_TYPES.length)
    .refine((list) => new Set(list).size === list.length, { message: "entityTypes must be unique" })
    .describe(
      "Entity types to search. message, event or chatMessage must be searched alone; drive, driveItem, list, listItem and site may be combined with each other only.",
    ),
  queryString: z
    .string()
    .max(MAX_QUERY_STRING)
    .refine((s) => s.trim().length > 0, { message: "queryString must not be blank" })
    .describe(
      `Search text, optionally KQL (1-${MAX_QUERY_STRING} characters, not blank). Lexical Microsoft Search, not semantic retrieval.`,
    ),
  from: z
    .number()
    .int()
    .min(0)
    .max(MAX_FROM)
    .optional()
    .describe(
      `0-based result offset (0-${MAX_FROM}, default 0). For message and event the first page must start at 0.`,
    ),
  size: z
    .number()
    .int()
    .min(1)
    .max(MAX_SIZE_OTHER)
    .optional()
    .describe(
      `Page size (default ${DEFAULT_SIZE}). At most ${MAX_SIZE_MAIL_EVENT} for message and event, at most ${MAX_SIZE_OTHER} for the other types.`,
    ),
} satisfies z.ZodRawShape;

// ------------------------------------------------------------------ response projection

const absent = (v: unknown): v is undefined | null => v === undefined || v === null;

/** `<c0>`, `</c0>` and `<ddd/>` are the Search highlight/ellipsis markers. */
const HIGHLIGHT_MARKUP = /<\/?c\d+>|<ddd\/>/g;

/** Strips Search highlight markup and bounds the length (truncation is marked with an ellipsis). */
export function projectSummary(raw: unknown): string | undefined {
  if (absent(raw)) return undefined;
  if (typeof raw !== "string") malformed("search summary");
  const clean = raw.replace(HIGHLIGHT_MARKUP, "");
  // A cut summary, ellipsis included, stays within MAX_SUMMARY characters.
  return clean.length > MAX_SUMMARY ? `${clean.slice(0, MAX_SUMMARY - 1)}…` : clean;
}

/** Keeps only scalar `keys` of the nested record at `obj[key]`; empty or absent yields undefined. */
function sub(obj: Record<string, unknown>, key: string, keys: readonly string[]) {
  const v = obj[key];
  if (!isRecord(v)) return undefined;
  const picked = pickScalars(v, keys);
  return Object.keys(picked).length > 0 ? picked : undefined;
}

function setIf(out: Record<string, unknown>, key: string, value: unknown): void {
  if (value !== undefined) out[key] = value;
}

/** from.emailAddress.{name,address} (mail, chat search shape) and from.user.displayName. */
function projectFrom(
  obj: Record<string, unknown>,
  key: string,
): Record<string, unknown> | undefined {
  const v = obj[key];
  if (!isRecord(v)) return undefined;
  const out: Record<string, unknown> = {};
  const email = sub(v, "emailAddress", ["name", "address"]);
  if (email !== undefined) out["emailAddress"] = email;
  const user = sub(v, "user", ["displayName"]);
  if (user !== undefined) out["user"] = user;
  return Object.keys(out).length > 0 ? out : undefined;
}

const PROJECTORS: Record<
  SearchEntityType,
  (r: Record<string, unknown>) => Record<string, unknown>
> = {
  message: (r) => {
    const out = pickScalars(r, [
      "id",
      "subject",
      "receivedDateTime",
      "sentDateTime",
      "hasAttachments",
      "webLink",
    ]);
    setIf(out, "from", projectFrom(r, "from"));
    return out;
  },
  event: (r) => {
    const out = pickScalars(r, ["id", "subject", "isAllDay", "webLink"]);
    setIf(out, "start", sub(r, "start", ["dateTime", "timeZone"]));
    setIf(out, "end", sub(r, "end", ["dateTime", "timeZone"]));
    const organizer = isRecord(r["organizer"])
      ? sub(r["organizer"], "emailAddress", ["name"])
      : undefined;
    if (organizer !== undefined) out["organizer"] = { emailAddress: organizer };
    return out;
  },
  chatMessage: (r) => {
    const out = pickScalars(r, [
      "id",
      "createdDateTime",
      "lastModifiedDateTime",
      "chatId",
      "webLink",
      "webUrl",
    ]);
    setIf(out, "from", projectFrom(r, "from"));
    setIf(out, "channelIdentity", sub(r, "channelIdentity", ["teamId", "channelId"]));
    return out;
  },
  drive: (r) => pickScalars(r, ["id", "name", "webUrl", "driveType"]),
  driveItem: (r) => {
    const out = pickScalars(r, [
      "id",
      "name",
      "webUrl",
      "size",
      "createdDateTime",
      "lastModifiedDateTime",
    ]);
    setIf(out, "parentReference", sub(r, "parentReference", ["driveId", "siteId"]));
    return out;
  },
  list: (r) => pickScalars(r, ["id", "displayName", "name", "webUrl"]),
  listItem: (r) => {
    const out = pickScalars(r, ["id", "webUrl", "createdDateTime", "lastModifiedDateTime"]);
    setIf(out, "parentReference", sub(r, "parentReference", ["siteId"]));
    return out;
  },
  site: (r) => pickScalars(r, ["id", "displayName", "name", "webUrl"]),
};

/** `#microsoft.graph.driveItem` or `microsoft.graph.chatMessage` (Learn shows both) -> entity type. */
function resourceType(resource: Record<string, unknown>): SearchEntityType | undefined {
  const t = resource["@odata.type"];
  if (typeof t !== "string") return undefined;
  const name = t.replace(/^#/, "");
  if (!name.startsWith("microsoft.graph.")) return undefined;
  const short = name.slice("microsoft.graph.".length);
  return (SEARCH_ENTITY_TYPES as readonly string[]).includes(short)
    ? (short as SearchEntityType)
    : undefined;
}

function boundStrings(value: unknown): void {
  if (typeof value === "string") {
    if (value.length > MAX_RESOURCE_TEXT) malformed("search resource field");
  } else if (isRecord(value)) {
    Object.values(value).forEach(boundStrings);
  }
}

/**
 * Projects one hit's resource. Returns undefined when the resource is absent, not an object, of an
 * unknown type, or of a type outside the requested family (the envelope is still returned).
 */
export function projectResource(
  resource: unknown,
  requested: readonly SearchEntityType[],
): { type: SearchEntityType; resource: Record<string, unknown> } | undefined {
  if (!isRecord(resource)) return undefined;
  const type = resourceType(resource);
  if (type === undefined) return undefined;
  const allowed = requested.some((t) => SHAREPOINT_TYPES.includes(t))
    ? SHAREPOINT_TYPES
    : requested;
  if (!allowed.includes(type)) return undefined;
  const projected = PROJECTORS[type](resource);
  boundStrings(projected);
  return { type, resource: projected };
}

function projectHit(hit: unknown, requested: readonly SearchEntityType[]): Record<string, unknown> {
  if (!isRecord(hit)) malformed("search hit");
  const hitId = hit["hitId"];
  if (typeof hitId !== "string" || hitId.length === 0 || hitId.length > MAX_HIT_ID) {
    malformed("search hit");
  }
  const out: Record<string, unknown> = { hitId };
  const rank = hit["rank"];
  if (!absent(rank)) {
    if (typeof rank !== "number" || !Number.isFinite(rank)) malformed("search hit rank");
    out["rank"] = rank;
  }
  setIf(out, "summary", projectSummary(hit["summary"]));
  const res = projectResource(hit["resource"], requested);
  if (res !== undefined) {
    out["resourceType"] = res.type;
    out["resource"] = res.resource;
  }
  return out;
}

function projectContainer(
  container: unknown,
  requested: readonly SearchEntityType[],
  size: number,
): Record<string, unknown> {
  if (!isRecord(container)) malformed("search hits container");
  const total = container["total"];
  // An absent `hits` must not read as a zero-result success. It is accepted only for an explicit
  // empty result (`total: 0` and no more results), a shape Graph may use for zero hits; every
  // other absent `hits` fails closed.
  const explicitEmpty = total === 0 && container["moreResultsAvailable"] !== true;
  const hits = container["hits"] === undefined && explicitEmpty ? [] : container["hits"];
  if (!Array.isArray(hits) || hits.length > size) malformed("search hits list");
  const out: Record<string, unknown> = {};
  if (!absent(total)) {
    if (typeof total !== "number" || !Number.isFinite(total) || total < 0)
      malformed("search total");
    out["total"] = total;
  }
  const more = container["moreResultsAvailable"];
  if (!absent(more)) {
    if (typeof more !== "boolean") malformed("search moreResultsAvailable flag");
    out["moreResultsAvailable"] = more;
  }
  out["hits"] = hits.map((h) => projectHit(h, requested));
  return out;
}

/**
 * Projects the Graph response `{ value: [ searchResponse ] }`. Exactly one searchResponse is
 * expected (one request was sent); anything else fails closed. `queryAlterationResponse`,
 * `resultTemplates`, aggregations and every unknown property are dropped.
 */
export function projectSearchResponse(
  data: unknown,
  requested: readonly SearchEntityType[],
  size: number,
): Record<string, unknown> {
  if (!isRecord(data)) malformed("search response");
  const value = data["value"];
  if (!Array.isArray(value) || value.length !== 1) malformed("search response");
  const response: unknown = value[0];
  if (!isRecord(response)) malformed("search response");

  // searchTerms only echoes the query; its absence cannot misstate results, so [] is safe here.
  const termsRaw = response["searchTerms"] ?? [];
  if (
    !Array.isArray(termsRaw) ||
    termsRaw.length > MAX_SEARCH_TERMS ||
    !termsRaw.every((t) => typeof t === "string" && t.length <= MAX_TERM)
  ) {
    malformed("search terms list");
  }
  // Absent `hitsContainers` must not read as a zero-result success: fail closed.
  const containers = response["hitsContainers"];
  if (!Array.isArray(containers) || containers.length > MAX_CONTAINERS) {
    malformed("search hits container list");
  }
  return {
    searchTerms: termsRaw,
    hitsContainers: containers.map((c) => projectContainer(c, requested, size)),
  };
}

// ------------------------------------------------------------------ tool

/** `search-query`. Not batch-reachable: no `batchPolicy`. */
export function registerSearchQueryTool(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const pathPattern = postPath(catalog, "search-query");
  const entry = catalog.find((e) => e.toolName === "search-query");
  if (entry === undefined || entry["apiVersion"] === "beta") {
    throw new Error("Catalog has no v1.0 POST entry for search-query.");
  }

  handlers.set("search-query", {
    inputSchema: SEARCH_QUERY_SHAPE,
    handler: async (args) => {
      try {
        const a = args as {
          entityTypes: SearchEntityType[];
          queryString: string;
          from?: number;
          size?: number;
        };
        const combination = entityTypeCombinationIssue(a.entityTypes);
        if (combination !== undefined) return errorResult("invalid_input", combination);
        const size = a.size ?? DEFAULT_SIZE;
        const maxSize = maxSizeFor(a.entityTypes);
        if (size > maxSize) {
          return errorResult(
            "invalid_input",
            `size must be at most ${maxSize} for the selected entityTypes.`,
          );
        }
        const body = {
          requests: [
            {
              entityTypes: [...a.entityTypes],
              query: { queryString: a.queryString },
              from: a.from ?? 0,
              size,
            },
          ],
        };
        const res = await client.request({ pathPattern, method: "post", body });
        const data = projectSearchResponse(res.data, a.entityTypes, size);
        const meta: Record<string, unknown> = { clientRequestId: res.clientRequestId };
        if (res.requestId !== undefined) meta["requestId"] = res.requestId;
        return successResult(data, meta);
      } catch (err) {
        return failure(err);
      }
    },
  });
}
