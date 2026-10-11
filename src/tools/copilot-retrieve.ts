import { z } from "zod";
import type { EndpointCatalog } from "../catalog.js";
import { errorResult, successResult } from "../envelope.js";
import { GraphClientError } from "../graph/errors.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { postPath } from "./calendar-availability.js";
import { isRecord } from "./chats.js";

/**
 * SM-14 part C: `copilot-retrieve` (`POST /copilot/retrieval`, v1.0), decision D41.
 *
 * Retrieval-only POST. The Graph body is CONSTRUCTED here from validated typed top-level inputs and
 * contains only the fields the caller provided: `queryString`, `dataSource`, and optionally
 * `filterExpression`, `resourceMetadata`, `maximumNumberOfResults`. Never sent: `dataSourceConfiguration`
 * (connector connection ids), the beta-only `includeThumbnails`, any other body property, or any
 * caller header. The registry validates input with a strict object, so every other top-level
 * property (including a nested `body`) is rejected before the handler runs. There is no
 * `batchPolicy`, so `graph-batch` (GET-only) can never reach this tool.
 *
 * Microsoft documents (Learn, October 2026): `queryString` is required with a limit of 1,500
 * characters; `dataSource` is `sharePoint`, `oneDriveBusiness` or `externalItem`;
 * `maximumNumberOfResults` is 1-25 (default up to 25); invalid KQL in `filterExpression` makes the
 * query run unscoped (so it is not a security boundary). The `filterExpression` and `resourceMetadata`
 * limits and every response bound below are PROJECT limits, not Microsoft ones.
 *
 * Responses are untrusted. They are projected through allowlists: `webUrl`, `resourceType`, bounded
 * `extracts`, requested-only scalar `resourceMetadata` and a minimal `sensitivityLabel`. A body
 * without a `retrievalHits` array fails closed (Learn: an EMPTY array means "no relevant results",
 * an absent one is not a zero-result success).
 */

export const COPILOT_DATA_SOURCES = ["sharePoint", "oneDriveBusiness", "externalItem"] as const;

/** Microsoft-documented limits. */
export const MAX_QUERY_STRING = 1500;
export const MAX_RESULTS = 25;
/** Project limits (not Microsoft limits). */
export const MAX_FILTER_EXPRESSION = 1000;
export const MAX_RESOURCE_METADATA_KEYS = 20;
export const RESOURCE_METADATA_KEY = /^[A-Za-z][A-Za-z0-9_]{0,63}$/;
/** Response bounds (project). Over-long text is truncated with an ellipsis within the limit. */
export const MAX_EXTRACT_TEXT = 4000;
export const MAX_EXTRACTS_PER_HIT = 20;
export const MAX_METADATA_VALUE = 1000;
export const MAX_LABEL_TEXT = 256;
const MAX_WEB_URL = 2048;

const RESOURCE_TYPES: readonly string[] = [
  "site",
  "list",
  "listItem",
  "externalItem",
  "drive",
  "driveItem",
];

function malformed(what: string): never {
  throw new GraphClientError("graph_error", `Graph returned a malformed ${what}.`);
}

function failure(err: unknown): ReturnType<typeof errorResult> {
  if (err instanceof GraphClientError) return err.toResult();
  return errorResult("graph_error", "Tool copilot-retrieve failed unexpectedly.");
}

const COPILOT_RETRIEVE_SHAPE = {
  queryString: z
    .string()
    .max(MAX_QUERY_STRING)
    .refine((s) => s.trim().length > 0, { message: "queryString must not be blank" })
    .describe(
      `Natural-language query, preferably one sentence (1-${MAX_QUERY_STRING} characters, not blank). Avoid spelling errors in context-rich keywords.`,
    ),
  dataSource: z
    .enum(COPILOT_DATA_SOURCES)
    .describe(
      "Exactly one source per call: sharePoint, oneDriveBusiness, or externalItem (Copilot connectors).",
    ),
  filterExpression: z
    .string()
    .max(MAX_FILTER_EXPRESSION)
    .refine((s) => s.trim().length > 0, { message: "filterExpression must not be blank" })
    .optional()
    .describe(
      `Optional KQL scoping expression (1-${MAX_FILTER_EXPRESSION} characters). Invalid KQL may be silently ignored and the query run unscoped, so this is not a security boundary.`,
    ),
  resourceMetadata: z
    .array(z.string().regex(RESOURCE_METADATA_KEY))
    .min(1)
    .max(MAX_RESOURCE_METADATA_KEYS)
    // Case-insensitive, matching the case-insensitive response lookup (no duplicate outputs).
    .refine((list) => new Set(list.map((key) => key.toLowerCase())).size === list.length, {
      message: "resourceMetadata must be unique (case-insensitive)",
    })
    .optional()
    .describe(
      `Optional retrievable metadata field names to return per hit (for example title, author); at most ${MAX_RESOURCE_METADATA_KEYS}, each a plain identifier. Only these keys are returned.`,
    ),
  maximumNumberOfResults: z
    .number()
    .int()
    .min(1)
    .max(MAX_RESULTS)
    .optional()
    .describe(
      `Optional cap on hits (1-${MAX_RESULTS}; the service default is up to ${MAX_RESULTS}).`,
    ),
} satisfies z.ZodRawShape;

// ------------------------------------------------------------------ response projection

const absent = (v: unknown): v is undefined | null => v === undefined || v === null;

/** Bounds untrusted text; a cut string, ellipsis included, stays within `max` characters. */
export function boundText(raw: string, max: number): string {
  return raw.length > max ? `${raw.slice(0, max - 1)}…` : raw;
}

function projectExtract(extract: unknown): Record<string, unknown> {
  if (!isRecord(extract)) malformed("retrieval extract");
  const text = extract["text"];
  if (typeof text !== "string") malformed("retrieval extract");
  const out: Record<string, unknown> = { text: boundText(text, MAX_EXTRACT_TEXT) };
  const score = extract["relevanceScore"];
  if (!absent(score)) {
    // Learn: an extract can be returned without a relevanceScore. Graph commonly serializes an unset
    // nullable property as null, so null counts as absent; any other non-finite value fails closed.
    if (typeof score !== "number" || !Number.isFinite(score)) malformed("retrieval relevanceScore");
    out["relevanceScore"] = score;
  }
  return out;
}

/** Only keys the caller requested (case-insensitive match, returned under the requested spelling). */
function projectMetadata(raw: unknown, requested: readonly string[]): Record<string, unknown> {
  if (absent(raw)) return {};
  if (!isRecord(raw)) malformed("retrieval resourceMetadata");
  const out: Record<string, unknown> = {};
  for (const key of requested) {
    const source = Object.keys(raw).find((k) => k.toLowerCase() === key.toLowerCase());
    if (source === undefined) continue;
    const v = raw[source];
    if (typeof v === "string") out[key] = boundText(v, MAX_METADATA_VALUE);
    else if (typeof v === "boolean" || (typeof v === "number" && Number.isFinite(v))) out[key] = v;
    // Non-scalar and null values are dropped.
  }
  return out;
}

function projectLabel(raw: unknown): Record<string, unknown> | undefined {
  if (absent(raw)) return undefined;
  if (!isRecord(raw)) malformed("retrieval sensitivityLabel");
  const out: Record<string, unknown> = {};
  for (const k of ["sensitivityLabelId", "displayName"]) {
    const v = raw[k];
    if (typeof v === "string") out[k] = boundText(v, MAX_LABEL_TEXT);
  }
  const priority = raw["priority"];
  if (typeof priority === "number" && Number.isFinite(priority)) out["priority"] = priority;
  return Object.keys(out).length > 0 ? out : undefined;
}

function projectHit(hit: unknown, requested: readonly string[]): Record<string, unknown> {
  if (!isRecord(hit)) malformed("retrieval hit");
  const out: Record<string, unknown> = {};
  const webUrl = hit["webUrl"];
  if (!absent(webUrl)) {
    if (typeof webUrl !== "string" || webUrl.length > MAX_WEB_URL) malformed("retrieval webUrl");
    if (/^https?:\/\//i.test(webUrl)) out["webUrl"] = webUrl; // never surface other schemes
  }
  const resourceType = hit["resourceType"];
  if (typeof resourceType === "string" && RESOURCE_TYPES.includes(resourceType)) {
    out["resourceType"] = resourceType; // unknownFutureValue and other members are dropped
  }
  const extracts = absent(hit["extracts"]) ? [] : hit["extracts"];
  if (!Array.isArray(extracts) || extracts.length > MAX_EXTRACTS_PER_HIT) {
    malformed("retrieval extracts list");
  }
  out["extracts"] = extracts.map(projectExtract);
  out["resourceMetadata"] = projectMetadata(hit["resourceMetadata"], requested);
  const label = projectLabel(hit["sensitivityLabel"]);
  if (label !== undefined) out["sensitivityLabel"] = label;
  return out;
}

/**
 * Projects a `retrievalResponse`. `retrievalHits` must be an array no longer than the requested
 * maximum (an empty array is a genuine empty result); anything else fails closed.
 */
export function projectRetrievalResponse(
  data: unknown,
  maxHits: number,
  requestedMetadata: readonly string[],
): Record<string, unknown> {
  if (!isRecord(data)) malformed("retrieval response");
  const hits = data["retrievalHits"];
  if (!Array.isArray(hits) || hits.length > maxHits) malformed("retrieval hits list");
  return { retrievalHits: hits.map((h) => projectHit(h, requestedMetadata)) };
}

// ------------------------------------------------------------------ tool

/** `copilot-retrieve`. Not batch-reachable: no `batchPolicy`. */
export function registerCopilotRetrieveTool(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const pathPattern = postPath(catalog, "copilot-retrieve");
  const entry = catalog.find((e) => e.toolName === "copilot-retrieve");
  if (entry === undefined || entry["apiVersion"] === "beta") {
    throw new Error("Catalog has no v1.0 POST entry for copilot-retrieve.");
  }

  handlers.set("copilot-retrieve", {
    inputSchema: COPILOT_RETRIEVE_SHAPE,
    handler: async (args) => {
      try {
        const a = args as {
          queryString: string;
          dataSource: (typeof COPILOT_DATA_SOURCES)[number];
          filterExpression?: string;
          resourceMetadata?: string[];
          maximumNumberOfResults?: number;
        };
        const body: Record<string, unknown> = {
          queryString: a.queryString,
          dataSource: a.dataSource,
        };
        if (a.filterExpression !== undefined) body["filterExpression"] = a.filterExpression;
        if (a.resourceMetadata !== undefined) body["resourceMetadata"] = [...a.resourceMetadata];
        if (a.maximumNumberOfResults !== undefined) {
          body["maximumNumberOfResults"] = a.maximumNumberOfResults;
        }
        const res = await client.request({ pathPattern, method: "post", body });
        const data = projectRetrievalResponse(
          res.data,
          a.maximumNumberOfResults ?? MAX_RESULTS,
          a.resourceMetadata ?? [],
        );
        const meta: Record<string, unknown> = { clientRequestId: res.clientRequestId };
        if (res.requestId !== undefined) meta["requestId"] = res.requestId;
        return successResult(data, meta);
      } catch (err) {
        return failure(err);
      }
    },
  });
}
