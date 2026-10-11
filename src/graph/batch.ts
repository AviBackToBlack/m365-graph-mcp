import type { EndpointCatalog, EndpointDefinition } from "../catalog.js";
import type { ErrorCode } from "../envelope.js";
import { GraphClientError, redactToken, sanitizeGraphText } from "./errors.js";
import { validateExtraHeaders } from "./headers.js";
import { validateRelativeGraphUrl, type GraphVersion } from "./url.js";

/** Graph's documented JSON batch ceiling. */
export const MAX_BATCH_REQUESTS = 20;

const BATCH_ID_RE = /^[A-Za-z0-9_-]{1,64}$/;
const REQUEST_KEYS: ReadonlySet<string> = new Set(["id", "method", "url", "headers"]);

/** One validated subrequest. Only the validator constructs these. */
export interface ValidatedSubrequest {
  readonly id: string;
  readonly url: string;
  readonly headers: Readonly<Record<string, string>>;
  /**
   * The target's response projector (see `BatchPolicy.project`) bound to this subrequest's
   * query, applied to a 2xx JSON body.
   */
  readonly project?: (body: unknown) => unknown;
  /** Catalog entry the url matched (used to interpret its response). */
  readonly entry: EndpointDefinition;
}

declare const validatedBatchBrand: unique symbol;

/**
 * Opaque, validator-produced batch. `GraphClient.batch` accepts nothing else; the WeakSet
 * makes the brand real at runtime, so a structurally similar object cannot be forged.
 */
export interface ValidatedBatch {
  readonly [validatedBatchBrand]: true;
  readonly version: GraphVersion;
  readonly requests: readonly ValidatedSubrequest[];
}

const issued = new WeakSet<object>();

export function isValidatedBatch(value: unknown): value is ValidatedBatch {
  return typeof value === "object" && value !== null && issued.has(value);
}

function invalid(message: string): GraphClientError {
  return new GraphClientError("invalid_input", message);
}

/**
 * A batchable GET tool's batch contract. `check` receives the subrequest's decoded query
 * pairs and its (already canonicalized, shared-allowlist-checked) headers, and throws a
 * `GraphClientError` (invalid_input) for anything its single-request tool would not accept
 * (unknown parameters, unbounded `$top`, `$expand`, a `Prefer` value the tool never sends...).
 * Batch must never bypass the per-tool limits in `SECURITY_INVARIANTS.md`.
 * `project`, if present, reshapes each successful JSON sub-response body (after the token
 * scan), given the same decoded query pairs `check` accepted (e.g. to clip to `$top`), so a
 * tool's response-shaping guarantee also holds in batch.
 */
export interface BatchPolicy {
  check(
    query: ReadonlyArray<readonly [string, string]>,
    headers: Readonly<Record<string, string>>,
  ): void;
  project?(body: unknown, query: ReadonlyArray<readonly [string, string]>): unknown;
}

/** Batchable endpoints and the query policy each one enforces. */
export type BatchTargets = ReadonlyMap<EndpointDefinition, BatchPolicy>;

/**
 * The endpoints a batch subrequest may target: catalog GET entries whose tool is
 * implemented AND explicitly opts in by declaring a `batchPolicy`
 * (`TOOL_SURFACE.md`: "an implemented allowlisted GET endpoint"). Handler presence alone is
 * not enough, because it says nothing about that tool's query contract.
 */
export function batchTargets(
  catalog: EndpointCatalog,
  handlers: ReadonlyMap<string, { batchPolicy?: BatchPolicy }>,
): BatchTargets {
  const targets = new Map<EndpointDefinition, BatchPolicy>();
  for (const entry of catalog) {
    const policy = handlers.get(entry.toolName)?.batchPolicy;
    if (entry.method === "get" && policy !== undefined) targets.set(entry, policy);
  }
  return targets;
}

/** Decoded `[name, value]` pairs of a canonical (validator-produced) relative URL. */
function queryPairs(url: string): Array<readonly [string, string]> {
  const at = url.indexOf("?");
  if (at === -1) return [];
  return url
    .slice(at + 1)
    .split("&")
    .map((pair) => {
      const eq = pair.indexOf("=");
      return [
        decodeURIComponent(pair.slice(0, eq)),
        decodeURIComponent(pair.slice(eq + 1)),
      ] as const;
    });
}

/**
 * Validates untrusted `graph-batch` input. Rules: exactly `{ requests }`; 1..20 requests;
 * each exactly `{ id, method, url, headers? }` (no `body`, no `dependsOn`: every subrequest
 * is an independent read, so sequencing and its cycle/424 semantics are never needed);
 * unique (case-insensitive, as Graph compares ids) `[A-Za-z0-9_-]{1,64}` ids; method is
 * exactly the string "GET" (allowlist, no normalization); headers limited to the shared
 * allowlist; url validated by `validateRelativeGraphUrl` against the `targets` entries at
 * `version`, then its query checked by that entry's `BatchPolicy`. The result is deeply
 * frozen, so a validated batch cannot be altered before it is sent.
 */
export function validateBatch(
  input: unknown,
  targets: BatchTargets,
  version: GraphVersion,
): ValidatedBatch {
  const permitted = [...targets.keys()];
  if (typeof input !== "object" || input === null || Array.isArray(input)) {
    throw invalid("Batch input must be an object.");
  }
  const keys = Object.keys(input);
  if (keys.length !== 1 || keys[0] !== "requests") {
    throw invalid("Batch input must contain only `requests`.");
  }
  const list = (input as { requests?: unknown }).requests;
  if (!Array.isArray(list) || list.length < 1 || list.length > MAX_BATCH_REQUESTS) {
    throw invalid(`A batch must contain 1 to ${MAX_BATCH_REQUESTS} requests.`);
  }
  const seen = new Set<string>();
  const requests: ValidatedSubrequest[] = list.map((item: unknown, index) => {
    const at = `Subrequest #${index + 1}`;
    if (typeof item !== "object" || item === null || Array.isArray(item)) {
      throw invalid(`${at} must be an object.`);
    }
    for (const key of Object.keys(item)) {
      if (!REQUEST_KEYS.has(key)) throw invalid(`${at} has an unsupported field.`);
    }
    const r = item as Record<string, unknown>;
    const id = r["id"];
    if (typeof id !== "string" || !BATCH_ID_RE.test(id)) {
      throw invalid(`${at} has an invalid id (use 1-64 of A-Z a-z 0-9 _ -).`);
    }
    if (seen.has(id.toLowerCase())) throw invalid(`${at} reuses an id.`);
    seen.add(id.toLowerCase());
    if (r["method"] !== "GET") throw invalid(`${at} method must be exactly "GET".`);
    let headers: Record<string, string>;
    try {
      headers = validateExtraHeaders(r["headers"]);
    } catch {
      throw invalid(`${at} has a header that is not permitted.`);
    }
    let checked: { url: string; entry: EndpointDefinition };
    try {
      checked = validateRelativeGraphUrl(r["url"], permitted, version);
    } catch (err) {
      if (err instanceof GraphClientError) {
        const reason = err.message.replace(/^Rejected unsafe relative URL: /, "");
        throw invalid(`${at} url rejected: ${reason}`);
      }
      throw err;
    }
    const policy = targets.get(checked.entry) as BatchPolicy;
    const pairs = Object.freeze(queryPairs(checked.url));
    try {
      policy.check(pairs, headers);
    } catch (err) {
      if (err instanceof GraphClientError)
        throw invalid(`${at} rejected by endpoint policy: ${err.message}`);
      throw err;
    }
    // The catalog entry is shared and not frozen here; it only selects error mapping.
    return Object.freeze({
      id,
      url: checked.url,
      headers: Object.freeze(headers),
      ...(policy.project === undefined
        ? {}
        : { project: (body: unknown) => policy.project?.(body, pairs) }),
      entry: checked.entry,
    });
  });
  const batch = Object.freeze({ version, requests: Object.freeze(requests) });
  issued.add(batch);
  return batch as unknown as ValidatedBatch;
}

/** Wire body for Graph's `$batch`; the method is fixed here, never taken from input. */
export function encodeBatchBody(batch: ValidatedBatch): string {
  return JSON.stringify({
    requests: batch.requests.map((r) => ({
      id: r.id,
      method: "GET",
      url: r.url,
      ...(Object.keys(r.headers).length > 0 ? { headers: r.headers } : {}),
    })),
  });
}

export interface BatchItemResult {
  id: string;
  status: number;
  /** Parsed JSON body of a successful subrequest (untrusted content). */
  body?: unknown;
  /** Present instead of a body for a non-JSON (base64) subresponse; contents are never returned. */
  bodyOmitted?: { reason: "non_json"; approxBytes: number };
  contentType?: string;
  /** Sanitized error summary for a non-2xx subresponse; raw bodies are never passed through. */
  error?: {
    code: ErrorCode;
    graphCode?: string;
    graphMessage?: string;
    requestId?: string;
    retryAfter?: string;
  };
}

function malformed(reason: string): GraphClientError {
  return new GraphClientError(
    "graph_error",
    `Graph returned an invalid batch response: ${reason}.`,
  );
}

function ciGet(headers: unknown, name: string): string | undefined {
  if (typeof headers !== "object" || headers === null || Array.isArray(headers)) return undefined;
  for (const [key, value] of Object.entries(headers)) {
    if (key.toLowerCase() === name && typeof value === "string") return value;
  }
  return undefined;
}

function errorCodeFor(status: number, entry: EndpointDefinition): ErrorCode {
  if (status === 401) return "reauth_required";
  if (status === 403) return "capability_gap";
  if (status === 429) return "rate_limited";
  if (status === 410 && entry.pathPattern.endsWith("/delta()")) return "restart_required";
  return "graph_error";
}

/**
 * Normalizes an (already token-echo-scanned) batch response. Subresponses are untrusted:
 * the structure must match the request exactly (no unknown, duplicate, or missing ids),
 * and only a fixed set of fields leaves this function (never Location or other headers,
 * raw error bodies, or base64 payloads).
 */
export function normalizeBatchResponse(
  data: unknown,
  batch: ValidatedBatch,
  token: string,
): { responses: BatchItemResult[] } {
  const list = (data as { responses?: unknown } | null | undefined)?.responses;
  if (typeof data !== "object" || data === null || !Array.isArray(list)) {
    throw malformed("missing responses array");
  }
  if (list.length > batch.requests.length) throw malformed("more responses than requests");
  const byId = new Map(batch.requests.map((r) => [r.id.toLowerCase(), r]));
  const results = new Map<string, BatchItemResult>();
  for (const raw of list as unknown[]) {
    if (typeof raw !== "object" || raw === null || Array.isArray(raw)) {
      throw malformed("malformed subresponse");
    }
    const sub = raw as Record<string, unknown>;
    const sid = sub["id"];
    const request = typeof sid === "string" ? byId.get(sid.toLowerCase()) : undefined;
    if (request === undefined) throw malformed("subresponse id does not match any request");
    if (results.has(request.id)) throw malformed("duplicate subresponse id");
    const status = sub["status"];
    if (typeof status !== "number" || !Number.isInteger(status) || status < 100 || status > 599) {
      throw malformed("invalid subresponse status");
    }
    const item: BatchItemResult = { id: request.id, status };
    const contentType = sanitizeGraphText(ciGet(sub["headers"], "content-type"), 100);
    if (contentType !== undefined) item.contentType = redactToken(contentType, token);
    const body = sub["body"];
    if (status >= 200 && status < 300) {
      if (typeof body === "string") {
        // Non-JSON subresponses arrive base64-encoded; never relay opaque payloads.
        item.bodyOmitted = { reason: "non_json", approxBytes: Math.floor((body.length * 3) / 4) };
      } else if (body !== undefined && body !== null) {
        item.body = request.project === undefined ? body : request.project(body);
      }
    } else {
      const err = (body as { error?: { code?: unknown; message?: unknown; innerError?: unknown } })
        ?.error;
      const out: NonNullable<BatchItemResult["error"]> = {
        code: errorCodeFor(status, request.entry),
      };
      const graphCode = sanitizeGraphText(err?.code, 64);
      const graphMessage = sanitizeGraphText(err?.message);
      if (graphCode !== undefined) out.graphCode = redactToken(graphCode, token);
      if (graphMessage !== undefined) out.graphMessage = redactToken(graphMessage, token);
      const inner = err?.innerError as Record<string, unknown> | null | undefined;
      const rid = inner?.["request-id"];
      if (typeof rid === "string" && /^[A-Za-z0-9._-]{1,64}$/.test(rid)) {
        out.requestId = redactToken(rid, token);
      }
      const retryAfter = sanitizeGraphText(ciGet(sub["headers"], "retry-after"), 40);
      if (retryAfter !== undefined) out.retryAfter = redactToken(retryAfter, token);
      item.error = out;
    }
    results.set(request.id, item);
  }
  if (results.size !== batch.requests.length) throw malformed("missing subresponses");
  // Deterministic order: the request order, not Graph's arbitrary order.
  return { responses: batch.requests.map((r) => results.get(r.id) as BatchItemResult) };
}
