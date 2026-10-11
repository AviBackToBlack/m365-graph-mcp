import { randomUUID } from "node:crypto";
import type { EndpointCatalog, EndpointDefinition } from "../catalog.js";
import { TokenError, type TokenProvider } from "../auth/token-provider.js";
import { GraphClientError, redactToken, sanitizeGraphText } from "./errors.js";
import {
  encodeBatchBody,
  isValidatedBatch,
  normalizeBatchResponse,
  type BatchItemResult,
  type ValidatedBatch,
} from "./batch.js";
import { validateExtraHeaders } from "./headers.js";
import {
  assertFixedOrigin,
  buildPath,
  buildUrl,
  endpointVersion,
  refersToBatch,
  validateGraphLink,
  type GraphVersion,
  type QueryParams,
} from "./url.js";

export type GraphMethod = "get" | "post";
export type FetchLike = (url: string, init: RequestInit) => Promise<Response>;

export const DEFAULTS = {
  timeoutMs: 30_000,
  maxResponseBytes: 5 * 1024 * 1024,
  maxRequestBytes: 256 * 1024,
  maxRetries: 3,
  maxRetryAfterMs: 30_000,
  maxTotalWaitMs: 60_000,
  maxPages: 10,
  maxItems: 1000,
} as const;

export interface GraphClientOptions {
  tokenProvider: TokenProvider;
  /** The loaded catalog; the only source of allowed (method, pathPattern) pairs. */
  catalog: EndpointCatalog;
  fetch?: FetchLike;
  now?: () => number;
  sleep?: (ms: number) => Promise<void>;
  timeoutMs?: number;
  maxResponseBytes?: number;
  maxRequestBytes?: number;
  maxRetries?: number;
  maxRetryAfterMs?: number;
  maxTotalWaitMs?: number;
  maxPages?: number;
  maxItems?: number;
}

export interface GraphRequest {
  /** Exact catalog `pathPattern`, for example `/me/messages/{message-id}`. */
  pathPattern: string;
  /** Defaults to `get`. `post` only works for catalog POST entries. */
  method?: GraphMethod;
  /** Values for catalog `{placeholders}`, using the catalog's own placeholder names. */
  pathParams?: Readonly<Record<string, string>>;
  query?: QueryParams;
  /** JSON body; POST only. */
  body?: unknown;
  /**
   * Must equal the catalog entry's version (`beta` only for entries with
   * `apiVersion: "beta"`). Omit to use the catalog's version.
   */
  version?: GraphVersion;
  /** Only `ConsistencyLevel` and `Prefer` are accepted. */
  headers?: Readonly<Record<string, string>>;
}

export interface GraphResponse {
  status: number;
  /** Parsed JSON body (`undefined` for an empty body). Untrusted content. */
  data: unknown;
  /** Raw `@odata.nextLink` from the body; pass it only to `fetchLink`/`paginate`. */
  nextLink?: string;
  /** Raw `@odata.deltaLink` from the body. Opaque capability; pass it to `fetchLink`. */
  deltaLink?: string;
  requestId?: string;
  clientRequestId: string;
}

export interface PaginateOptions {
  maxPages?: number;
  maxItems?: number;
}

export interface PaginatedResult {
  items: unknown[];
  pages: number;
  /** True when a page/item ceiling stopped collection before the end of the data. */
  truncated: boolean;
  /** Present when more pages exist and whole pages were consumed; resume via `fetchLink`. */
  nextLink?: string;
  /** Present when a delta sequence completed; store it for the next incremental sync. */
  deltaLink?: string;
}

interface ResolvedTarget {
  entry: EndpointDefinition;
  method: GraphMethod;
  version: GraphVersion;
}

/** Hook applied to the parsed, token-scanned success body; returns the data to hand back. */
type PostProcess = (data: unknown, token: string) => unknown;

export interface BatchResult {
  /** One normalized entry per request, in request order. Untrusted content. */
  responses: BatchItemResult[];
  requestId?: string;
  clientRequestId: string;
}

const HTTP_DATE_RE = /^[A-Za-z]{3}, \d{2} [A-Za-z]{3} \d{4} \d{2}:\d{2}:\d{2} GMT$/;

const isDeltaPattern = (pattern: string): boolean => pattern.endsWith("/delta()");

function batchRefused(): GraphClientError {
  return new GraphClientError(
    "invalid_input",
    "$batch is not available through the generic request path.",
  );
}

function describe(entry: EndpointDefinition, method: GraphMethod): string {
  return `${method.toUpperCase()} ${entry.pathPattern}`;
}

function positiveInt(value: number | undefined, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0
    ? Math.floor(value)
    : fallback;
}

export class GraphClient {
  private readonly tokenProvider: TokenProvider;
  private readonly index = new Map<string, EndpointDefinition>();
  private readonly fetchImpl: FetchLike;
  private readonly now: () => number;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly timeoutMs: number;
  private readonly maxResponseBytes: number;
  private readonly maxRequestBytes: number;
  private readonly maxRetries: number;
  private readonly maxRetryAfterMs: number;
  private readonly maxTotalWaitMs: number;
  private readonly maxPages: number;
  private readonly maxItems: number;

  constructor(options: GraphClientOptions) {
    this.tokenProvider = options.tokenProvider;
    for (const entry of options.catalog) {
      this.index.set(`${entry.method} ${entry.pathPattern}`, entry);
    }
    this.fetchImpl = options.fetch ?? ((url, init) => fetch(url, init));
    this.now = options.now ?? Date.now;
    this.sleep = options.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.timeoutMs = positiveInt(options.timeoutMs, DEFAULTS.timeoutMs);
    this.maxResponseBytes = positiveInt(options.maxResponseBytes, DEFAULTS.maxResponseBytes);
    this.maxRequestBytes = positiveInt(options.maxRequestBytes, DEFAULTS.maxRequestBytes);
    this.maxRetries = positiveInt(options.maxRetries, DEFAULTS.maxRetries);
    this.maxRetryAfterMs = positiveInt(options.maxRetryAfterMs, DEFAULTS.maxRetryAfterMs);
    this.maxTotalWaitMs = positiveInt(options.maxTotalWaitMs, DEFAULTS.maxTotalWaitMs);
    this.maxPages = positiveInt(options.maxPages, DEFAULTS.maxPages);
    this.maxItems = positiveInt(options.maxItems, DEFAULTS.maxItems);
  }

  /** Issues one allowlisted request. Throws `GraphClientError` (envelope-coded) on failure. */
  async request(req: GraphRequest): Promise<GraphResponse> {
    const target = this.resolve(req);
    const path = buildPath(target.entry, req.pathParams);
    if (refersToBatch(path)) throw batchRefused();
    const url = buildUrl(target.version, path, req.query);
    const body = this.encodeBody(req, target);
    const headers = this.extraHeaders(req.headers);
    return this.send(target, url, body, headers);
  }

  /**
   * Sends a validator-produced batch through the same `send()` pipeline as every other
   * request (token, timeout, retries, size cap, origin assertion, token-echo scan). Accepts
   * only a `ValidatedBatch`; `request()` still refuses `$batch`.
   */
  async batch(validated: ValidatedBatch): Promise<BatchResult> {
    if (!isValidatedBatch(validated)) {
      throw new GraphClientError("invalid_input", "Batch was not produced by the batch validator.");
    }
    const entry = this.index.get("post /$batch");
    if (entry === undefined) {
      throw new GraphClientError("invalid_input", "Request is not in the endpoint catalog.");
    }
    const version = endpointVersion(entry);
    if (validated.version !== version) {
      throw new GraphClientError("invalid_input", "Batch API version does not match the catalog.");
    }
    const body = encodeBatchBody(validated);
    if (Buffer.byteLength(body, "utf8") > this.maxRequestBytes) {
      throw new GraphClientError("invalid_input", "Request body exceeds the size limit.", {
        maxRequestBytes: this.maxRequestBytes,
      });
    }
    const target: ResolvedTarget = { entry, method: "post", version };
    const url = buildUrl(version, "/$batch", undefined);
    const res = await this.send(target, url, body, {}, (data, token) =>
      normalizeBatchResponse(data, validated, token),
    );
    const out: BatchResult = {
      responses: (res.data as { responses: BatchItemResult[] }).responses,
      clientRequestId: res.clientRequestId,
    };
    if (res.requestId !== undefined) out.requestId = res.requestId;
    return out;
  }

  /**
   * Follows a `@odata.nextLink`/`@odata.deltaLink` produced by `base`. The link is validated
   * against the fixed origin, the base request's version, and the base request's catalog
   * endpoint, then rebuilt; the raw string is never fetched. GET endpoints only.
   */
  async fetchLink(base: GraphRequest, link: string): Promise<GraphResponse> {
    const target = this.resolve(base);
    if (target.method !== "get") {
      throw new GraphClientError("invalid_input", "Links can only be followed for GET endpoints.");
    }
    const url = validateGraphLink(
      link,
      target.entry,
      target.version,
      base.pathParams,
      base.query ?? {},
    );
    // Continuations keep the original request's allowlisted headers (ConsistencyLevel, Prefer).
    return this.send(target, url, undefined, this.extraHeaders(base.headers));
  }

  /**
   * Collects `value` items across pages with page and item ceilings. For delta endpoints
   * the final `deltaLink` is returned. Every follow-up page goes through `fetchLink`.
   */
  async paginate(base: GraphRequest, options: PaginateOptions = {}): Promise<PaginatedResult> {
    const maxPages = Math.max(1, Math.min(positiveInt(options.maxPages, this.maxPages), 1000));
    const maxItems = Math.max(1, Math.min(positiveInt(options.maxItems, this.maxItems), 100_000));
    const items: unknown[] = [];
    const seen = new Set<string>();
    let pages = 0;
    let response = await this.request(base);
    for (;;) {
      pages += 1;
      const value = (response.data as { value?: unknown } | null | undefined)?.value;
      // A successful collection page must carry a `value` array; anything else fails closed
      // instead of being reported as an empty (or shorter) listing.
      if (!Array.isArray(value)) {
        throw new GraphClientError(
          "graph_error",
          "Graph returned a collection page without a value array.",
        );
      }
      // Append only the remaining budget; never spread an unbounded array into push().
      const room = maxItems - items.length;
      const take = Math.min(value.length, room);
      for (let i = 0; i < take; i += 1) items.push(value[i]);
      if (value.length > room) {
        return { items, pages, truncated: true };
      }
      const next = response.nextLink;
      if (next === undefined) {
        const done: PaginatedResult = { items, pages, truncated: false };
        if (response.deltaLink !== undefined) done.deltaLink = response.deltaLink;
        return done;
      }
      if (items.length >= maxItems || pages >= maxPages) {
        return { items, pages, truncated: true, nextLink: next };
      }
      if (seen.has(next)) {
        throw new GraphClientError("graph_error", "Graph returned a repeating nextLink.");
      }
      seen.add(next);
      response = await this.fetchLink(base, next);
    }
  }

  // ---------------------------------------------------------------- internals

  private resolve(req: GraphRequest): ResolvedTarget {
    if (refersToBatch(String(req.pathPattern))) throw batchRefused();
    const method = req.method ?? "get";
    if (method !== "get" && method !== "post") {
      throw new GraphClientError("invalid_input", "Unsupported HTTP method.");
    }
    const entry = this.index.get(`${method} ${req.pathPattern}`);
    if (entry === undefined) {
      throw new GraphClientError("invalid_input", "Request is not in the endpoint catalog.");
    }
    const catalogVersion = endpointVersion(entry);
    const version = req.version ?? catalogVersion;
    if (version !== catalogVersion) {
      throw new GraphClientError(
        "invalid_input",
        `Endpoint ${describe(entry, method)} requires API version ${catalogVersion}.`,
      );
    }
    return { entry, method, version };
  }

  private encodeBody(req: GraphRequest, target: ResolvedTarget): string | undefined {
    if (target.method === "get") {
      if (req.body !== undefined) {
        throw new GraphClientError("invalid_input", "GET requests cannot have a body.");
      }
      return undefined;
    }
    if (req.body === undefined) {
      throw new GraphClientError("invalid_input", "POST requests require a body.");
    }
    let json: string | undefined;
    try {
      json = JSON.stringify(req.body);
    } catch {
      throw new GraphClientError("invalid_input", "Request body is not JSON-serializable.");
    }
    if (typeof json !== "string") {
      throw new GraphClientError("invalid_input", "Request body is not JSON-serializable.");
    }
    if (Buffer.byteLength(json, "utf8") > this.maxRequestBytes) {
      throw new GraphClientError("invalid_input", "Request body exceeds the size limit.", {
        maxRequestBytes: this.maxRequestBytes,
      });
    }
    return json;
  }

  private extraHeaders(
    headers: Readonly<Record<string, string>> | undefined,
  ): Record<string, string> {
    return validateExtraHeaders(headers);
  }

  private async getToken(): Promise<string> {
    let token: string;
    let timer: ReturnType<typeof setTimeout> | undefined;
    // Providers may do local I/O (e.g. the token bridge); never wait on them unbounded.
    const timedOut = new Promise<never>((_resolve, reject) => {
      timer = setTimeout(
        () =>
          reject(
            new GraphClientError(
              "reauth_required",
              "The token provider did not respond in time. Provide a fresh token.",
            ),
          ),
        this.timeoutMs,
      );
    });
    try {
      token = await Promise.race([this.tokenProvider.getAccessToken(), timedOut]);
    } catch (err) {
      if (err instanceof GraphClientError) throw err;
      if (err instanceof TokenError) {
        // Provider messages are fixed, token-free text; preserved as-is.
        throw new GraphClientError("reauth_required", err.message, { tokenError: err.code });
      }
      throw new GraphClientError("graph_error", "The token provider failed unexpectedly.");
    } finally {
      clearTimeout(timer);
    }
    // Visible ASCII only: anything else cannot travel safely in an HTTP header value.
    if (typeof token !== "string" || token.length === 0 || /[^\x21-\x7e]/.test(token)) {
      throw new GraphClientError(
        "reauth_required",
        "The access token is unusable as a header value.",
      );
    }
    return token;
  }

  private async send(
    target: ResolvedTarget,
    url: string,
    body: string | undefined,
    extra: Record<string, string>,
    post?: PostProcess,
  ): Promise<GraphResponse> {
    const label = describe(target.entry, target.method);
    const clientRequestId = randomUUID();
    let waited = 0;

    for (let attempt = 0; ; attempt += 1) {
      assertFixedOrigin(url);
      const token = await this.getToken();
      const outcome = await this.attempt(target, url, body, extra, token, clientRequestId, post);

      if (outcome.kind === "response") {
        const res = outcome.response;
        if (res.status >= 200 && res.status < 300) return res.parsed(clientRequestId);
        const retriable = res.status === 429 || res.status === 503 || res.status === 504;
        if (retriable && attempt < this.maxRetries) {
          const wait = this.retryDelayMs(res.retryAfter, attempt);
          if (wait <= this.maxRetryAfterMs && waited + wait <= this.maxTotalWaitMs) {
            waited += wait;
            await this.sleep(wait);
            continue;
          }
        }
        throw this.normalize(target, label, res, clientRequestId, token);
      }
      throw outcome.error;
    }
  }

  private retryDelayMs(retryAfter: string | undefined, attempt: number): number {
    if (retryAfter !== undefined) {
      const trimmed = retryAfter.trim();
      if (/^\d+$/.test(trimmed)) return Number(trimmed) * 1000;
      // Only an IMF-fixdate counts: V8 parses bare numerics such as "5.5" as a past date,
      // which would turn a throttle into an immediate retry (zero wait).
      if (HTTP_DATE_RE.test(trimmed)) {
        const date = Date.parse(trimmed);
        if (Number.isFinite(date)) return Math.max(0, date - this.now());
      }
    }
    return Math.min(1000 * 2 ** attempt, this.maxRetryAfterMs);
  }

  private async attempt(
    target: ResolvedTarget,
    url: string,
    body: string | undefined,
    extra: Record<string, string>,
    token: string,
    clientRequestId: string,
    post: PostProcess | undefined,
  ): Promise<
    { kind: "error"; error: GraphClientError } | { kind: "response"; response: ReadResponse }
  > {
    const label = describe(target.entry, target.method);
    const controller = new AbortController();
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      controller.abort();
    }, this.timeoutMs);
    try {
      const headers: Record<string, string> = {
        ...extra,
        Accept: "application/json",
        "client-request-id": clientRequestId,
        Authorization: `Bearer ${token}`,
      };
      if (body !== undefined) headers["Content-Type"] = "application/json";
      const init: RequestInit = {
        method: target.method.toUpperCase(),
        headers,
        redirect: "manual",
        signal: controller.signal,
      };
      if (body !== undefined) init.body = body;
      const res = await this.fetchImpl(url, init);

      if (res.type === "opaqueredirect" || (res.status >= 300 && res.status < 400)) {
        void res.body?.cancel().catch(() => undefined);
        return {
          kind: "error",
          error: new GraphClientError(
            "graph_error",
            "Graph returned a redirect; redirects are never followed.",
            {
              endpoint: label,
              status: res.status,
            },
          ),
        };
      }

      const isSuccess = res.status >= 200 && res.status < 300;
      const read = await this.readBody(res, isSuccess);
      if (read.kind === "oversize") {
        return {
          kind: "error",
          error: new GraphClientError("graph_error", "Graph response exceeded the size limit.", {
            endpoint: label,
            maxResponseBytes: this.maxResponseBytes,
          }),
        };
      }
      return {
        kind: "response",
        response: {
          status: res.status,
          retryAfter: res.headers.get("retry-after") ?? undefined,
          requestId: safeId(res.headers.get("request-id"), token),
          echoedClientRequestId: safeId(res.headers.get("client-request-id"), token),
          text: read.text,
          parsed: (id: string) => this.parseSuccess(res, read.text, label, id, token, post),
        },
      };
    } catch (err) {
      if (err instanceof GraphClientError) return { kind: "error", error: err };
      if (timedOut) {
        return {
          kind: "error",
          error: new GraphClientError("graph_error", "Graph request timed out.", {
            endpoint: label,
            timeoutMs: this.timeoutMs,
          }),
        };
      }
      return {
        kind: "error",
        error: new GraphClientError("graph_error", "Network error contacting Graph.", {
          endpoint: label,
          networkCode: networkCode(err),
        }),
      };
    } finally {
      clearTimeout(timer);
    }
  }

  private async readBody(
    res: Response,
    strict: boolean,
  ): Promise<{ kind: "ok"; text: string } | { kind: "oversize" }> {
    const declared = Number(res.headers.get("content-length"));
    if (Number.isFinite(declared) && declared > this.maxResponseBytes) {
      void res.body?.cancel().catch(() => undefined);
      return strict ? { kind: "oversize" } : { kind: "ok", text: "" };
    }
    if (res.body === null) return { kind: "ok", text: "" };
    const reader = res.body.getReader();
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > this.maxResponseBytes) {
        await reader.cancel().catch(() => undefined);
        return strict ? { kind: "oversize" } : { kind: "ok", text: "" };
      }
      chunks.push(value);
    }
    return { kind: "ok", text: Buffer.concat(chunks).toString("utf8") };
  }

  private parseSuccess(
    res: Response,
    text: string,
    label: string,
    clientRequestId: string,
    token: string,
    post: PostProcess | undefined,
  ): GraphResponse {
    let data: unknown;
    if (text.trim().length > 0) {
      try {
        // A UTF-8 BOM is not JSON whitespace; tolerate it rather than reject the response.
        data = JSON.parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text);
      } catch {
        throw new GraphClientError("graph_error", "Graph returned a non-JSON response.", {
          endpoint: label,
        });
      }
    }
    if (containsToken(data, token)) {
      // Never hand an echoed live token back through a tool result (or its links).
      throw new GraphClientError("graph_error", "Graph response contained credential material.", {
        endpoint: label,
      });
    }
    // Runs only after the token-echo scan above; may throw a GraphClientError.
    if (post !== undefined) data = post(data, token);
    const out: GraphResponse = { status: res.status, data, clientRequestId };
    const obj = typeof data === "object" && data !== null ? (data as Record<string, unknown>) : {};
    if (typeof obj["@odata.nextLink"] === "string") out.nextLink = obj["@odata.nextLink"];
    if (typeof obj["@odata.deltaLink"] === "string") out.deltaLink = obj["@odata.deltaLink"];
    const requestId = safeId(res.headers.get("request-id"), token);
    if (requestId !== undefined) out.requestId = requestId;
    return out;
  }

  private normalize(
    target: ResolvedTarget,
    label: string,
    res: ReadResponse,
    clientRequestId: string,
    token: string,
  ): GraphClientError {
    let graphCode: string | undefined;
    let graphMessage: string | undefined;
    try {
      const parsed = JSON.parse(res.text) as { error?: { code?: unknown; message?: unknown } };
      graphCode = sanitizeGraphText(parsed?.error?.code, 64);
      graphMessage = sanitizeGraphText(parsed?.error?.message);
    } catch {
      // Non-JSON error bodies are never echoed.
    }
    const details: Record<string, unknown> = { endpoint: label, status: res.status };
    if (graphCode !== undefined) details["graphCode"] = redactToken(graphCode, token);
    if (graphMessage !== undefined) details["graphMessage"] = redactToken(graphMessage, token);
    if (res.requestId !== undefined) details["requestId"] = res.requestId;
    details["clientRequestId"] = res.echoedClientRequestId ?? clientRequestId;

    if (res.status === 401) {
      return new GraphClientError(
        "reauth_required",
        "Microsoft Graph rejected the access token (401). Provide a fresh token.",
        details,
      );
    }
    if (res.status === 403) {
      return new GraphClientError(
        "capability_gap",
        "Microsoft Graph denied access (403): the token lacks a required permission or a tenant policy blocks it. Report for human review; do not broaden scopes automatically.",
        details,
      );
    }
    if (res.status === 410 && isDeltaPattern(target.entry.pathPattern)) {
      return new GraphClientError(
        "restart_required",
        "The delta token expired or is invalid (410 Gone). Restart the delta sync from the beginning.",
        details,
      );
    }
    if (res.status === 429) {
      const ra = sanitizeGraphText(res.retryAfter, 40);
      if (ra !== undefined) details["retryAfter"] = redactToken(ra, token);
      return new GraphClientError(
        "rate_limited",
        "Microsoft Graph throttled the request (429) and retries were exhausted or too long.",
        details,
      );
    }
    return new GraphClientError(
      "graph_error",
      `Microsoft Graph returned HTTP ${res.status}.`,
      details,
    );
  }
}

interface ReadResponse {
  status: number;
  retryAfter: string | undefined;
  requestId: string | undefined;
  echoedClientRequestId: string | undefined;
  text: string;
  parsed: (clientRequestId: string) => GraphResponse;
}

/** True when any string key or value in parsed JSON contains the live token (iterative walk). */
function containsToken(data: unknown, token: string): boolean {
  const stack: unknown[] = [data];
  while (stack.length > 0) {
    const node = stack.pop();
    if (typeof node === "string") {
      if (node.includes(token)) return true;
    } else if (Array.isArray(node)) {
      for (const item of node) stack.push(item);
    } else if (typeof node === "object" && node !== null) {
      for (const [key, value] of Object.entries(node)) {
        if (key.includes(token)) return true;
        stack.push(value);
      }
    }
  }
  return false;
}

/** Request ids are short opaque identifiers; anything else (or token-like) is dropped. */
function safeId(value: string | null, token: string): string | undefined {
  if (value === null || !/^[A-Za-z0-9._-]{1,64}$/.test(value)) return undefined;
  return redactToken(value, token);
}

function networkCode(err: unknown): string | undefined {
  const candidates: unknown[] = [
    (err as { code?: unknown } | null)?.code,
    (err as { cause?: { code?: unknown } } | null)?.cause?.code,
  ];
  for (const c of candidates) {
    if (typeof c === "string" && /^[A-Z0-9_]{3,40}$/.test(c)) return c;
  }
  return undefined;
}

export { GRAPH_ORIGIN } from "./url.js";
