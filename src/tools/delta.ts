import { z } from "zod";
import type { EndpointDefinition } from "../catalog.js";
import { errorResult, successResult } from "../envelope.js";
import { GraphClientError } from "../graph/errors.js";
import type { GraphClient, GraphRequest } from "../graph/client.js";
import { endpointVersion, validateGraphLink } from "../graph/url.js";
import type { ToolArgs, ToolImplementation } from "../registry.js";

/**
 * Generic opaque-token delta tool (SM-6 mail, SM-8 calendarView). A tool supplies its catalog
 * entry, its initial-only options and how they become the initial query; this module owns
 * everything security-relevant, identically for every delta tool.
 *
 * Continuation is an opaque, typed TOKEN, never a Graph URL. Graph returns
 * `@odata.nextLink` (`$skiptoken`, more pages in this round) or `@odata.deltaLink`
 * (`$deltatoken`, round complete). Each returned link is validated (fixed origin, same
 * catalog endpoint, same path principal; `validateGraphLink`), reduced to its single token
 * and discarded. The caller passes the token back as `skipToken` / `deltaToken`, and EVERY
 * request (including the pages followed inside one call) is rebuilt from the fixed catalog
 * path + the validated path parameters + that one token parameter. A Graph-supplied link is
 * never fetched, so nothing Graph (or an attacker-influenced response) puts in a link
 * besides the token can reach a request.
 *
 * Per Graph docs the OData options are sent on the initial request only; the state token
 * encodes them. A token call therefore rejects the tool's initial-only options (fail closed)
 * instead of ignoring them silently. Headers are not part of the token, so
 * `Prefer: odata.maxpagesize` (and any tool-specific preference such as the response time
 * zone) is sent on every request.
 *
 * Bounds: at most MAX_PAGES requests per call and stop once `top` items are collected.
 * Ceilings are applied at PAGE boundaries only: a page is never cut, because dropping part
 * of a delta page would silently lose changes (the token only resumes after whole pages).
 * Graph honours `maxpagesize`, so a call returns at most about `top` items. Delta tools are
 * not batch-reachable (no `batchPolicy`): a batch cannot carry tokens between requests or
 * follow pages, and a collection projector would drop `@removed` markers.
 */

const MAX_PAGES = 10;
const DEFAULT_TOP = 100;
const MAX_TOP = 1000;

/**
 * Charset of a state token after percent-decoding. Graph's tokens are opaque base64-like
 * strings; anything else is refused. The cap keeps the rebuilt URL well under the 8192 limit.
 */
const TOKEN_RE = /^[A-Za-z0-9_\-.~=+/]{8,4096}$/;

const invalid = (message: string): GraphClientError =>
  new GraphClientError("invalid_input", message);

function linkRejected(reason: string): GraphClientError {
  return invalid(`Rejected unsafe Graph link: ${reason}.`);
}

function decode(part: string): string | undefined {
  try {
    return decodeURIComponent(part);
  } catch {
    return undefined;
  }
}

/**
 * The single state token of a Graph `@odata.nextLink` (`skiptoken`) or `@odata.deltaLink`
 * (`deltatoken`). The link first passes `validateGraphLink` (fixed origin, same endpoint,
 * folder-canonicalization only, bound principal). Its query must then be exactly one pair
 * with the expected name and a token the input validator would also accept. The query is
 * split manually: `URLSearchParams` would turn a raw `+` of a base64 token into a space.
 */
function extractStateToken(
  link: unknown,
  kind: "skiptoken" | "deltatoken",
  entry: EndpointDefinition,
  pathParams: Readonly<Record<string, string>>,
): string {
  const safe = validateGraphLink(link, entry, endpointVersion(entry), pathParams);
  const at = safe.indexOf("?");
  if (at === -1) throw linkRejected("no state token");
  // Exactly one pair, the token: Microsoft documents that the initial query is encoded into
  // the token and shows links carrying only $skiptoken/$deltatoken. Anything else (an echoed
  // window included) fails closed.
  const pairs = safe.slice(at + 1).split("&");
  if (pairs.length !== 1) throw linkRejected("unexpected query parameters");
  const pair = pairs[0] as string;
  const eq = pair.indexOf("=");
  const name = eq === -1 ? undefined : decode(pair.slice(0, eq));
  if (name === undefined || name.toLowerCase() !== `$${kind}`) {
    throw linkRejected("query is not the expected state token");
  }
  const value = decode(pair.slice(eq + 1));
  if (value === undefined || !TOKEN_RE.test(value)) throw linkRejected("malformed state token");
  return value;
}

/**
 * Allowlist projection that, unlike the list projectors, keeps the delta `@removed` marker
 * (reduced to a known `reason`) so a deletion never reads like an update. Applied to every
 * page, whatever the resumed token encodes.
 */
function projectChange(item: unknown, keepKeys: ReadonlySet<string>): Record<string, unknown> {
  // Every change must be identifiable and every deletion marker well-formed; otherwise the
  // whole page fails closed (before its token is returned or followed), because silently
  // dropping or degrading a change would let a caller checkpoint past it.
  if (typeof item !== "object" || item === null || Array.isArray(item)) {
    throw malformedChange();
  }
  const raw = item as Record<string, unknown>;
  const id = raw["id"];
  if (typeof id !== "string" || id.length === 0) throw malformedChange();
  const out: Record<string, unknown> = Object.fromEntries(
    Object.entries(raw).filter(([k]) => keepKeys.has(k)),
  );
  out["id"] = id;
  if ("@removed" in raw) {
    const removed = raw["@removed"];
    const reason =
      typeof removed === "object" && removed !== null && !Array.isArray(removed)
        ? (removed as Record<string, unknown>)["reason"]
        : undefined;
    if (reason !== "deleted" && reason !== "changed") throw malformedChange();
    out["@removed"] = { reason };
  }
  return out;
}

function malformedChange(): GraphClientError {
  return new GraphClientError("graph_error", "Graph returned a malformed delta change.");
}

const SYNC_RESET_CODE_RE = /^(?:syncstate|resync)/i;

/** An expired/invalid delta state can surface as 410 (mapped by the client) or a 4xx code. */
function failure(toolName: string, err: unknown): ReturnType<typeof errorResult> {
  if (err instanceof GraphClientError) {
    const status = err.details?.["status"];
    const code = err.details?.["graphCode"];
    if (
      err.code === "graph_error" &&
      typeof status === "number" &&
      status >= 400 &&
      status < 500 &&
      typeof code === "string" &&
      SYNC_RESET_CODE_RE.test(code)
    ) {
      return errorResult(
        "restart_required",
        "The delta token expired or is invalid. Restart the delta sync from the beginning (call without deltaToken or skipToken).",
        err.details,
      );
    }
    return err.toResult();
  }
  return errorResult("graph_error", `Tool ${toolName} failed unexpectedly.`);
}

export interface DeltaToolSpec {
  toolName: string;
  /** The catalog GET entry; its `pathPattern` is the only path ever requested. */
  entry: EndpointDefinition;
  /** Names of the entry's `{placeholders}`; each becomes a required string input. */
  pathParamNames: readonly string[];
  /** Options valid on the initial call only (the state token encodes them afterwards). */
  initialShape: z.ZodRawShape;
  /** Options accepted on every call, token calls included (header-level, not in the token). */
  alwaysShape?: z.ZodRawShape;
  /** Error text when an initial-only option accompanies a token. */
  initialOnlyMessage: string;
  /** Names the initial-only options in the token input descriptions. */
  initialOnlyList: string;
  /** Initial query from the parsed args; throws `invalid_input` for bad/missing options. */
  buildInitialQuery(typed: Readonly<Record<string, unknown>>): Record<string, string>;
  /** Extra `Prefer` preferences sent on EVERY request (after `odata.maxpagesize`). */
  preferences?(typed: Readonly<Record<string, unknown>>): string[];
  /** Properties a returned change may keep (`id` and annotations included by the caller). */
  keepKeys: ReadonlySet<string>;
  /** Largest `odata.maxpagesize` requested. */
  pageSizeMax: number;
}

export function createDeltaTool(client: GraphClient, spec: DeltaToolSpec): ToolImplementation {
  const { entry, toolName } = spec;
  if (entry.method !== "get") throw new Error(`Catalog has no GET entry for ${toolName}.`);
  const pathPattern = entry.pathPattern;

  const tokenSchema = (what: string) =>
    z
      .string()
      .regex(TOKEN_RE)
      .optional()
      .describe(
        `${what} Opaque value copied unchanged from a previous result of this tool; never a URL. Mutually exclusive with the other token and with ${spec.initialOnlyList}.`,
      );
  const common = {
    top: z
      .number()
      .int()
      .min(1)
      .max(MAX_TOP)
      .optional()
      .describe(
        `Approximate maximum changes to return per call (1-${MAX_TOP}; default ${DEFAULT_TOP}); also the page size (max ${spec.pageSizeMax}). Pages are never cut: when the limit is reached the result has hasMore=true and a skipToken.`,
      ),
    deltaToken: tokenSchema(
      "Continue with changes since the last COMPLETED round (meta.deltaToken).",
    ),
    skipToken: tokenSchema(
      "Continue the CURRENT round (meta.skipToken, when meta.hasMore is true).",
    ),
  } satisfies z.ZodRawShape;
  const shape: z.ZodRawShape = {
    ...spec.initialShape,
    ...common,
    ...(spec.alwaysShape ?? {}),
  };
  const pathShape: z.ZodRawShape = Object.fromEntries(
    spec.pathParamNames.map((n) => [n, z.string().min(1)]),
  );
  const argsSchema = z.strictObject({ ...shape, ...pathShape });
  const initialKeys = Object.keys(spec.initialShape);

  const run = async (args: ToolArgs) => {
    const parsed = argsSchema.safeParse(args);
    if (!parsed.success) throw invalid("Invalid delta options.");
    const typed = parsed.data as Record<string, unknown>;
    const deltaToken = typed["deltaToken"] as string | undefined;
    const skipToken = typed["skipToken"] as string | undefined;
    if (deltaToken !== undefined && skipToken !== undefined) {
      throw invalid("deltaToken and skipToken are mutually exclusive.");
    }
    const resume = deltaToken ?? skipToken;
    if (resume !== undefined && initialKeys.some((k) => typed[k] !== undefined)) {
      throw invalid(spec.initialOnlyMessage);
    }
    const pathParams = Object.fromEntries(spec.pathParamNames.map((n) => [n, typed[n] as string]));
    const top = (typed["top"] as number | undefined) ?? DEFAULT_TOP;
    const headers = {
      Prefer: [
        `odata.maxpagesize=${Math.min(top, spec.pageSizeMax)}`,
        ...(spec.preferences?.(typed) ?? []),
      ].join(", "),
    };
    const requestFor = (query: Record<string, string>): GraphRequest => ({
      pathPattern,
      pathParams,
      query,
      headers,
    });

    let query: Record<string, string> =
      deltaToken !== undefined
        ? { $deltatoken: deltaToken }
        : skipToken !== undefined
          ? { $skiptoken: skipToken }
          : spec.buildInitialQuery(typed);
    const seen = new Set<string>(resume === undefined ? [] : [resume]);
    const items: unknown[] = [];
    let pages = 0;
    let response = await client.request(requestFor(query));
    for (;;) {
      pages += 1;
      const value = (response.data as { value?: unknown } | null | undefined)?.value;
      if (!Array.isArray(value)) {
        // Fail closed: returning (or following) this page's token without its changes would
        // let a caller advance its sync state past changes it never received.
        throw new GraphClientError(
          "graph_error",
          "Graph returned a delta page without a value array.",
        );
      }
      // Whole pages only (see header comment); Graph bounds the page via maxpagesize and the
      // client bounds the response size.
      for (const raw of value) {
        items.push(projectChange(raw, spec.keepKeys));
      }
      const { nextLink, deltaLink } = response;
      if (nextLink !== undefined && deltaLink !== undefined) {
        throw new GraphClientError(
          "graph_error",
          "Graph returned both a nextLink and a deltaLink.",
        );
      }
      const meta: Record<string, unknown> = {
        count: items.length,
        pages,
        clientRequestId: response.clientRequestId,
      };
      if (response.requestId !== undefined) meta["requestId"] = response.requestId;
      if (deltaLink !== undefined) {
        meta["roundComplete"] = true;
        meta["hasMore"] = false;
        meta["deltaToken"] = extractStateToken(deltaLink, "deltatoken", entry, pathParams);
        return successResult(items, meta);
      }
      if (nextLink === undefined) {
        throw new GraphClientError(
          "graph_error",
          "Graph returned neither a nextLink nor a deltaLink for a delta request.",
        );
      }
      const skip = extractStateToken(nextLink, "skiptoken", entry, pathParams);
      // Checked before the token is either returned or followed: handing back the token we
      // were given (or one already seen this call) would make callers replay pages forever.
      if (seen.has(skip)) {
        throw new GraphClientError("graph_error", "Graph returned a repeating skip token.");
      }
      seen.add(skip);
      if (items.length >= top || pages >= MAX_PAGES) {
        meta["roundComplete"] = false;
        meta["hasMore"] = true;
        meta["skipToken"] = skip;
        return successResult(items, meta);
      }
      query = { $skiptoken: skip };
      response = await client.request(requestFor(query));
    }
  };

  return {
    inputSchema: shape,
    handler: async (args) => {
      try {
        return await run(args);
      } catch (err) {
        return failure(toolName, err);
      }
    },
  };
}
