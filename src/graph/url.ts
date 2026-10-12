import { exposedParamName, type EndpointDefinition } from "../catalog.js";
import { GraphClientError } from "./errors.js";
import { pathValueViolation, placeholderViolation } from "./placeholders.js";

/** The only origin the bearer token may ever be sent to. */
export const GRAPH_ORIGIN = "https://graph.microsoft.com";
export type GraphVersion = "v1.0" | "beta";

const MAX_URL_LENGTH = 8192;
const PARAM_RE = /\{([^{}/]+)\}/g;
const QUERY_NAME_RE = /^[$A-Za-z][A-Za-z0-9_.-]*$/;
// Real Graph query names are short, fixed names. Bound diagnostics for malformed
// runtime input before interpolating any accepted name into an error message.
const MAX_QUERY_NAME_LENGTH = 128;
// eslint-disable-next-line no-control-regex
const UNSAFE_RAW_RE = /[\u0000- \u007f\\#]/;

export type QueryValue = string | number | boolean | undefined;
export type QueryParams = Readonly<Record<string, QueryValue>>;

export function endpointVersion(entry: EndpointDefinition): GraphVersion {
  return entry["apiVersion"] === "beta" ? "beta" : "v1.0";
}

export function pathParamNames(pathPattern: string): string[] {
  return [...pathPattern.matchAll(PARAM_RE)].map((m) => m[1] as string);
}

function skipEncodingParams(entry: EndpointDefinition): ReadonlySet<string> {
  const raw = entry["skipEncoding"];
  return new Set(Array.isArray(raw) ? raw.filter((v): v is string => typeof v === "string") : []);
}

function encodeSegment(name: string, value: string): string {
  // Errors name the exposed input (`relative-path`), not the catalog placeholder (`path`).
  const label = exposedParamName(name);
  if (value.length === 0) {
    throw new GraphClientError("invalid_input", `Path parameter ${label} must not be empty.`);
  }
  if (value === "." || value === "..") {
    throw new GraphClientError("invalid_input", `Path parameter ${label} must not be "." or "..".`);
  }
  // "=" stays literal: base64 Exchange ids end in "=", "%3D" has caused Graph 404s
  // (Softeria #245), and "=" is an RFC 3986 sub-delim that cannot alter routing here.
  return encodeURIComponent(value).replace(/%3D/g, "=");
}

/**
 * Substitutes catalog placeholders. Every value becomes exactly one encoded path
 * segment. The only exception is a placeholder the catalog marks `skipEncoding`
 * (a site-relative path): it is split on "/" and each segment is validated and
 * encoded separately, so it still cannot traverse or inject.
 */
export function buildPath(
  entry: EndpointDefinition,
  params: Readonly<Record<string, string>> = {},
): string {
  const expected = pathParamNames(entry.pathPattern);
  for (const key of Object.keys(params)) {
    if (!expected.includes(key)) {
      throw new GraphClientError("invalid_input", "Unexpected path parameter supplied.");
    }
  }
  const multi = skipEncodingParams(entry);
  return entry.pathPattern.replace(PARAM_RE, (_m, name: string) => {
    const value = params[name];
    if (typeof value !== "string") {
      throw new GraphClientError(
        "invalid_input",
        `Missing path parameter ${exposedParamName(name)}.`,
      );
    }
    if (multi.has(name)) {
      // ":" in any decoding would close Graph's path-based addressing early
      // (`root:/a:/other`) and reach a resource outside the catalog.
      const reason = pathValueViolation(value);
      if (reason !== undefined) {
        throw new GraphClientError(
          "invalid_input",
          `Path parameter ${exposedParamName(name)} ${reason}.`,
        );
      }
      return value
        .split("/")
        .map((segment) => encodeSegment(name, segment))
        .join("/");
    }
    const reason = placeholderViolation(name, value, entry.pathPattern);
    if (reason !== undefined) {
      throw new GraphClientError(
        "invalid_input",
        `Path parameter ${exposedParamName(name)} ${reason}.`,
      );
    }
    return encodeSegment(name, value);
  });
}

/** Encodes a query, keeping Graph `$` option names literal. Spaces become %20. */
export function buildQuery(query: QueryParams = {}): string {
  // QueryParams is a TypeScript type, not a runtime trust boundary. Reject invalid
  // records and values before String() can silently coerce arrays/objects into
  // Graph options (for example, $select: ["id", "subject"] -> "id,subject").
  if (
    query === null ||
    typeof query !== "object" ||
    Array.isArray(query) ||
    (Object.getPrototypeOf(query) !== Object.prototype && Object.getPrototypeOf(query) !== null)
  ) {
    throw new GraphClientError("invalid_input", "Query parameters must be a plain record.");
  }
  const parts: string[] = [];
  for (const [name, value] of Object.entries(query)) {
    if (value === undefined) continue;
    if (name.length > MAX_QUERY_NAME_LENGTH || !QUERY_NAME_RE.test(name)) {
      throw new GraphClientError("invalid_input", "Invalid query parameter name.");
    }
    if (typeof value !== "string" && typeof value !== "boolean" && typeof value !== "number") {
      throw new GraphClientError(
        "invalid_input",
        `Query parameter ${name} must be a string, boolean, or finite number.`,
      );
    }
    if (typeof value === "number" && !Number.isFinite(value)) {
      throw new GraphClientError(
        "invalid_input",
        `Query parameter ${name} is not a finite number.`,
      );
    }
    const encodedName = encodeURIComponent(name).replace(/%24/g, "$");
    parts.push(`${encodedName}=${encodeURIComponent(String(value))}`);
  }
  return parts.join("&");
}

export function buildUrl(
  version: GraphVersion,
  path: string,
  query: QueryParams | undefined,
): string {
  const qs = buildQuery(query);
  const url = `${GRAPH_ORIGIN}/${version}${path}${qs.length > 0 ? `?${qs}` : ""}`;
  if (url.length > MAX_URL_LENGTH) {
    throw new GraphClientError("invalid_input", "Request URL is too long.");
  }
  assertFixedOrigin(url);
  return url;
}

/** Final guard run on every URL immediately before it is fetched. */
export function assertFixedOrigin(url: string): void {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    throw new GraphClientError("invalid_input", "Malformed request URL.");
  }
  if (
    parsed.origin !== GRAPH_ORIGIN ||
    parsed.username !== "" ||
    parsed.password !== "" ||
    parsed.hash !== ""
  ) {
    throw new GraphClientError("invalid_input", "Request URL is not on the fixed Graph origin.");
  }
}

/**
 * True when any segment of `text` is `$batch` after (repeated) percent-decoding and
 * case folding. `$batch` is never reachable through the generic request path: it needs
 * the dedicated batch validator and `GraphClient.batch` (SM-4).
 */
export function refersToBatch(text: string): boolean {
  let current = text;
  for (let i = 0; i < 4; i += 1) {
    if (
      current
        .toLowerCase()
        .split(/[/\\?#;]/)
        .some((segment) => segment.trim() === "$batch")
    ) {
      return true;
    }
    let decoded: string;
    try {
      decoded = decodeURIComponent(current);
    } catch {
      return true; // malformed escapes cannot be reasoned about; fail closed
    }
    if (decoded === current) return false;
    current = decoded;
  }
  return true;
}

function escapeRegex(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Regex (matched against the version-less, key-normalized path) for one catalog pattern.
 * Each placeholder is one capture group, in `pathParamNames` order.
 */
export function patternToRegex(entry: EndpointDefinition): RegExp {
  const multi = skipEncodingParams(entry);
  const pattern = entry.pathPattern;
  // `delta()` is returned by Graph as `delta` in nextLink/deltaLink URLs.
  const literal = (text: string): string =>
    escapeRegex(text).replace(/delta\\\(\\\)/g, "delta(?:\\(\\))?");
  let source = "";
  let last = 0;
  for (const m of pattern.matchAll(PARAM_RE)) {
    source += literal(pattern.slice(last, m.index));
    source += multi.has(m[1] as string) ? "(.+)" : "([^/]+)";
    last = m.index + m[0].length;
  }
  source += literal(pattern.slice(last));
  return new RegExp(`^${source}/?$`, "i");
}

/**
 * Graph returns OData key syntax in some links (`/mailFolders('inbox')/messages/delta`).
 * Expand `name('value')` into `name/value` for matching only; the sent path is unchanged.
 */
function normalizeKeySyntax(path: string): string {
  return path.replace(
    /\/([A-Za-z][A-Za-z0-9]*)\('((?:[^'/]|'')+)'\)/g,
    (_m, name: string, value: string) => `/${name}/${value.replace(/''/g, "'")}`,
  );
}

/**
 * Placeholders whose value Graph may legitimately canonicalize inside the SAME principal
 * scope: a well-known folder name (`inbox`) comes back as that folder's opaque id.
 * Every other placeholder (user, group, site, calendar, drive, ...) must be unchanged.
 * Documented paging/delta examples keep `/me/...` as `/me/...`; if a link ever rewrote it
 * to `/users/{id}/...` it would fail the template match and be rejected (fail closed)
 * rather than risk a principal swap.
 */
const CANONICALIZABLE_PARAMS: ReadonlySet<string> = new Set(["mailFolder-id"]);

/**
 * Placeholders that hold Entra object ids (GUIDs) or UPNs, which are case-insensitive.
 * Every other id (calendar, message, drive, site, ...) is opaque and compared exactly.
 */
const CASE_INSENSITIVE_PARAMS: ReadonlySet<string> = new Set(["user-id", "group-id", "team-id"]);

/**
 * Checks every captured placeholder of `match` (a `patternToRegex` result): the value Graph
 * will see (one percent-decoding of the capture) must have an allowed shape, and a
 * skipEncoding value must be colon-free. Undefined when all captures pass.
 */
function capturesViolation(entry: EndpointDefinition, match: RegExpExecArray): string | undefined {
  const multi = skipEncodingParams(entry);
  const names = pathParamNames(entry.pathPattern);
  for (let i = 0; i < names.length; i += 1) {
    const name = names[i] as string;
    const captured = match[i + 1] ?? "";
    const reason = multi.has(name)
      ? pathValueViolation(captured)
      : (() => {
          const actual = decodeCaptured(captured);
          return actual === undefined
            ? "has malformed percent-encoding"
            : placeholderViolation(name, actual, entry.pathPattern);
        })();
    if (reason !== undefined) return `path parameter ${name} ${reason}`;
  }
  return undefined;
}

function decodeCaptured(captured: string): string | undefined {
  try {
    return decodeURIComponent(captured);
  } catch {
    return undefined;
  }
}

function linkError(reason: string): GraphClientError {
  return new GraphClientError("invalid_input", `Rejected unsafe Graph link: ${reason}.`);
}

/**
 * Validates an `@odata.nextLink` / `@odata.deltaLink` and rebuilds it against the fixed
 * origin. The link must be absolute https on exactly graph.microsoft.com (no userinfo,
 * port, or fragment), carry the same API version, and its path must match the SAME
 * catalog endpoint that produced it. Every path placeholder must also keep the value of
 * the original request (after OData key-syntax and percent decoding; exact, except that
 * `CASE_INSENSITIVE_PARAMS` ignore case),
 * so a continuation can never switch user/group/site/etc.; only `CANONICALIZABLE_PARAMS`
 * may differ. The raw string is never fetched.
 */
export function validateGraphLink(
  raw: unknown,
  entry: EndpointDefinition,
  version: GraphVersion,
  pathParams: Readonly<Record<string, string>> = {},
  /** The original request's query; when given, the link may not widen it (`linkWidensQuery`). */
  originalQuery?: QueryParams,
): string {
  if (typeof raw !== "string" || raw.length === 0 || raw.length > MAX_URL_LENGTH) {
    throw linkError("not a usable string");
  }
  if (UNSAFE_RAW_RE.test(raw)) throw linkError("contains forbidden characters");
  if (!/^https:\/\//i.test(raw)) throw linkError("scheme is not https");
  // Inspect the RAW authority: the URL parser drops a default `:443` and a trailing dot.
  const rawAuthority = /^https:\/\/([^/?#\\]*)/i.exec(raw)?.[1] ?? "";
  if (rawAuthority.toLowerCase() !== "graph.microsoft.com") {
    throw linkError("authority is not exactly graph.microsoft.com");
  }
  let parsed: URL;
  try {
    parsed = new URL(raw);
  } catch {
    throw linkError("not an absolute URL");
  }
  if (parsed.protocol !== "https:") throw linkError("scheme is not https");
  if (parsed.username !== "" || parsed.password !== "") throw linkError("contains userinfo");
  if (parsed.port !== "") throw linkError("contains a port");
  if (parsed.hostname !== "graph.microsoft.com" || parsed.origin !== GRAPH_ORIGIN) {
    throw linkError("origin is not the fixed Graph origin");
  }
  // Inspect the RAW path: the URL parser would silently collapse encoded dot segments.
  const pathStart = raw.indexOf("/", "https://".length);
  const rawPath = pathStart === -1 ? "" : (raw.slice(pathStart).split("?")[0] ?? "");
  if (/%(?:2f|5c|2e|00)/i.test(rawPath)) throw linkError("encoded path trickery");

  const prefix = `/${version}`;
  if (!parsed.pathname.startsWith(`${prefix}/`)) throw linkError("API version mismatch");
  const rest = normalizeKeySyntax(parsed.pathname.slice(prefix.length));
  const match = patternToRegex(entry).exec(rest);
  if (match === null) throw linkError("path does not match the original endpoint");
  // Applies to every placeholder, including the (binding-exempt) mailFolder-id.
  const violation = capturesViolation(entry, match);
  if (violation !== undefined) throw linkError(violation);
  pathParamNames(entry.pathPattern).forEach((name, i) => {
    if (CANONICALIZABLE_PARAMS.has(name)) return;
    const expected = pathParams[name];
    const actual = decodeCaptured(match[i + 1] ?? "");
    const fold = (v: string): string => (CASE_INSENSITIVE_PARAMS.has(name) ? v.toLowerCase() : v);
    if (typeof expected !== "string" || actual === undefined || fold(actual) !== fold(expected)) {
      throw linkError(`path parameter ${name} differs from the original request`);
    }
  });

  if (originalQuery !== undefined) {
    // The continuation does not rebuild the original URL, but it must enforce
    // the exact same runtime query contract before linkWidensQuery() examines
    // values; otherwise String(array/object) silently widens the input boundary.
    buildQuery(originalQuery);
    const reason = linkWidensQuery(parsed.searchParams, originalQuery);
    if (reason !== undefined) throw linkError(reason);
  }

  // The link is otherwise followed as Graph returned it (Microsoft: use the whole nextLink).
  const url = `${GRAPH_ORIGIN}${parsed.pathname}${parsed.search}`;
  assertFixedOrigin(url);
  return url;
}

// OData system options that change what a page contains. A continuation may carry paging state
// ($skiptoken, $skip, $deltatoken, $top) and endpoint parameters freely, but never one of these
// unless the original request sent it (Graph accepts each with or without the "$").
const SHAPE_OPTIONS = new Set([
  "expand",
  "select",
  "filter",
  "orderby",
  "search",
  "count",
  "apply",
  "compute",
  "format",
]);

const optionName = (key: string): string => key.toLowerCase().replace(/^\$/, "");

/** Top-level comma-separated items of an `$expand`/`$select` value (commas inside "(...)" kept). */
function topLevelItems(value: string): Set<string> {
  const items = new Set<string>();
  let depth = 0;
  let start = 0;
  for (let i = 0; i <= value.length; i += 1) {
    const ch = value[i];
    if (ch === "(") depth += 1;
    else if (ch === ")") depth -= 1;
    else if ((ch === "," && depth === 0) || i === value.length) {
      const item = value.slice(start, i).trim().toLowerCase();
      if (item !== "") items.add(item);
      start = i + 1;
    }
  }
  return items;
}

/**
 * Why a continuation's query would widen the original request (a shape option the original did
 * not send, or `$expand`/`$select` items beyond the original's), or undefined. Values of options
 * present in both are otherwise not compared: Graph re-encodes them in its links.
 */
function linkWidensQuery(link: URLSearchParams, originalQuery: QueryParams): string | undefined {
  const original = new Map<string, string[]>();
  for (const [key, value] of Object.entries(originalQuery)) {
    if (value === undefined) continue;
    const name = optionName(key);
    original.set(name, [...(original.get(name) ?? []), String(value)]);
  }
  for (const [key, value] of link) {
    const name = optionName(key);
    if (!SHAPE_OPTIONS.has(name)) continue;
    const sent = original.get(name);
    if (sent === undefined)
      return `continuation adds query option ${name} not in the original request`;
    if (name === "expand" || name === "select") {
      const allowed = new Set(sent.flatMap((v) => [...topLevelItems(v)]));
      for (const item of topLevelItems(value)) {
        if (!allowed.has(item)) return `continuation widens ${name} beyond the original request`;
      }
    }
  }
  return undefined;
}

// RFC 3986 pchar plus "/" and "%": the only characters a raw relative path may contain.
const RELATIVE_PATH_RE = /^[A-Za-z0-9\-._~!$&'()*+,;=:@%/]*$/;
const MAX_QUERY_PAIRS = 50;

export interface RelativeGraphUrl {
  /** The catalog entry the path matched. */
  entry: EndpointDefinition;
  /** Canonical relative form (`/path?query`), safe to hand to Graph relative to `version`. */
  url: string;
}

function relativeError(reason: string): GraphClientError {
  return new GraphClientError("invalid_input", `Rejected unsafe relative URL: ${reason}.`);
}

/**
 * Validates a caller-supplied relative Graph URL (a `$batch` subrequest) and returns its
 * canonical form. The URL must start with a single "/" (the only form this server accepts;
 * Graph also tolerates a missing slash, which we refuse to keep the grammar tight), be
 * relative to the Graph version root (no scheme, authority, backslash, version prefix,
 * dot segment, encoded / \ . NUL, or `$batch`), and its path must match one of `permitted`
 * whose Graph version equals `version`. The query is parsed into pairs and re-encoded with
 * the same rules as `buildQuery`. Nothing is fetched here.
 */
export function validateRelativeGraphUrl(
  raw: unknown,
  permitted: readonly EndpointDefinition[],
  version: GraphVersion,
): RelativeGraphUrl {
  if (typeof raw !== "string" || raw.length === 0 || raw.length > MAX_URL_LENGTH) {
    throw relativeError("not a usable string");
  }
  if (!raw.startsWith("/")) throw relativeError("must start with a single '/'");
  if (raw.startsWith("//")) throw relativeError("network-path reference");
  if (raw.includes("\\")) throw relativeError("contains a backslash");
  if (raw.includes("#")) throw relativeError("contains a fragment");
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001f\u007f]/.test(raw)) throw relativeError("contains control characters");
  // Check escape syntax first so a malformed escape is reported as such, not as "$batch".
  if (/%(?![0-9A-Fa-f]{2})/.test(raw)) throw relativeError("malformed percent-encoding");
  if (refersToBatch(raw)) throw relativeError("nested $batch is not allowed");

  const queryAt = raw.indexOf("?");
  const rawPath = queryAt === -1 ? raw : raw.slice(0, queryAt);
  const rawQuery = queryAt === -1 ? undefined : raw.slice(queryAt + 1);

  if (!RELATIVE_PATH_RE.test(rawPath)) throw relativeError("path contains forbidden characters");
  if (/%(?![0-9A-Fa-f]{2})/.test(rawPath)) throw relativeError("malformed percent-encoding");
  if (/%(?:2f|5c|2e|00)/i.test(rawPath) || /%25(?:2f|5c|2e|00|25)/i.test(rawPath)) {
    throw relativeError("encoded path trickery");
  }
  const segments = rawPath.slice(1).split("/");
  if (segments.some((s) => s === "." || s === "..")) throw relativeError("dot segment");
  // The final segment may be empty (trailing slash); any other empty segment is "//".
  if (segments.slice(0, -1).some((s) => s === "")) throw relativeError("empty path segment");
  if (/^(?:v1\.0|beta)$/i.test(segments[0] ?? "")) {
    throw relativeError("must not include an API version prefix");
  }

  const rest = normalizeKeySyntax(rawPath);
  // A candidate qualifies only if its pattern matches AND every captured placeholder passes
  // its allowlisted shape (so `/groups/delta` cannot ride `/groups/{group-id}`). Among
  // qualifying candidates the most specific (fewest placeholders) wins, so
  // `/me/events/delta()` resolves to the delta entry rather than `/me/events/{event-id}`.
  let rejection: string | undefined;
  const qualified: EndpointDefinition[] = [];
  for (const e of permitted) {
    if (endpointVersion(e) !== version) continue;
    const m = patternToRegex(e).exec(rest);
    if (m === null) continue;
    const violation = capturesViolation(e, m);
    if (violation === undefined) qualified.push(e);
    else rejection ??= violation;
  }
  const entry = qualified.sort(
    (x, y) => pathParamNames(x.pathPattern).length - pathParamNames(y.pathPattern).length,
  )[0];
  if (entry === undefined) {
    throw relativeError(rejection ?? "path is not in the permitted endpoint catalog");
  }

  let query = "";
  if (rawQuery !== undefined && rawQuery.length > 0) {
    const pairs = rawQuery.split("&");
    if (pairs.length > MAX_QUERY_PAIRS) throw relativeError("too many query parameters");
    query = pairs
      .map((pair) => {
        const eq = pair.indexOf("=");
        if (eq <= 0) throw relativeError("malformed query parameter");
        const name = pair.slice(0, eq);
        const rawValue = pair.slice(eq + 1);
        // A bare "+" is ambiguous between a space and a plus; require %20 / %2B instead.
        if (rawValue.includes("+")) throw relativeError("unencoded '+' in query value");
        let value: string;
        try {
          value = decodeURIComponent(rawValue);
        } catch {
          throw relativeError("malformed percent-encoding in query");
        }
        return buildQuery({ [name]: value });
      })
      .join("&");
  }
  const url = query.length > 0 ? `${rawPath}?${query}` : rawPath;
  if (url.length > MAX_URL_LENGTH) throw relativeError("URL is too long");
  return { entry, url };
}
