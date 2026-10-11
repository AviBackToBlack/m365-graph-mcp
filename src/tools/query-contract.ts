import { z } from "zod";
import type { BatchPolicy } from "../graph/batch.js";
import { GraphClientError } from "../graph/errors.js";
import type { QueryParams } from "../graph/url.js";
import type { ToolArgs } from "../registry.js";
import {
  EXTENDED_PROPERTIES_KEY,
  MAX_EXTENDED_PROPERTY_ID_LENGTH,
  extendedPropertyExpand,
  isExtendedPropertyId,
  parseExtendedPropertyExpand,
  projectExtendedProperties,
} from "./extended-properties.js";
import { canonicalInstant, instantTicks } from "./timestamps.js";
import { canonicalTimeZone } from "./timezones.js";

/**
 * Per-tool query contract: ONE definition that drives both the tool's typed MCP inputs and
 * the `graph-batch` query policy (`SECURITY_INVARIANTS.md`: "Permit only documented query
 * parameters per tool", "Bound `$top`", "Treat `$search`/`$filter`/KQL as data").
 *
 * Callers never supply raw OData strings named after a URL: they supply typed options
 * (`top`, `select`, `filter`, `search`, `orderby`, `skip`, ...) and this module is the only
 * place that turns them into wire query parameters. Batch subrequests carry wire pairs;
 * `batchPolicy` parses those back into the same typed options and runs the same validation,
 * so a subrequest is accepted if and only if the single-request tool would accept it.
 *
 * Wire rules (exact, case-sensitive; anything else is rejected, never normalized):
 * - names: `$top $select $filter $search $orderby $skip includeHiddenFolders`, once each. The
 *   `$`-less spellings Graph tolerates, other casings, `$skiptoken`, `$count`, `$format` and
 *   every other name are not part of any contract. `$expand` exists only for contracts that
 *   declare `extendedProperty` (below); there it is never free text.
 * - `$top`/`$skip`: canonical decimal digits (no sign, leading zero, exponent, fraction).
 * - `$select`: comma-separated, each name in the tool's allowlist.
 * - `$orderby`: comma-separated `property[ asc|desc]`, property in the tool's allowlist. With a
 *   `$filter`, Graph requires every `$orderby` property to appear in the filter, in the same
 *   order, before other filter properties; the contract checks the first two conditions
 *   locally (`orderbyFilterViolation`) and Graph still enforces the remainder.
 * - `startDateTime`/`endDateTime` (calendarView and instances; `window`): both REQUIRED, once each,
 *   real instants, start strictly before end (compared exactly, 100 ns), and a window of at most
 *   `window.maxDays` days (a project safety limit; Graph documents none). The single-request tool
 *   accepts ISO-8601 with `Z` or an explicit `+/-HH:MM` offset (within +/-14:00; offset-less local
 *   times are rejected) and ALWAYS sends the canonical UTC `Z` form (`canonicalInstant`: fraction
 *   digits trimmed of trailing zeros). A batch subrequest carries the wire value verbatim, so it
 *   must already BE that canonical form: offsets (even `%2B`-encoded) and non-canonical spellings
 *   such as `.0000000Z` are rejected, which keeps batch bytes identical to the single path.
 * - `select` is optional per contract (none offered/sent, e.g. mailboxSettings, profile); a default, when
 *   present, is never empty. `$expand` (`expand`; typed input `expand`): a comma list of names from the
 *   tool's allowlist, no nested options (profile only); never combined with `extendedProperty`.
 * - `$expand` (`extendedProperty`; typed input `extendedPropertyId`): exactly
 *   `singleValueExtendedProperties($filter=id eq '<id>')` for one documented single-value extended
 *   property id (see `extended-properties.ts`), built by the server; batch accepts only that form.
 * - `Prefer` (never a query name): the single request path emits ONE header, preferences in
 *   this order separated by ", " (RFC 7240 list form): `outlook.timezone="<name>"` when the
 *   tool declares `timeZone` and the caller chose a zone, then `outlook.body-content-type="<t>"`
 *   when it declares `bodyType`. The batch policy accepts exactly those forms (zone from the
 *   allowlist in canonical spelling) and nothing else.
 * - `$search`: KQL wrapped in ONE pair of double quotes with `\` and `"` backslash-escaped
 *   (`kqlQuote`). The caller supplies the bare KQL text; it is data, never evaluated here.
 *   Exception, `search.siteSearch` (`GET /sites?search=`): the documented bare name `search`, required,
 *   text sent verbatim (percent-encoded), never `$search` (see `QueryContractSpec.search.siteSearch`).
 */

const WIRE_NAMES = {
  top: "$top",
  select: "$select",
  filter: "$filter",
  search: "$search",
  orderby: "$orderby",
  skip: "$skip",
  includeHiddenFolders: "includeHiddenFolders",
  startDateTime: "startDateTime",
  endDateTime: "endDateTime",
  extendedPropertyId: "$expand",
  expand: "$expand",
} as const;
type OptionKey = keyof typeof WIRE_NAMES;

export interface QueryContractSpec {
  /**
   * Collection tools only. `max` is the project hard limit on returned items. `pageMax`
   * (default `max`) caps the `$top` sent to Graph per request: a larger `max` is reached by
   * following pages in the single-request path, while batch cannot paginate and so is bound
   * to `pageMax`.
   */
  top?: {
    default: number;
    max: number;
    pageMax?: number;
    /**
     * Default true. False for collections whose Graph page documents no `$top` (directReports):
     * `top` is then only the local item limit (pages are followed up to it), no `$top` is ever
     * sent, and a batch subrequest carrying `$top` is rejected.
     */
    wire?: boolean;
    /**
     * Offset-paged collections (`/me/people`: `$top`/`$skip`, no reliable nextLink). A page that
     * fills the limit means more MAY exist, so batch bodies report `hasMore` on a full page and
     * carry `nextSkip` (see `nextSkipFor`), matching the single-request tool.
     */
    offsetPaging?: boolean;
  };
  /**
   * Optional. When omitted the tool offers no `select` input, never sends `$select`, and a
   * batch subrequest carrying one is rejected (e.g. `/me/mailboxSettings`, `/me/profile`).
   */
  select?: {
    allowed: readonly [string, ...string[]];
    /**
     * Applied when the caller sends no `select` (compact defaults). Must be non-empty when
     * set (an empty list would put `$select=` on the wire); omit it to send NO `$select`
     * unless the caller chose one (Graph's own default property set applies).
     */
    default?: readonly [string, ...string[]];
    /**
     * Batch cannot inject defaults (a policy only accepts or rejects), so tools whose
     * unselected response is not compact/safe require an explicit allowlisted `$select`.
     */
    requiredInBatch: boolean;
  };
  /**
   * Single-entity tools with a navigation-property `$expand` (`/me/profile`): optional typed
   * `expand`, a list of names from `allowed`, sent as `$expand=a,b` (no nested options, no
   * free text). Mutually exclusive with `extendedProperty`, which owns `$expand` otherwise.
   */
  expand?: { allowed: readonly [string, ...string[]] };
  /**
   * Single-entity tools: project the response to `id`, the select/expand/`keep` names
   * and the kept annotations, dropping `@odata.context` and anything unexpected. A response
   * that is not a JSON object fails closed (`projectsEntity`, see `getTool`).
   * On a collection contract (used via `listTool`), `keep` is still folded into the item and
   * batch keep-set, so it can widen that set for a later sanitizer
   * (`list-trending-insights`); `projectsEntity` is then set but unused.
   */
  projectEntity?: { keep?: readonly string[] };
  filter?: { maxLength: number };
  search?: {
    maxLength: number;
    /**
     * Directory-object search (`/users`): the typed text must be ONE `property:text` clause with
     * `property` from this list (text free of quotes, backslashes, colons, parentheses, ampersands and the operators AND/OR/NOT). The
     * server wraps it as `$search="property:text"` and, because Graph requires advanced query
     * mode for directory `$search`, ITSELF adds `$count=true` and `ConsistencyLevel: eventual`.
     * Callers never supply either. graph-batch rejects `$search` for such a contract (it cannot
     * reproduce the header + `$count` pair), so batch and single stay identical by exclusion.
     */
    directoryProperties?: readonly [string, ...string[]];
    /**
     * People search (`/me/people`): plain fuzzy name/e-mail text only, restricted to letters,
     * digits, space and `. @ ' _ -` (PLAIN_NAME_RE). That excludes quotes, backslashes, colons
     * (the documented `topic:` keyword) and every other operator character, so the wire form
     * is exactly the documented `$search="tyler lee"` with nothing to escape. No header or
     * `$count` is involved, so graph-batch accepts it (the same refine runs on the batch path).
     */
    plainName?: boolean;
    /**
     * SharePoint site search (`GET /sites?search={query}`): the page documents the bare name
     * `search` (no `$`) with free-text keywords, no KQL quoting. The typed text is REQUIRED (a
     * bare `/sites` must never become an enumeration; a text without a letter or digit, such as
     * `*`, is rejected for the same reason) and is sent verbatim, percent-encoded, as
     * `search=<text>`. `$search` is not part of this contract (single path never emits it; a batch
     * subrequest carrying it is rejected). Incompatible with `directoryProperties`/`plainName`.
     */
    siteSearch?: boolean;
    /** Replaces the mail-oriented default description. */
    description?: string;
  };
  orderby?: {
    allowed: readonly [string, ...string[]];
    maxClauses: number;
    /**
     * Mail rule (user-list-messages): with a `$filter`, every `$orderby` property must also
     * appear in it. Default true; calendar tools document no such rule and set false.
     */
    requireInFilter?: boolean;
    /**
     * Teams chats (`/me/chats`): Graph supports only DESCENDING order. Each clause then needs an
     * explicit `desc` (the literal), so a bare property (Graph's ascending default), `asc` and
     * every other direction are rejected by the one `validate()` both paths share.
     */
    descOnly?: boolean;
  };
  skip?: { max: number };
  includeHiddenFolders?: boolean;
  /** Single-item tools: `Prefer: outlook.body-content-type` (a header, not a query option). */
  bodyType?: { default: "text" | "html" };
  /**
   * Calendar tools: optional typed `timeZone` sent as `Prefer: outlook.timezone="<name>"`
   * (response start/end only; it never affects the request window). Absent means Graph's
   * default (UTC).
   */
  timeZone?: boolean;
  /**
   * calendarView / instances: required `startDateTime` and `endDateTime` (`Z` or `+/-HH:MM`
   * offset, sent as canonical UTC `Z`), start < end, and a span of at most `maxDays` days (a
   * project safety limit; Graph documents none).
   */
  window?: { maxDays: number };
  /**
   * Event tools: optional typed `extendedPropertyId` sent as the documented filtered
   * `$expand=singleValueExtendedProperties($filter=id eq '<id>')`. Needs Calendars.Read.
   */
  extendedProperty?: boolean;
}

export interface QueryRequest {
  query: QueryParams;
  headers: Record<string, string>;
  /** Item ceiling for collection tools (requested or default `top`); 0 when not applicable. */
  limit: number;
}

export interface QueryContract {
  /** Typed MCP input shape (query options only; path params are derived from the catalog). */
  readonly shape: z.ZodRawShape;
  /** Validates typed tool args and builds the wire query/headers. Throws `invalid_input`. */
  toRequest(args: ToolArgs): QueryRequest;
  /**
   * The tool's batch policy: the same validation applied to decoded batch query pairs and to
   * the subrequest headers (only the exact `Prefer` the single-request path emits, if any).
   */
  readonly batchPolicy: BatchPolicy;
  /** True when `projectItem` is a strict entity projection (`projectEntity`): non-objects yield undefined. */
  readonly projectsEntity: boolean;
  /**
   * Item projector for collection tools: keeps only the `select` allowlist plus `id` and the
   * `@odata.type`/`@odata.etag` annotations; non-objects yield undefined. Defense in depth
   * for responses that ignored (or whose nextLink dropped) `$select`. Contracts with
   * `extendedProperty` also keep the expanded `singleValueExtendedProperties` as id/value
   * pairs only (this is also applied to single-entity responses). Identity for others.
   */
  projectItem(item: unknown): unknown;
  /**
   * Single-entity response projector for one request, given that request's wire query (the
   * `query` from `toRequest`). Strict entity contracts (`projectEntity`) keep only `id`, the
   * request's effective `$select` (the contract default when none was sent) and `$expand`
   * names, the contract's `keep` names and the allowed annotations; a non-object yields
   * undefined. Other contracts behave like `projectItem` for an entity.
   */
  projectResponse(item: unknown, query: Readonly<Record<string, unknown>>): unknown;
  /**
   * Collection item projector for one request, given its wire query (`toRequest().query`):
   * keeps `id`, the request's effective `$select` and allowed annotations (see
   * `projectItem` for the extended-property handling). Used by `listTool`.
   */
  itemProjector(query: Readonly<Record<string, unknown>>): (item: unknown) => unknown;
}

/** Annotation keys kept on projected items (never `@odata.nextLink`, `@odata.context`, ...). */
const ITEM_ANNOTATIONS: ReadonlySet<string> = new Set(["@odata.type", "@odata.etag"]);

/** The canonical `Prefer` preference tokens the single-request path emits. */
const bodyPrefer = (type: string): string => `outlook.body-content-type="${type}"`;
const timeZonePrefer = (name: string): string => `outlook.timezone="${name}"`;
const TZ_PREFER_RE = /^outlook\.timezone="([^"]{1,64})"$/;
const BODY_PREFER_RE = /^outlook\.body-content-type="(text|html)"$/;
const DAY_TICKS = 864_000_000_000n; // 24 h in 100 ns ticks

function invalid(message: string): GraphClientError {
  return new GraphClientError("invalid_input", message);
}

/** Typed `startDateTime`/`endDateTime` input (shared by the query contracts and the calendarView delta tool). */
export function windowStampSchema(what: string, maxDays: number, required = true) {
  return z
    .string()
    .refine((v) => instantTicks(v) !== undefined)
    .describe(
      `${what} as an ISO-8601 timestamp with a Z or an explicit UTC offset, e.g. 2024-01-31T00:00:00Z or 2024-01-31T09:00:00+01:00 (offset within +/-14:00; local times without Z/offset are rejected). It is sent to Graph as the equivalent UTC value. ${required ? "Required; s" : "S"}tart must be before end and, as a project safety limit (not a Graph limit), the window may span at most ${maxDays} days.`,
    );
}

/** Exact (sub-millisecond) start < end and span checks on two already-validated instants. */
export function assertWindow(startValue: string, endValue: string, maxDays: number): void {
  const start = instantTicks(startValue) as bigint;
  const end = instantTicks(endValue) as bigint;
  if (start >= end) throw invalid("startDateTime must be before endDateTime.");
  if (end - start > BigInt(maxDays) * DAY_TICKS) {
    throw invalid(`The time window may span at most ${maxDays} days.`);
  }
}

/** Typed `timeZone` input (allowlisted name; sent only as `Prefer: outlook.timezone`). */
export function timeZoneSchema() {
  return z
    .string()
    .min(1)
    .max(64)
    .refine((v) => canonicalTimeZone(v) !== undefined)
    .optional()
    .describe(
      'Time zone for the start/end of returned events: a Microsoft-documented IANA name (e.g. "Europe/Berlin"), a Windows name (e.g. "Pacific Standard Time") or "UTC". Default UTC. Does not change the requested window, which is always UTC.',
    );
}

/** The canonical `Prefer` preference for an allowlisted zone name (validated by the caller). */
export function timeZonePreferFor(name: string): string {
  return timeZonePrefer(canonicalTimeZone(name) as string);
}

// eslint-disable-next-line no-control-regex
const CONTROL_RE = /[\u0000-\u001f\u007f]/;
const noControl = (value: string): boolean => !CONTROL_RE.test(value);

/** Wraps bare KQL in one pair of double quotes, escaping `\` and `"` with a backslash. */
export function kqlQuote(text: string): string {
  return `"${text.replace(/[\\"]/g, "\\$&")}"`;
}

/** Inverse of `kqlQuote`; undefined when `wire` is not exactly one well-formed quoted run. */
export function kqlUnquote(wire: string): string | undefined {
  if (wire.length < 2 || !wire.startsWith('"') || !wire.endsWith('"')) return undefined;
  const inner = wire.slice(1, -1);
  let out = "";
  for (let i = 0; i < inner.length; i += 1) {
    const c = inner[i] as string;
    if (c === '"') return undefined; // an unescaped quote would close the run early
    if (c === "\\") {
      const next = inner[i + 1];
      if (next !== "\\" && next !== '"') return undefined; // dangling or unknown escape
      out += next;
      i += 1;
    } else {
      out += c;
    }
  }
  return out;
}

interface Orderby {
  property: string;
  direction?: "asc" | "desc";
}

function orderbyWire(clauses: readonly Orderby[]): string {
  return clauses
    .map((c) => (c.direction === undefined ? c.property : `${c.property} ${c.direction}`))
    .join(",");
}

function parseOrderbyWire(value: string): Orderby[] {
  return value.split(",").map((part) => {
    const m = /^([^\s,]+)(?: (asc|desc))?$/.exec(part);
    if (m === null) throw invalid("Malformed $orderby.");
    return m[2] === undefined
      ? { property: m[1] as string }
      : { property: m[1] as string, direction: m[2] as "asc" | "desc" };
  });
}

/** One `property:text` clause (directory `$search`); see `QueryContractSpec.search`. */
function directorySearchOk(value: string, properties: readonly string[]): boolean {
  const m = /^([A-Za-z]+):(.+)$/s.exec(value);
  if (m === null || !properties.includes(m[1] as string)) return false;
  const text = m[2] as string;
  // `&` is excluded too: Graph documents encoded ampersands in directory $search values as
  // failing, and every query value is percent-encoded on the wire.
  // One clause only: the KQL operators (case-sensitive, uppercase) may not appear as words,
  // so the text cannot add free-text terms outside the named property.
  if (/(^|\s)(AND|OR|NOT)(\s|$)/.test(text)) return false;
  return /^[^"\\:()&]+$/.test(text) && /[\p{L}\p{N}]/u.test(text);
}

/**
 * Next offset for an offset-paged collection (`top.offsetPaging`), or undefined when it would
 * exceed the contract's `skip` maximum (the caller could not pass it back).
 */
export function nextSkipFor(skip: number, returned: number, skipMax: number): number | undefined {
  const next = skip + returned;
  return next <= skipMax ? next : undefined;
}

/** Characters allowed in a people `$search` (see `QueryContractSpec.search.plainName`). */
const PLAIN_NAME_RE = /^[\p{L}\p{N} .@'_-]+$/u;

function escapeRe(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\/]/g, "\\$&");
}

/**
 * Conservative local check of Graph's `$filter` + `$orderby` rule (user-list-messages): every
 * orderby property must occur in the filter as a whole property token and the first
 * occurrences must follow orderby order. Single-quoted literals are ignored. This is not an
 * OData parser; Graph still rejects the remainder (e.g. filter-only properties placed before
 * orderby properties). The message never echoes caller values.
 */
function orderbyFilterViolation(filter: string, orderby: readonly Orderby[]): string | undefined {
  const text = filter.replace(/'(?:[^']|'')*'/g, " ");
  let last = -1;
  for (const { property } of orderby) {
    const re = new RegExp(`(?<![\\w/.])${escapeRe(property)}(?![\\w/])`);
    const at = re.exec(text)?.index ?? -1;
    if (at <= last) {
      return "When filter is combined with orderby, every orderby property must also appear in the filter, in the same order (Microsoft Graph rule).";
    }
    last = at;
  }
  return undefined;
}

const DIGITS_RE = /^(?:0|[1-9][0-9]{0,8})$/;

function strictInt(value: string, label: string): number {
  if (!DIGITS_RE.test(value)) throw invalid(`${label} must be a plain non-negative integer.`);
  return Number(value);
}

export function defineQueryContract(spec: QueryContractSpec): QueryContract {
  const shape: Record<string, z.ZodType> = {};
  if (spec.top !== undefined) {
    shape["top"] = z
      .number()
      .int()
      .min(1)
      .max(spec.top.max)
      .optional()
      .describe(
        `Maximum items to return (1-${spec.top.max}; default ${spec.top.default}).${
          spec.top.pageMax !== undefined && spec.top.pageMax < spec.top.max
            ? ` Fetched in pages of at most ${spec.top.pageMax}.`
            : ""
        }`,
      );
  }
  if (spec.select !== undefined) {
    shape["select"] = z
      .array(z.enum(spec.select.allowed))
      .min(1)
      .max(spec.select.allowed.length)
      .optional()
      .describe(
        `Properties to return. Default: ${
          spec.select.default?.join(", ") ?? "the properties Microsoft Graph returns by default"
        }.`,
      );
  }
  if (spec.filter !== undefined) {
    shape["filter"] = z
      .string()
      .min(1)
      .max(spec.filter.maxLength)
      .refine(noControl)
      .optional()
      .describe(
        `OData $filter expression, e.g. isRead eq false (max ${spec.filter.maxLength} chars). Passed to Microsoft Graph as data. Cannot be combined with search.`,
      );
  }
  if (spec.search !== undefined) {
    const directoryProps = spec.search.directoryProperties;
    const searchText: z.ZodType<string> = z
      .string()
      .min(1)
      .max(spec.search.maxLength)
      .refine(noControl)
      .refine((v) => /[\p{L}\p{N}]/u.test(v))
      .refine((v) => directoryProps === undefined || directorySearchOk(v, directoryProps))
      .refine((v) => spec.search?.plainName !== true || PLAIN_NAME_RE.test(v));
    shape["search"] = (
      spec.search.siteSearch === true ? searchText : searchText.optional()
    ).describe(
      spec.search.description ??
        `KQL search text WITHOUT an enclosing quote pair, e.g. from:john AND subject:budget (max ${spec.search.maxLength} chars). The server wraps it in double quotes and backslash-escapes any " or \\ you include. Returns at most 1000 results ordered by sent date; cannot be combined with filter, orderby or skip.`,
    );
  }
  if (spec.orderby !== undefined) {
    const descOnly = spec.orderby.descOnly === true;
    shape["orderby"] = z
      .array(
        z.strictObject({
          property: z.enum(spec.orderby.allowed),
          direction: descOnly ? z.literal("desc") : z.enum(["asc", "desc"]).optional(),
        }),
      )
      .min(1)
      .max(spec.orderby.maxClauses)
      .optional()
      .describe(
        descOnly
          ? 'Sort order. Microsoft Graph supports only descending order here: every clause must carry direction "desc".'
          : spec.orderby.requireInFilter === false
            ? "Sort order."
            : "Sort order. When combined with filter, Microsoft Graph requires every orderby property to also appear in the filter (same order, first).",
      );
  }
  if (spec.skip !== undefined) {
    shape["skip"] = z
      .number()
      .int()
      .min(0)
      .max(spec.skip.max)
      .optional()
      .describe(`Items to skip (0-${spec.skip.max}). Not supported together with search.`);
  }
  if (spec.includeHiddenFolders === true) {
    shape["includeHiddenFolders"] = z.boolean().optional().describe("Include hidden folders.");
  }
  if (spec.window !== undefined) {
    shape["startDateTime"] = windowStampSchema("Start of the time range", spec.window.maxDays);
    shape["endDateTime"] = windowStampSchema("End of the time range", spec.window.maxDays);
  }
  if (spec.timeZone === true) {
    shape["timeZone"] = timeZoneSchema();
  }
  if (spec.extendedProperty === true) {
    shape["extendedPropertyId"] = z
      .string()
      .min(1)
      .max(MAX_EXTENDED_PROPERTY_ID_LENGTH)
      .refine(isExtendedPropertyId)
      .optional()
      .describe(
        'Also return one single-value extended property of each event, by its documented id: "{type} {guid} Name {name}" (e.g. "String {8ECCC264-6880-4EBE-992F-8888D2EEAA1D} Name TestProperty"), "{type} {guid} Id 0x8012" or "{type} 0x4001". Returned as singleValueExtendedProperties [{id, value}] (absent when the event has no such property). Requires the Calendars.Read permission (Calendars.ReadBasic does not cover extended properties).',
      );
  }
  // Strict entity projection derives its keep set from select/expand only, so it would drop a
  // requested extended property on the single path while batch kept it: refuse the combination.
  // A directory search needs a `property:` clause, which the plain-name charset forbids, so the
  // combination would silently reject every search value.
  if (spec.search?.plainName === true && spec.search.directoryProperties !== undefined) {
    throw new Error(
      "A query contract cannot declare both search.plainName and search.directoryProperties.",
    );
  }
  if (
    spec.search?.siteSearch === true &&
    (spec.search.plainName === true || spec.search.directoryProperties !== undefined)
  ) {
    throw new Error(
      "A query contract cannot combine search.siteSearch with search.plainName or search.directoryProperties.",
    );
  }
  if (spec.projectEntity !== undefined && spec.extendedProperty === true) {
    throw new Error("A query contract cannot declare both projectEntity and extendedProperty.");
  }
  if (spec.expand !== undefined) {
    if (spec.extendedProperty === true) {
      throw new Error("A query contract cannot declare both expand and extendedProperty.");
    }
    shape["expand"] = z
      .array(z.enum(spec.expand.allowed))
      .min(1)
      .max(spec.expand.allowed.length)
      .optional()
      .describe(
        `Related entities or collections to include inline (names only): ${spec.expand.allowed.join(", ")}. Default: none.`,
      );
  }
  if (spec.bodyType !== undefined) {
    shape["bodyType"] = z
      .enum(["text", "html"])
      .optional()
      .describe(`Body format (default ${spec.bodyType.default}).`);
  }

  const pageMax = spec.top === undefined ? 0 : (spec.top.pageMax ?? spec.top.max);
  const keepKeys: ReadonlySet<string> = new Set([
    "id",
    ...(spec.select?.allowed ?? []),
    ...(spec.expand?.allowed ?? []),
    ...(spec.projectEntity?.keep ?? []),
    ...(spec.extendedProperty === true ? [EXTENDED_PROPERTIES_KEY] : []),
  ]);
  const projectItemWith =
    (keys: ReadonlySet<string>) =>
    (item: unknown): unknown => {
      if (typeof item !== "object" || item === null || Array.isArray(item)) return undefined;
      const kept = Object.fromEntries(
        Object.entries(item).filter(([k]) => keys.has(k) || ITEM_ANNOTATIONS.has(k)),
      );
      return spec.extendedProperty === true ? projectExtendedProperties(kept) : kept;
    };
  const projectItem = projectItemWith(keepKeys);
  /**
   * Collection item keep set for ONE request: `id`, that request's effective `$select` (the
   * wire value, or the contract default when none was sent) intersected with the allowlist,
   * and the expanded extended-property key where the contract has one. Over-returned
   * allowlisted fields the caller did not select are dropped (single path and batch alike).
   * A contract without `select` keeps its allowlist-wide set.
   */
  const collectionKeys = (selectWire: unknown): ReadonlySet<string> => {
    if (spec.select === undefined) return keepKeys;
    const allowed = new Set<string>(spec.select.allowed);
    const selected =
      typeof selectWire === "string" && selectWire.length > 0
        ? selectWire.split(",")
        : [...(spec.select.default ?? spec.select.allowed)];
    return new Set([
      "id",
      ...selected.filter((n) => allowed.has(n)),
      ...(spec.extendedProperty === true ? [EXTENDED_PROPERTIES_KEY] : []),
    ]);
  };
  /**
   * Batch collection projector: only `{ value: [...projected items] }`, never collection
   * metadata or continuation links. Items are clipped to the subrequest's `$top` (or this
   * contract's default when omitted, never above the page maximum), with `hasMore: true`
   * when Graph returned more, matching the single-request result limits.
   */
  const projectCollection = (
    body: unknown,
    query: ReadonlyArray<readonly [string, string]>,
  ): unknown => {
    if (typeof body !== "object" || body === null || Array.isArray(body)) return {};
    const { value } = body as Record<string, unknown>;
    if (!Array.isArray(value)) return {};
    const requested = query.find(([name]) => name === WIRE_NAMES.top)?.[1];
    const limit = Math.min(
      requested === undefined ? (spec.top?.default ?? 0) : Number(requested),
      pageMax,
    );
    const items = value
      .slice(0, limit)
      .map(projectItemWith(collectionKeys(query.find(([n]) => n === WIRE_NAMES.select)?.[1])))
      .filter((item) => item !== undefined);
    // Graph honours $top and signals further pages with a nextLink, so its presence (not only
    // an oversized page) means the listing is incomplete; the link itself is never returned.
    const hasMore =
      value.length > limit ||
      (body as Record<string, unknown>)["@odata.nextLink"] !== undefined ||
      (spec.top?.offsetPaging === true && value.length >= limit);
    if (!hasMore) return { value: items };
    const out: Record<string, unknown> = { value: items, hasMore: true };
    if (spec.top?.offsetPaging === true && spec.skip !== undefined) {
      const searched = query.some(([name]) => name === WIRE_NAMES.search);
      const skip = Number(query.find(([name]) => name === WIRE_NAMES.skip)?.[1] ?? 0);
      const next = searched
        ? undefined
        : nextSkipFor(skip, Math.min(value.length, limit), spec.skip.max);
      if (next !== undefined) out["nextSkip"] = next;
    }
    return out;
  };

  const queryKeys = (Object.keys(WIRE_NAMES) as OptionKey[]).filter(
    (k) => k in shape && !(k === "top" && spec.top?.wire === false),
  );
  // Site search is documented as the bare `search` (no `$`); every other contract uses `$search`.
  const siteSearch = spec.search?.siteSearch === true;
  const searchWire = siteSearch ? "search" : WIRE_NAMES.search;
  const wireToKey = new Map<string, OptionKey>(
    queryKeys.map((k) => [k === "search" ? searchWire : WIRE_NAMES[k], k]),
  );
  const schema = z.strictObject(shape);

  /** The single validation used by both the tool handler and the batch policy. */
  function validate(raw: Record<string, unknown>): Record<string, unknown> {
    const parsed = schema.safeParse(raw);
    if (!parsed.success) {
      // Field names come from the schema itself; offending values are never echoed.
      const bad = new Set<string>();
      let unknown = 0;
      for (const issue of parsed.error.issues) {
        if (issue.code === "unrecognized_keys") unknown += issue.keys.length;
        else if (typeof issue.path[0] === "string" && issue.path[0] in shape) {
          bad.add(issue.path[0]);
        }
      }
      const parts = [...bad, ...(unknown > 0 ? [`${unknown} unsupported option(s)`] : [])];
      throw invalid(`Invalid query options: ${parts.join(", ") || "(arguments)"}.`);
    }
    const typed = parsed.data as Record<string, unknown>;
    if (typed["search"] !== undefined) {
      // Conservative project rule: Graph's $search ranks/sorts by sent date and pages itself;
      // learn.microsoft.com documents no supported combination with these options.
      if (typed["orderby"] !== undefined) throw invalid("search cannot be combined with orderby.");
      if (typed["skip"] !== undefined) throw invalid("search cannot be combined with skip.");
      if (typed["filter"] !== undefined) throw invalid("search cannot be combined with filter.");
    }
    if (spec.window !== undefined) {
      // Exact (sub-millisecond) comparison: fraction digits are part of the instant.
      assertWindow(
        typed["startDateTime"] as string,
        typed["endDateTime"] as string,
        spec.window.maxDays,
      );
    }
    if (
      typeof typed["filter"] === "string" &&
      typed["orderby"] !== undefined &&
      spec.orderby?.requireInFilter !== false
    ) {
      const violation = orderbyFilterViolation(typed["filter"], typed["orderby"] as Orderby[]);
      if (violation !== undefined) throw invalid(violation);
    }
    return typed;
  }

  const toRequest = (args: ToolArgs): QueryRequest => {
    const raw: Record<string, unknown> = {};
    for (const key of Object.keys(shape)) {
      if (args[key] !== undefined) raw[key] = args[key];
    }
    const typed = validate(raw);
    const query: Record<string, string | number | boolean> = {};
    const limit =
      spec.top === undefined ? 0 : ((typed["top"] as number | undefined) ?? spec.top.default);
    if (spec.top !== undefined && spec.top.wire !== false) {
      query[WIRE_NAMES.top] = Math.min(limit, pageMax);
    }
    const chosen = (typed["select"] as string[] | undefined) ?? spec.select?.default;
    if (chosen !== undefined) query[WIRE_NAMES.select] = [...new Set(chosen)].join(",");
    if (typed["expand"] !== undefined) {
      query[WIRE_NAMES.expand] = [...new Set(typed["expand"] as string[])].join(",");
    }
    if (typeof typed["filter"] === "string") query[WIRE_NAMES.filter] = typed["filter"];
    if (typeof typed["search"] === "string") {
      query[searchWire] = siteSearch ? typed["search"] : kqlQuote(typed["search"]);
    }
    if (typed["orderby"] !== undefined) {
      query[WIRE_NAMES.orderby] = orderbyWire(typed["orderby"] as Orderby[]);
    }
    if (typeof typed["skip"] === "number") query[WIRE_NAMES.skip] = typed["skip"];
    if (typeof typed["includeHiddenFolders"] === "boolean") {
      query[WIRE_NAMES.includeHiddenFolders] = typed["includeHiddenFolders"];
    }
    if (spec.window !== undefined) {
      // Validated above, so canonicalization cannot fail; the UTC form is what Graph sees.
      query[WIRE_NAMES.startDateTime] = canonicalInstant(
        typed["startDateTime"] as string,
      ) as string;
      query[WIRE_NAMES.endDateTime] = canonicalInstant(typed["endDateTime"] as string) as string;
    }
    if (typeof typed["extendedPropertyId"] === "string") {
      query[WIRE_NAMES.extendedPropertyId] = extendedPropertyExpand(typed["extendedPropertyId"]);
    }
    const headers: Record<string, string> = {};
    const prefer: string[] = [];
    if (typeof typed["timeZone"] === "string") {
      // The allowlist's canonical spelling is what goes on the wire.
      prefer.push(timeZonePrefer(canonicalTimeZone(typed["timeZone"]) as string));
    }
    if (spec.bodyType !== undefined) {
      const type = (typed["bodyType"] as string | undefined) ?? spec.bodyType.default;
      prefer.push(bodyPrefer(type));
    }
    if (prefer.length > 0) headers["Prefer"] = prefer.join(", ");
    if (spec.search?.directoryProperties !== undefined && typeof typed["search"] === "string") {
      // Directory `$search` works only in advanced query mode (Graph: header AND `$count=true`).
      query["$count"] = true;
      headers["ConsistencyLevel"] = "eventual";
    }
    return { query, headers, limit };
  };

  const checkHeaders = (headers: Readonly<Record<string, string>>): void => {
    // Exact allowlist: ConsistencyLevel is never sent, and Prefer only as the canonical list
    // `toRequest` emits: [outlook.timezone="<allowlisted, canonical>"][, ]
    // [outlook.body-content-type="text|html"], each part only if this tool declares it.
    // Anything else (unsafe-html, odata.*, other casings, duplicates, reordering) fails.
    const denied = (): GraphClientError =>
      invalid("Request header is not permitted for this endpoint.");
    const needBody = (): GraphClientError =>
      invalid("A body-content-type Prefer header is required in a batch for this endpoint.");
    for (const name of Object.keys(headers)) {
      if (name !== "Prefer") throw denied();
    }
    const value = headers["Prefer"];
    if (value === undefined) {
      // Batch cannot inject the single-request text default (Graph would return HTML).
      if (spec.bodyType !== undefined) throw needBody();
      return;
    }
    const parts = value.split(", ");
    let at = 0;
    const zone = spec.timeZone === true ? TZ_PREFER_RE.exec(parts[at] ?? "")?.[1] : undefined;
    if (zone !== undefined) {
      if (canonicalTimeZone(zone) !== zone) throw denied();
      at += 1;
    }
    if (spec.bodyType !== undefined) {
      if (!BODY_PREFER_RE.test(parts[at] ?? "")) throw needBody();
      at += 1;
    }
    if (at !== parts.length) throw denied();
  };

  const check: BatchPolicy["check"] = (pairs, headers) => {
    checkHeaders(headers);
    const raw: Record<string, unknown> = {};
    for (const [name, value] of pairs) {
      const key = wireToKey.get(name);
      if (key === undefined) throw invalid("Unsupported query parameter.");
      if (key in raw) throw invalid("Duplicate query parameter.");
      switch (key) {
        case "top":
        case "skip":
          raw[key] = strictInt(value, WIRE_NAMES[key]);
          break;
        case "select":
        case "expand":
          raw[key] = value.split(",");
          break;
        case "orderby":
          raw[key] = parseOrderbyWire(value);
          break;
        case "search": {
          if (spec.search?.directoryProperties !== undefined) {
            throw invalid(
              "$search on this endpoint needs advanced query mode, which graph-batch does not support; call the tool directly.",
            );
          }
          if (siteSearch) {
            // Bare `search=<text>`: the wire value IS the typed text; `validate` checks it.
            raw[key] = value;
            break;
          }
          const text = kqlUnquote(value);
          if (text === undefined) {
            throw invalid(
              spec.search?.plainName === true
                ? "$search must be one double-quoted plain search string."
                : "$search must be one double-quoted KQL expression.",
            );
          }
          raw[key] = text;
          break;
        }
        case "startDateTime":
        case "endDateTime":
          // The wire value is what Graph receives, so it must be the canonical UTC form the
          // single-request path emits: no offsets (a raw "+" decodes to a space), no spellings
          // like ".0000000Z".
          if (canonicalInstant(value) !== value) {
            throw invalid(`${WIRE_NAMES[key]} must be a canonical UTC timestamp ending in Z.`);
          }
          raw[key] = value;
          break;
        case "extendedPropertyId": {
          const id = parseExtendedPropertyExpand(value);
          if (id === undefined) {
            throw invalid(
              "$expand must be exactly singleValueExtendedProperties($filter=id eq '<documented id>').",
            );
          }
          raw[key] = id;
          break;
        }
        case "includeHiddenFolders":
          if (value !== "true" && value !== "false") {
            throw invalid("includeHiddenFolders must be true or false.");
          }
          raw[key] = value === "true";
          break;
        default:
          raw[key] = value;
      }
    }
    const typed = validate(raw);
    if (spec.select?.requiredInBatch === true && typed["select"] === undefined) {
      throw invalid("$select is required in a batch for this endpoint.");
    }
    // Batch cannot paginate, so it is bound to the per-request page size, not the item limit.
    if (typeof typed["top"] === "number" && typed["top"] > pageMax) {
      throw invalid(`$top exceeds the batch maximum of ${pageMax}.`);
    }
  };

  /**
   * Strict entity projection for ONE request: the keep set is derived from that request's
   * effective `$select`/`$expand` wire values, never the whole allowlist, so an over-returning
   * Graph response cannot surface allowlisted fields the caller did not ask for. Wire names are
   * intersected with the allowlists (defence in depth; both paths already validated them).
   */
  const requestKeys = (selectWire: unknown, expandWire: unknown): ReadonlySet<string> => {
    const names = (wire: unknown): string[] =>
      typeof wire === "string" && wire.length > 0 ? wire.split(",") : [];
    const allowedSelect = new Set<string>(spec.select?.allowed ?? []);
    const allowedExpand = new Set<string>(spec.expand?.allowed ?? []);
    // No $select on the wire: Graph's own default set applies, which is not known here, so
    // the contract default (when declared) bounds it; with neither, keep the allowlist.
    const selected =
      selectWire !== undefined
        ? names(selectWire)
        : [...(spec.select?.default ?? spec.select?.allowed ?? [])];
    return new Set([
      "id",
      ...selected.filter((n) => allowedSelect.has(n)),
      ...names(expandWire).filter((n) => allowedExpand.has(n)),
      ...(spec.projectEntity?.keep ?? []),
    ]);
  };
  const projectForRequest = (item: unknown, keys: ReadonlySet<string>): unknown => {
    if (typeof item !== "object" || item === null || Array.isArray(item)) return undefined;
    return Object.fromEntries(
      Object.entries(item).filter(([k]) => keys.has(k) || ITEM_ANNOTATIONS.has(k)),
    );
  };

  // Single-entity tools: identity, except that expanded extended properties are reduced.
  const entityProjector = (item: unknown): unknown => {
    if (spec.projectEntity !== undefined) return projectItem(item);
    return spec.extendedProperty === true ? projectExtendedProperties(item) : item;
  };
  const batchPolicyFor = (checkFn: BatchPolicy["check"]): BatchPolicy => {
    if (spec.top !== undefined) return { check: checkFn, project: projectCollection };
    if (spec.extendedProperty === true) {
      return { check: checkFn, project: (body) => entityProjector(body) };
    }
    // Strict entity projection from the subrequest's own (already validated) query, exactly
    // like the single path; a non-object body becomes {} (never raw), like collections.
    if (spec.projectEntity !== undefined) {
      return {
        check: checkFn,
        project: (body, pairs) =>
          projectForRequest(
            body,
            requestKeys(
              pairs.find(([n]) => n === WIRE_NAMES.select)?.[1],
              pairs.find(([n]) => n === WIRE_NAMES.expand)?.[1],
            ),
          ) ?? {},
      };
    }
    return { check: checkFn };
  };

  return {
    shape,
    toRequest,
    batchPolicy: batchPolicyFor(check),
    projectsEntity: spec.projectEntity !== undefined,
    projectItem: spec.top !== undefined ? projectItem : entityProjector,
    itemProjector: (query) => projectItemWith(collectionKeys(query[WIRE_NAMES.select])),
    projectResponse: (item, query) =>
      spec.projectEntity !== undefined
        ? projectForRequest(item, requestKeys(query[WIRE_NAMES.select], query[WIRE_NAMES.expand]))
        : entityProjector(item),
  };
}
