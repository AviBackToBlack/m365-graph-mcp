import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { isRecord, pickScalars } from "./chats.js";
import { withProjection } from "./chat-messages.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";
import { withSanitizer } from "./teams.js";

/**
 * SharePoint lists, list items and columns (SM-13 part C): `list-sharepoint-site-lists`,
 * `get-sharepoint-site-list`, `list-sharepoint-site-list-items`, `get-sharepoint-site-list-item`,
 * `list-sharepoint-list-columns` and `get-sharepoint-list-column` (all v1.0 GET).
 *
 * Delegated permissions (current Microsoft docs, October 2026; see PERMISSIONS.md): list-list,
 * list-get, listitem-list and listitem-get document `Sites.Read.All` as least privileged;
 * list-list-columns and columndefinition-get document `Sites.Read.All` too (higher:
 * `Sites.Manage.All`, `Sites.FullControl.All`, `Sites.ReadWrite.All`). Personal Microsoft accounts
 * are "Not supported" on all six. No page lists `Sites.Selected`; D30 keeps it as the
 * resource-scoped alternative (it needs a grant on the target site). The server never pre-checks
 * scopes; Graph decides, and a 403 is a `capability_gap`.
 *
 * Query options (typed only; the pages are the authority):
 * - lists (collection): the page documents no OData option except that system-faceted lists are
 *   hidden unless `system` is in `$select`. Only a local `top` (no `$top` sent, pages followed) is
 *   offered, so hidden/system lists are never requested.
 * - list item collection / item: `$expand` (and, on the collection, `$filter`) are documented; no
 *   `$top`. Only a local `top` (collection) and the typed `expand: ["fields"]` are offered. No
 *   `$filter` (so no `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`), no nested
 *   `fields($select=...)`, no `$expand=driveItem` (content) and no arbitrary `$expand`.
 * - columns (collection): "some of the OData query parameters" and none named, so only a local
 *   `top`. Hidden columns are shown only when `hidden` is in `$select`, which is never sent.
 * - list, column: nothing documented that is offered.
 * `list-id` and `columnDefinition-id` are GUIDs and `listItem-id` is digits (see
 * `graph/placeholders.ts`); the documented `{list-title}` addressing is deliberately not accepted.
 *
 * Content policy (single path and `graph-batch` alike, allowlist-only, all values untrusted data):
 * - list: `id`, `name`, `displayName`, `description`, `webUrl`, `createdDateTime`,
 *   `lastModifiedDateTime` and `list.{template, hidden, contentTypesEnabled}`.
 * - list item: `id`, `webUrl`, `createdDateTime`, `lastModifiedDateTime`, `contentType.{id, name}`
 *   and, ONLY when that request's `$expand` named it, `fields` reduced by `sanitizeFields`.
 * - column: `id`, `name`, `displayName`, `description`, `columnGroup`, `required`, `indexed`,
 *   `hidden`, `readOnly`, `enforceUniqueValues`, `columnType` (the one type facet present); the
 *   single-column tool additionally returns that facet reduced to allowlisted scalars.
 * Identities (`createdBy`, `lastModifiedBy`), `eTag`, `parentReference`, `sharepointIds`,
 * `driveItem`, relationships, `defaultValue`, `validation`, formulas, lookup list ids and every
 * unknown property are dropped.
 */

type Rec = Record<string, unknown>;

// --- Lists ---------------------------------------------------------------------------------------

const LIST_KEYS = [
  "name",
  "displayName",
  "description",
  "webUrl",
  "createdDateTime",
  "lastModifiedDateTime",
] as const;
const LIST_FACET_KEYS = ["template", "hidden", "contentTypesEnabled"] as const;

/** Allowlist projection for one list (single, collection item, batch body). */
export function sanitizeList(item: unknown): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, ["id", ...LIST_KEYS]);
  const facet = item["list"];
  if (isRecord(facet)) out["list"] = pickScalars(facet, LIST_FACET_KEYS);
  return out;
}

// --- List items ----------------------------------------------------------------------------------

const ITEM_KEYS = ["webUrl", "createdDateTime", "lastModifiedDateTime"] as const;

/** Bounds on the untrusted `fields` bag. */
export const FIELDS_MAX_COUNT = 200;
export const FIELD_STRING_MAX = 2000;
export const FIELD_ARRAY_MAX = 50;
export const FIELD_TRUNCATION_MARKER = "…[truncated]";
/** SharePoint internal column names (`_x0020_` style encoding included); no `@` annotations. */
const FIELD_NAME_RE = /^[A-Za-z_][A-Za-z0-9_]{0,254}$/;
/**
 * `{InternalName}LookupId` scalars: for person/group columns (including the creator/modifier
 * `AuthorLookupId`, `EditorLookupId`, `AppAuthorLookupId`, `AppEditorLookupId`) they are site
 * user-information ids, i.e. identities the content policy does not expose; for lookup columns
 * they are ids into a list that is never joined. Every such key is dropped (D35, after review).
 */
const LOOKUP_ID_RE = /lookupid$/i;

/** Cuts a string so that it, including the truncation marker, fits `FIELD_STRING_MAX`. */
const clip = (value: string): { text: string; cut: boolean } =>
  value.length > FIELD_STRING_MAX
    ? {
        text: `${value.slice(0, FIELD_STRING_MAX - FIELD_TRUNCATION_MARKER.length)}${FIELD_TRUNCATION_MARKER}`,
        cut: true,
      }
    : { text: value, cut: false };

/**
 * Reduces a `fieldValueSet` to bounded scalar column values: string, finite number, boolean, and
 * arrays of strings (multi-choice). Annotation keys (anything outside the internal-name grammar,
 * e.g. `@odata.etag`, `Col@odata.type`), `*LookupId` user/lookup references, nested objects,
 * mixed or non-string arrays, nulls, `__proto__` and anything past `FIELDS_MAX_COUNT` are dropped;
 * long strings are cut. `truncated` is true when a cap (not a type rule) removed data.
 */
export function sanitizeFields(fields: unknown): { fields: Rec; truncated: boolean } | undefined {
  if (!isRecord(fields)) return undefined;
  const entries: Array<[string, unknown]> = [];
  let truncated = false;
  for (const [key, value] of Object.entries(fields)) {
    if (!FIELD_NAME_RE.test(key) || key === "__proto__" || LOOKUP_ID_RE.test(key)) continue;
    let kept: unknown;
    if (typeof value === "string") {
      const c = clip(value);
      truncated ||= c.cut;
      kept = c.text;
    } else if (typeof value === "number") {
      if (Number.isFinite(value)) kept = value;
    } else if (typeof value === "boolean") {
      kept = value;
    } else if (Array.isArray(value) && value.every((v) => typeof v === "string")) {
      const clipped = (value as string[]).slice(0, FIELD_ARRAY_MAX).map(clip);
      truncated ||= value.length > FIELD_ARRAY_MAX || clipped.some((c) => c.cut);
      kept = clipped.map((c) => c.text);
    }
    if (kept === undefined) continue;
    if (entries.length >= FIELDS_MAX_COUNT) {
      truncated = true;
      break;
    }
    entries.push([key, kept]);
  }
  // fromEntries defines own data properties (no prototype setter involved).
  return { fields: Object.fromEntries(entries), truncated };
}

/**
 * Allowlist projection for one list item. `fields` is kept only when that request asked for it
 * (`withFields`, from the wire `$expand`); an over-returned one is dropped.
 */
export function sanitizeListItem(item: unknown, withFields = false): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, ["id", ...ITEM_KEYS]);
  const contentType = item["contentType"];
  if (isRecord(contentType)) out["contentType"] = pickScalars(contentType, ["id", "name"]);
  if (withFields && Object.hasOwn(item, "fields")) {
    const reduced = sanitizeFields(item["fields"]);
    if (reduced !== undefined) {
      out["fields"] = reduced.fields;
      if (reduced.truncated) out["fieldsTruncated"] = true;
    }
  }
  return out;
}

const wantsFields = (query: Readonly<Record<string, unknown>>): boolean => {
  const expand = query["$expand"];
  return typeof expand === "string" && expand.split(",").includes("fields");
};

// --- Columns -------------------------------------------------------------------------------------

const COLUMN_KEYS = [
  "name",
  "displayName",
  "description",
  "columnGroup",
  "required",
  "indexed",
  "hidden",
  "readOnly",
  "enforceUniqueValues",
] as const;

/**
 * The mutually exclusive type facets of `columnDefinition` and the scalar settings of each that the
 * single-column tool returns. Not exposed: `lookup` list/column ids, `calculated.formula`,
 * `defaultValue`, `validation`, `sourceContentType`, `term` and every unknown property.
 */
const FACET_KEYS: Readonly<Record<string, readonly string[]>> = {
  boolean: [],
  calculated: ["format", "outputType"],
  choice: ["allowTextEntry", "displayAs"],
  contentApprovalStatus: [],
  currency: ["locale"],
  dateTime: ["displayAs", "format"],
  geolocation: [],
  hyperlinkOrPicture: ["isPicture"],
  lookup: ["allowMultipleValues", "allowUnlimitedLength"],
  number: ["decimalPlaces", "displayAs", "maximum", "minimum"],
  personOrGroup: ["allowMultipleSelection", "chooseFromType", "displayAs"],
  term: [],
  text: ["allowMultipleLines", "maxLength", "textType"],
  thumbnail: [],
};
const CHOICES_MAX = 100;
const CHOICE_MAX_CHARS = 255;

/**
 * Allowlist projection for one column. `columnType` names the one type facet present (none for a
 * type Graph does not model); `withDetail` (single-column tool) also returns that facet reduced to
 * its allowlisted scalars (choice values are bounded strings).
 */
export function sanitizeColumn(item: unknown, withDetail: boolean): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, ["id", ...COLUMN_KEYS]);
  const facet = Object.keys(FACET_KEYS).find((k) => isRecord(item[k]));
  if (facet === undefined) return out;
  out["columnType"] = facet;
  if (withDetail) {
    const body = item[facet] as Rec;
    const detail = pickScalars(body, FACET_KEYS[facet] ?? []);
    const choices = body["choices"];
    if (facet === "choice" && Array.isArray(choices)) {
      const strings = choices.filter((c): c is string => typeof c === "string");
      detail["choices"] = strings.slice(0, CHOICES_MAX).map((c) => c.slice(0, CHOICE_MAX_CHARS));
      if (strings.length > CHOICES_MAX || strings.some((c) => c.length > CHOICE_MAX_CHARS)) {
        detail["choicesTruncated"] = true;
      }
    }
    out[facet] = detail;
  }
  return out;
}

// --- Contracts -----------------------------------------------------------------------------------

const listsContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      // The page documents no `$top`: `top` is a local limit and pages are followed.
      top: { default: 25, max: 200, wire: false },
      // Load-bearing on a collection contract: it widens the inner batch keep-set
      // (`projectCollection`) so list scalars and the `list` facet survive until `sanitizeList`.
      projectEntity: { keep: [...LIST_KEYS, "list"] },
    }),
    true,
    sanitizeList,
  );

const listContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({ projectEntity: { keep: [...LIST_KEYS, "list"] } }),
    false,
    sanitizeList,
  );

// Query-aware projection (as for channel tabs): `fields` survives only when the request's own
// `$expand` named it, on the single path and in batch alike.
const itemsContract = (): QueryContract =>
  withProjection(
    defineQueryContract({
      // The page documents `$filter`/`$expand` but no `$top`: local limit, pages are followed.
      top: { default: 25, max: 100, wire: false },
      expand: { allowed: ["fields"] },
      projectEntity: { keep: [...ITEM_KEYS, "contentType"] },
    }),
    true,
    (item, query) => sanitizeListItem(item, wantsFields(query)),
  );

const itemContract = (): QueryContract =>
  withProjection(
    defineQueryContract({
      expand: { allowed: ["fields"] },
      projectEntity: { keep: [...ITEM_KEYS, "contentType"] },
    }),
    false,
    (item, query) => sanitizeListItem(item, wantsFields(query)),
  );

const columnsContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      // "Some of the OData query parameters" and none named: local limit, pages are followed.
      top: { default: 50, max: 200, wire: false },
      projectEntity: { keep: [...COLUMN_KEYS, ...Object.keys(FACET_KEYS)] },
    }),
    true,
    (item) => sanitizeColumn(item, false),
  );

const columnContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      projectEntity: { keep: [...COLUMN_KEYS, ...Object.keys(FACET_KEYS)] },
    }),
    false,
    (item) => sanitizeColumn(item, true),
  );

export function registerSharePointListTools(
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
  const list = (name: string, contract: QueryContract) =>
    handlers.set(name, listTool(client, name, patternOf(name), contract));
  const get = (name: string, contract: QueryContract) =>
    handlers.set(name, getTool(client, name, patternOf(name), contract));
  list("list-sharepoint-site-lists", listsContract());
  get("get-sharepoint-site-list", listContract());
  list("list-sharepoint-site-list-items", itemsContract());
  get("get-sharepoint-site-list-item", itemContract());
  list("list-sharepoint-list-columns", columnsContract());
  get("get-sharepoint-list-column", columnContract());
}
