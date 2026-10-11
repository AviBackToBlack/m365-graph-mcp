/**
 * Single-value extended property expansion for calendar event tools
 * (`supportsExpandExtendedProperties` in the catalog).
 *
 * Microsoft Graph documents exactly one way to read a single-value extended property:
 * `$expand=singleValueExtendedProperties($filter=id eq '<id>')` (singlevaluelegacyextendedproperty-get),
 * and says all extended properties of an instance cannot be read in one call. So the caller
 * supplies only a typed property id, never OData: `extendedPropertyExpand` is the single place
 * that builds the expansion, and batch accepts that exact form only (`parseExtendedPropertyExpand`).
 *
 * Documented id formats (extended-properties-overview, "Valid id formats for single-value
 * extended properties"):
 *   `{type} {guid} Name {name}`   e.g. `String {8ECCC264-6880-4EBE-992F-8888D2EEAA1D} Name TestProperty`
 *   `{type} {guid} Id {id}`       e.g. `Integer {8ECCC264-6880-4EBE-992F-8888D2EEAA1D} Id 0x8012`
 *   `{type} {proptag}`            e.g. `String 0x4001`
 * The Graph pages do not list the `{type}` values. They are the scalar (non-array) members of the
 * Exchange `MapiPropertyTypeType` enumeration, minus Error, Null and Object (not meaningful as
 * a single string value); anything else is rejected. Keywords and types are exact-case, as in
 * the docs; the GUID must be braced; the name is case-sensitive in Graph and kept as given.
 * Names are limited to a conservative charset (no quote, parenthesis, `$`, `/`, `\`, `&`,
 * control character), so no id can break out of the OData literal; `quoteOData` still doubles
 * single quotes as defense in depth.
 *
 * Reading extended properties needs Calendars.Read (Calendars.ReadBasic excludes them).
 */

const TYPES = [
  "ApplicationTime",
  "Binary",
  "Boolean",
  "CLSID",
  "Currency",
  "Double",
  "Float",
  "Integer",
  "Long",
  "Short",
  "String",
  "SystemTime",
] as const;

const TYPE = `(?:${TYPES.join("|")})`;
const GUID = "\\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\\}";
const HEX = "0x[0-9A-Fa-f]{1,8}";
const NAME = "[A-Za-z0-9_.:-](?:[A-Za-z0-9 _.:-]{0,126}[A-Za-z0-9_.:-])?";
const ID_RE = new RegExp(`^${TYPE} (?:${GUID} (?:Name ${NAME}|Id ${HEX})|${HEX})$`);

/** Hard cap checked before the pattern. */
export const MAX_EXTENDED_PROPERTY_ID_LENGTH = 200;

/** The navigation property that carries expanded single-value extended properties. */
export const EXTENDED_PROPERTIES_KEY = "singleValueExtendedProperties";

/**
 * Documented numeric ranges: a GUID-backed `Id` (named property id) is 0x8000-0xFFFE; a
 * bare proptag (`{type} 0x...`) is 0x0001-0x7FFF. Anything else would only fail at Graph.
 */
const GUID_ID_MIN = 0x8000;
const GUID_ID_MAX = 0xfffe;
const PROPTAG_MIN = 0x0001;
const PROPTAG_MAX = 0x7fff;

export function isExtendedPropertyId(value: string): boolean {
  if (value.length > MAX_EXTENDED_PROPERTY_ID_LENGTH || !ID_RE.test(value)) return false;
  // Anchored to the `{type} {guid} Id 0x..` form, so a Name-form id whose name merely ends in
  // " Id 0x.." is not range-checked as a numeric id.
  const guidId = /^\S+ \{[0-9A-Fa-f-]{36}\} Id (0x[0-9A-Fa-f]+)$/.exec(value)?.[1];
  if (guidId !== undefined) {
    const n = Number.parseInt(guidId, 16);
    return n >= GUID_ID_MIN && n <= GUID_ID_MAX;
  }
  const proptag = /^\S+ (0x[0-9A-Fa-f]+)$/.exec(value)?.[1];
  if (proptag !== undefined) {
    const n = Number.parseInt(proptag, 16);
    return n >= PROPTAG_MIN && n <= PROPTAG_MAX;
  }
  return true; // the `{guid} Name {name}` form has no numeric part
}

const quoteOData = (text: string): string => `'${text.replace(/'/g, "''")}'`;

/** The only `$expand` value this project ever sends. `id` must satisfy `isExtendedPropertyId`. */
export function extendedPropertyExpand(id: string): string {
  return `${EXTENDED_PROPERTIES_KEY}($filter=id eq ${quoteOData(id)})`;
}

const EXPAND_RE = /^singleValueExtendedProperties\(\$filter=id eq '([^']+)'\)$/;

/** Inverse of `extendedPropertyExpand`: the id, or undefined for anything but that exact form. */
export function parseExtendedPropertyExpand(wire: string): string | undefined {
  const id = EXPAND_RE.exec(wire)?.[1];
  return id !== undefined && isExtendedPropertyId(id) ? id : undefined;
}

/**
 * Reduces an entity's expanded extended properties to `{ id, value }` string pairs (anything
 * else in that collection, or a malformed collection, is dropped). Other keys are untouched.
 */
export function projectExtendedProperties(item: unknown): unknown {
  if (typeof item !== "object" || item === null || Array.isArray(item)) return item;
  const record = item as Record<string, unknown>;
  if (!(EXTENDED_PROPERTIES_KEY in record)) return item;
  const { [EXTENDED_PROPERTIES_KEY]: raw, ...rest } = record;
  if (!Array.isArray(raw)) return rest;
  const kept = raw.flatMap((entry): Array<{ id: string; value: string }> => {
    if (typeof entry !== "object" || entry === null) return [];
    const { id, value } = entry as Record<string, unknown>;
    return typeof id === "string" && typeof value === "string" ? [{ id, value }] : [];
  });
  return { ...rest, [EXTENDED_PROPERTIES_KEY]: kept };
}
