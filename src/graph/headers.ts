import { GraphClientError } from "./errors.js";

/** Request headers callers may add. Everything else (Authorization, Host, ...) is refused. */
const ALLOWED_EXTRA_HEADERS: ReadonlyMap<string, string> = new Map([
  ["consistencylevel", "ConsistencyLevel"],
  ["prefer", "Prefer"],
]);

const MAX_HEADER_VALUE = 256;

/**
 * Validates caller-supplied request headers against the allowlist and returns them with
 * canonical names. Values must be printable ASCII (space allowed, e.g. outlook.timezone="UTC").
 * Shared by single requests and batch subrequests so both honor one allowlist.
 */
export function validateExtraHeaders(headers: unknown): Record<string, string> {
  const out: Record<string, string> = {};
  if (headers === undefined) return out;
  if (typeof headers !== "object" || headers === null || Array.isArray(headers)) {
    throw new GraphClientError("invalid_input", "Request header is not permitted.");
  }
  for (const [name, value] of Object.entries(headers)) {
    const canonical = ALLOWED_EXTRA_HEADERS.get(name.toLowerCase());
    if (
      canonical === undefined ||
      typeof value !== "string" ||
      value.length > MAX_HEADER_VALUE ||
      /[^\x20-\x7e]/.test(value) ||
      Object.hasOwn(out, canonical)
    ) {
      throw new GraphClientError("invalid_input", "Request header is not permitted.");
    }
    out[canonical] = value;
  }
  return out;
}
