/**
 * Shared, advisory JWT payload inspection used by the manual token provider and
 * the Graph Explorer bridge acceptance gate. Nothing here verifies signatures.
 */

/** Accepted Graph audiences (`aud`): resource URI forms and the Graph app ID. */
export const GRAPH_AUDIENCES: ReadonlySet<string> = new Set([
  "https://graph.microsoft.com",
  "https://graph.microsoft.com/",
  "00000003-0000-0000-c000-000000000000",
]);

export const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const BASE64URL_RE = /^[A-Za-z0-9_-]+$/;
/** A bearer must be one header-safe value: visible ASCII only (no whitespace/control/non-ASCII). */
export const HEADER_SAFE_RE = /^[!-~]+$/;
/** Generous upper bound on token length (16 KiB). */
export const MAX_TOKEN_LENGTH = 16_384;
/** Valid JS Date range in epoch milliseconds (ECMAScript time value limit). */
export const MAX_DATE_MS = 8.64e15;

export interface Claims {
  aud?: unknown;
  exp?: unknown;
  nbf?: unknown;
  tid?: unknown;
  scp?: unknown;
  roles?: unknown;
  ver?: unknown;
  appid?: unknown;
  azp?: unknown;
}

/** Converts a seconds claim to epoch ms; undefined if non-finite or outside the Date range. */
export function claimSecondsToMs(seconds: number): number | undefined {
  const ms = seconds * 1000;
  return Number.isFinite(ms) && Math.abs(ms) <= MAX_DATE_MS ? ms : undefined;
}

/**
 * Best-effort, advisory decode of a JWS compact-form JWT payload WITHOUT
 * verifying the signature. Returns `undefined` when the token is not a
 * decodable JWT (opaque token, JWE, undecodable or non-object payload); that
 * is never an error by itself.
 *
 * Microsoft treats access tokens as opaque to clients (they may be non-JWT or
 * encrypted), and Graph performs the authoritative validation. The claim
 * checks built on this are early, clearer errors only, and run only when the
 * token is decodable. Never put raw input or parser error text in messages.
 */
export function tryDecodeClaims(token: string): Claims | undefined {
  const parts = token.split(".");
  const [header, payload, signature] = parts;
  if (
    parts.length !== 3 ||
    !header ||
    !payload ||
    !signature ||
    !BASE64URL_RE.test(header) ||
    !BASE64URL_RE.test(payload) ||
    !BASE64URL_RE.test(signature)
  ) {
    return undefined;
  }
  try {
    const parsed: unknown = JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
    if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) return undefined;
    return parsed as Claims;
  } catch {
    return undefined;
  }
}
