import {
  GRAPH_AUDIENCES,
  HEADER_SAFE_RE,
  MAX_DATE_MS,
  MAX_TOKEN_LENGTH,
  claimSecondsToMs,
  tryDecodeClaims,
} from "../auth/jwt-claims.js";
import type { Claims } from "../auth/jwt-claims.js";
import { DISABLED_ACCEPTANCE_DIAGNOSTICS } from "./acceptance-diagnostics.js";
import type {
  AcceptanceDiagnostics,
  AudienceCategory,
  GraphRequestOrigin,
  TokenShape,
} from "./acceptance-diagnostics.js";

/**
 * Local token acceptance sanity gate (D44). This is NOT cryptographic
 * validation: only the JWT payload is decoded, no signature is verified, and
 * decoded claims prove nothing about authenticity. The trust anchor is that the
 * bearer was observed leaving the bridge-owned Graph Explorer session for the
 * exact Graph origin. Microsoft tells clients to treat access tokens as opaque;
 * opaque or undecodable tokens are refused (`unsupported_token_format`).
 *
 * Rejection reasons are fixed category strings. They never contain the token,
 * claim values, tenant ids, scopes, audiences or client ids.
 */

export type RejectReason =
  | "wrong_source"
  | "wrong_origin"
  | "no_response"
  | "http_401"
  | "unsupported_token_format"
  | "clock_invalid"
  | "audience"
  | "tenant"
  | "not_delegated"
  | "lifetime"
  | "not_yet_valid"
  | "client_pin";

export type ClientPinClaim = "auto" | "appid" | "azp";

export interface TokenGateConfig {
  /** Required tenant GUID; `tid` must equal it (GUIDs compare case-insensitively). */
  tenantId: string;
  /**
   * Optional client pin, off when absent. `auto` uses `appid` for v1 tokens
   * (`ver` "1.0") and `azp` for v2 tokens (`ver` "2.0"); any other `ver` is
   * rejected. `appid`/`azp` force one claim. Absent claim or mismatch rejects.
   */
  clientPin?: { guid: string; claim: ClientPinClaim };
}

export interface TokenCandidate {
  /** Raw bearer value without the `Bearer ` prefix. */
  bearer: string;
  requestUrl: string;
  /** HTTP status of the correlated Graph response; undefined when none was seen. */
  responseStatus: number | undefined;
  /** True only for requests from the bridge-owned Graph Explorer page/session. */
  sourceIsBridgeTarget: boolean;
}

export type GateResult =
  | { ok: true; bearer: string; expMs: number; remainingMs: number }
  | { ok: false; reason: RejectReason };

/** Minimum remaining lifetime for a token to be accepted (D44). */
export const MIN_REMAINING_MS = 5 * 60_000;
/** Tolerated `nbf` clock skew. */
export const NBF_SKEW_MS = 60_000;

const GRAPH_HOST = "graph.microsoft.com";

/**
 * Classifies a request URL against the exact Graph origin: https, host graph.microsoft.com, default
 * port, no userinfo. This is the single matcher: `isExactGraphOrigin` is derived from it, and the
 * D48 acceptance diagnostics print only the returned categories (never the URL).
 */
export function classifyGraphOrigin(requestUrl: string): GraphRequestOrigin {
  let url: URL;
  try {
    url = new URL(requestUrl);
  } catch {
    return { exactOriginMatch: false, scheme: "other", host: "other", port: "other" };
  }
  const scheme = url.protocol === "https:" ? "https" : "other";
  // WHATWG URL lowercases the host; no suffix matching.
  const host = url.hostname === GRAPH_HOST ? "expected_graph_host" : "other";
  // Default port (an explicit :443 is normalized away); an empty port on another scheme is not it.
  const port = scheme === "https" && url.port === "" ? "default_https" : "other";
  return {
    exactOriginMatch:
      scheme === "https" &&
      host === "expected_graph_host" &&
      port === "default_https" &&
      url.username === "" &&
      url.password === "",
    scheme,
    host,
    port,
  };
}

/** Exact origin check: https, host graph.microsoft.com, default port, no userinfo. */
export function isExactGraphOrigin(requestUrl: string): boolean {
  return classifyGraphOrigin(requestUrl).exactOriginMatch;
}

const AUDIENCE_CATEGORIES: ReadonlyMap<string, AudienceCategory> = new Map([
  ["https://graph.microsoft.com", "graph_url"],
  ["https://graph.microsoft.com/", "graph_url_trailing_slash"],
  ["00000003-0000-0000-c000-000000000000", "graph_application_id"],
]);

/**
 * Classifies an ACCEPTED token's decoded claims into the D48 shape categories. Runs locally on the
 * claims and returns categories only; no claim value, `exp` or identifier leaves this function.
 * Called only after every gate check passed, so rejected tokens never get a shape. Returns undefined
 * (no line) for an audience outside the mapped set, so a future accepted audience is never mislabeled.
 */
export function classifyTokenShape(
  claims: Claims,
  tenantId: string,
  remainingMs: number,
): TokenShape | undefined {
  const audience = typeof claims.aud === "string" ? AUDIENCE_CATEGORIES.get(claims.aud) : undefined;
  if (audience === undefined) return undefined;
  const hasAppid = typeof claims.appid === "string";
  const hasAzp = typeof claims.azp === "string";
  return {
    audience,
    clientClaim: hasAppid && hasAzp ? "both" : hasAppid ? "appid" : hasAzp ? "azp" : "none",
    version: claims.ver === "1.0" ? "v1" : claims.ver === "2.0" ? "v2" : "other_or_absent",
    scp: typeof claims.scp === "string" && claims.scp.trim().length > 0 ? "present" : "absent",
    roles: claims.roles === undefined ? "absent" : "present",
    tenant: guidEquals(claims.tid, tenantId) ? "matched" : "mismatch",
    remainingMinutes: Math.floor(remainingMs / 60_000),
  };
}

const reject = (reason: RejectReason): GateResult => ({ ok: false, reason });

function guidEquals(value: unknown, expected: string): boolean {
  return typeof value === "string" && value.toLowerCase() === expected.toLowerCase();
}

export function evaluateCandidate(
  candidate: TokenCandidate,
  config: TokenGateConfig,
  nowMs: number,
  diagnostics: AcceptanceDiagnostics = DISABLED_ACCEPTANCE_DIAGNOSTICS,
): GateResult {
  if (candidate.sourceIsBridgeTarget !== true) return reject("wrong_source");
  if (!isExactGraphOrigin(candidate.requestUrl)) return reject("wrong_origin");
  const status = candidate.responseStatus;
  if (typeof status !== "number" || !Number.isInteger(status)) return reject("no_response");
  // 401 means Graph rejected the token; a 403 alone does not invalidate it.
  if (status === 401) return reject("http_401");

  const token = candidate.bearer;
  if (
    typeof token !== "string" ||
    token.length === 0 ||
    token.length > MAX_TOKEN_LENGTH ||
    !HEADER_SAFE_RE.test(token)
  ) {
    return reject("unsupported_token_format");
  }
  const claims = tryDecodeClaims(token);
  if (claims === undefined) return reject("unsupported_token_format");

  if (typeof nowMs !== "number" || !Number.isFinite(nowMs) || Math.abs(nowMs) > MAX_DATE_MS) {
    return reject("clock_invalid");
  }

  if (typeof claims.aud !== "string" || !GRAPH_AUDIENCES.has(claims.aud)) return reject("audience");
  if (!guidEquals(claims.tid, config.tenantId)) return reject("tenant");

  const scp = claims.scp;
  if (typeof scp !== "string" || scp.trim().length === 0) return reject("not_delegated");
  // Delegated-only profile: any `roles` claim, even an empty one, is refused (D44: roles absent).
  if (claims.roles !== undefined) return reject("not_delegated");

  if (typeof claims.exp !== "number" || !Number.isFinite(claims.exp)) return reject("lifetime");
  const expMs = claimSecondsToMs(claims.exp);
  if (expMs === undefined || expMs - nowMs < MIN_REMAINING_MS) return reject("lifetime");

  if (claims.nbf !== undefined) {
    if (typeof claims.nbf !== "number" || !Number.isFinite(claims.nbf)) return reject("lifetime");
    const nbfMs = claimSecondsToMs(claims.nbf);
    if (nbfMs === undefined) return reject("lifetime");
    if (nowMs < nbfMs - NBF_SKEW_MS) return reject("not_yet_valid");
  }

  const pin = config.clientPin;
  if (pin !== undefined) {
    let selected: "appid" | "azp" | undefined;
    if (pin.claim === "auto") {
      selected = claims.ver === "1.0" ? "appid" : claims.ver === "2.0" ? "azp" : undefined;
    } else {
      selected = pin.claim;
    }
    if (selected === undefined || !guidEquals(claims[selected], pin.guid)) {
      return reject("client_pin");
    }
  }

  // D48: observation only, on the final accepting return (every check above passed).
  if (diagnostics.enabled) {
    const shape = classifyTokenShape(claims, config.tenantId, expMs - nowMs);
    if (shape !== undefined) diagnostics.tokenShape(shape);
  }
  return { ok: true, bearer: token, expMs, remainingMs: expMs - nowMs };
}
