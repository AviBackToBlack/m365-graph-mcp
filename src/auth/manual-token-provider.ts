import { TokenError } from "./token-provider.js";
import type { TokenProvider } from "./token-provider.js";
import {
  GRAPH_AUDIENCES,
  GUID_RE,
  HEADER_SAFE_RE,
  MAX_DATE_MS,
  MAX_TOKEN_LENGTH,
  claimSecondsToMs,
  tryDecodeClaims,
} from "./jwt-claims.js";

export { MAX_TOKEN_LENGTH };

/** Environment variable holding the short-lived bearer token (never argv). */
export const ACCESS_TOKEN_ENV_VAR = "M365_GRAPH_ACCESS_TOKEN";
/** Optional environment variable pinning the expected tenant GUID (`tid`). */
export const TENANT_ID_ENV_VAR = "M365_TENANT_ID";

const BEARER_PREFIX_RE = /^bearer(?:\s+|$)/i;

/** Tokens within this window of `exp` are treated as expired (avoid mid-request expiry). */
export const DEFAULT_EXPIRY_BUFFER_MS = 30_000;
/** Tolerance for `nbf` being slightly in the future due to clock drift. */
export const DEFAULT_CLOCK_SKEW_MS = 60_000;

/** Upper bound for the configurable expiry buffer and clock skew. */
const MAX_TIME_OPTION_MS = 3_600_000;

function validateTimeOption(value: number | undefined, fallback: number): number {
  const v = value ?? fallback;
  if (typeof v !== "number" || !Number.isFinite(v) || v < 0 || v > MAX_TIME_OPTION_MS) {
    throw new TokenError(
      "auth_misconfigured",
      "expiryBufferMs and clockSkewMs must be finite, non-negative numbers of at most one hour.",
    );
  }
  return v;
}

export interface ManualTokenProviderOptions {
  /** Environment to read from; defaults to `process.env`. */
  env?: Readonly<Record<string, string | undefined>>;
  /** Clock in epoch milliseconds; defaults to `Date.now`. */
  now?: () => number;
  /** Tenant pin; overrides `M365_TENANT_ID` when given. */
  tenantId?: string;
  expiryBufferMs?: number;
  clockSkewMs?: number;
}

function normalizeTenantPin(value: string | undefined): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) return undefined;
  if (!GUID_RE.test(trimmed)) {
    throw new TokenError(
      "auth_misconfigured",
      `${TENANT_ID_ENV_VAR} must be a tenant GUID (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).`,
    );
  }
  return trimmed.toLowerCase();
}

/**
 * Bootstrap provider: a manually supplied short-lived bearer read from an
 * environment variable. The environment is re-read on every call. When the
 * token is a decodable JWT, advisory claim checks (audience, expiry, nbf,
 * tenant pin) run so an expired token is not handed out; opaque tokens are
 * passed through (unless a tenant pin is set) and Graph stays authoritative.
 */
export class ManualTokenProvider implements TokenProvider {
  private readonly env: Readonly<Record<string, string | undefined>>;
  private readonly now: () => number;
  private readonly tenantPin: string | undefined;
  private readonly expiryBufferMs: number;
  private readonly clockSkewMs: number;

  /** @throws TokenError `auth_misconfigured` if the tenant pin is not a GUID or a time option is invalid. */
  constructor(options: ManualTokenProviderOptions = {}) {
    this.env = options.env ?? process.env;
    this.now = options.now ?? Date.now;
    this.tenantPin = normalizeTenantPin(options.tenantId ?? this.env[TENANT_ID_ENV_VAR]);
    this.expiryBufferMs = validateTimeOption(options.expiryBufferMs, DEFAULT_EXPIRY_BUFFER_MS);
    this.clockSkewMs = validateTimeOption(options.clockSkewMs, DEFAULT_CLOCK_SKEW_MS);
  }

  getAccessToken(): Promise<string> {
    try {
      return Promise.resolve(this.readValidToken());
    } catch (error) {
      return Promise.reject(
        error instanceof TokenError ? error : new Error("Token validation failed."),
      );
    }
  }

  private readValidToken(): string {
    const raw = this.env[ACCESS_TOKEN_ENV_VAR];
    const token = raw?.trim().replace(BEARER_PREFIX_RE, "").trim();
    if (!token) {
      throw new TokenError(
        "reauth_required",
        `No access token available: ${ACCESS_TOKEN_ENV_VAR} is not set. Supply a fresh Graph access token and restart.`,
      );
    }

    if (token.length > MAX_TOKEN_LENGTH) {
      throw new TokenError(
        "reauth_required",
        `${ACCESS_TOKEN_ENV_VAR} exceeds the maximum supported token length; supply a valid Graph access token.`,
      );
    }
    if (!HEADER_SAFE_RE.test(token)) {
      throw new TokenError(
        "reauth_required",
        `${ACCESS_TOKEN_ENV_VAR} contains whitespace, control, or non-ASCII characters and cannot be used as a bearer token; supply a valid Graph access token.`,
      );
    }

    const claims = tryDecodeClaims(token);
    if (claims === undefined) {
      // Opaque (non-JWT, JWE, or undecodable payload): advisory checks cannot run.
      if (this.tenantPin !== undefined) {
        throw new TokenError(
          "reauth_required",
          "Access token is opaque (not a decodable JWT), so the configured tenant pin cannot be verified; supply a decodable Graph access token or remove the tenant pin.",
        );
      }
      return token; // Graph is authoritative; a 401 later surfaces as reauth.
    }
    const nowMs = this.now();
    // A broken injected clock must not make temporal checks pass vacuously.
    if (typeof nowMs !== "number" || !Number.isFinite(nowMs) || Math.abs(nowMs) > MAX_DATE_MS) {
      throw new TokenError(
        "auth_misconfigured",
        "Clock returned an invalid time; cannot validate access token expiry.",
      );
    }

    if (typeof claims.aud !== "string" || !GRAPH_AUDIENCES.has(claims.aud)) {
      throw new TokenError(
        "reauth_required",
        "Access token audience is not Microsoft Graph; supply a token issued for https://graph.microsoft.com.",
      );
    }

    if (typeof claims.exp !== "number" || !Number.isFinite(claims.exp)) {
      throw new TokenError(
        "reauth_required",
        "Access token has no valid expiry (exp) claim; supply a fresh Graph access token.",
      );
    }
    const expMs = claimSecondsToMs(claims.exp);
    if (expMs === undefined) {
      throw new TokenError(
        "reauth_required",
        "Access token has an out-of-range expiry (exp) claim; supply a fresh Graph access token.",
      );
    }
    if (nowMs >= expMs - this.expiryBufferMs) {
      const expDate = new Date(expMs);
      const when = Number.isFinite(expDate.getTime()) ? expDate.toISOString() : "an invalid time";
      throw new TokenError(
        "reauth_required",
        `Access token expired or expiring at ${when}; supply a fresh token and restart/re-auth.`,
      );
    }

    if (claims.nbf !== undefined) {
      if (typeof claims.nbf !== "number" || !Number.isFinite(claims.nbf)) {
        throw new TokenError(
          "reauth_required",
          "Access token has an invalid not-before (nbf) claim; supply a fresh Graph access token.",
        );
      }
      const nbfMs = claimSecondsToMs(claims.nbf);
      if (nbfMs === undefined) {
        throw new TokenError(
          "reauth_required",
          "Access token has an out-of-range not-before (nbf) claim; supply a fresh Graph access token.",
        );
      }
      if (nowMs < nbfMs - this.clockSkewMs) {
        throw new TokenError(
          "reauth_required",
          "Access token is not valid yet (nbf is in the future); check the system clock or supply a fresh token.",
        );
      }
    }

    if (this.tenantPin !== undefined) {
      if (typeof claims.tid !== "string" || claims.tid.toLowerCase() !== this.tenantPin) {
        throw new TokenError(
          "reauth_required",
          "Access token tenant (tid) does not match the configured tenant; sign in to the configured tenant and supply a fresh token.",
        );
      }
    }

    return token;
  }
}
