/**
 * Supplies a short-lived delegated Microsoft Graph access token.
 *
 * Implementations MUST NOT log the token and MUST NOT include any part of it
 * in errors. Callers must call `getAccessToken()` per request rather than
 * caching the returned value.
 */
export interface TokenProvider {
  getAccessToken(): Promise<string>;
}

/**
 * `reauth_required`: the token is missing, expired, not yet valid, unusable as a
 * header value, or fails an advisory check; the user must supply a fresh one.
 * `auth_misconfigured`: static configuration (for example a tenant pin) is invalid.
 */
export type TokenErrorCode = "reauth_required" | "auth_misconfigured";

/**
 * Token failure. Messages are written by this module from fixed text plus
 * non-secret values (timestamps); they never embed token material or raw
 * input. No `cause` is ever attached, so no underlying parser error (which may
 * quote input) can leak.
 */
export class TokenError extends Error {
  readonly code: TokenErrorCode;

  constructor(code: TokenErrorCode, message: string) {
    super(message);
    this.name = "TokenError";
    this.code = code;
  }

  toJSON(): { name: string; code: TokenErrorCode; message: string } {
    return { name: this.name, code: this.code, message: this.message };
  }
}
