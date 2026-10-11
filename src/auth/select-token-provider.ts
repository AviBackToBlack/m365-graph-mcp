import { BridgeConfigError, resolveBridgeEndpoint } from "../bridge/config.js";
import { BridgeTokenProvider } from "./bridge-token-provider.js";
import { ManualTokenProvider } from "./manual-token-provider.js";
import { TokenError } from "./token-provider.js";
import type { TokenProvider } from "./token-provider.js";

/** Selects the MCP server's `TokenProvider` (D45). */
export const TOKEN_SOURCE_ENV = "M365_GRAPH_TOKEN_SOURCE";

export interface SelectTokenProviderOptions {
  /** Environment to read from; defaults to `process.env`. */
  env?: Readonly<Record<string, string | undefined>>;
  platform?: NodeJS.Platform;
  homedir?: string;
}

/**
 * Chooses the TokenProvider from `M365_GRAPH_TOKEN_SOURCE` only (D45). Selection is explicit:
 * - unset: `manual` (the compatibility default);
 * - exactly `manual` or `bridge`: that provider;
 * - anything else, including an empty or whitespace-only value and case/whitespace variants such
 *   as `Bridge` or ` bridge`: a startup configuration error. No normalization is applied, so an
 *   operator who templated an empty or mistyped value never silently runs through the other
 *   provider.
 *
 * There is no auto-detection and no fallback in either direction. `manual` never reads bridge
 * configuration or contacts the bridge; `bridge` never reads a manual bearer and neither requires
 * nor reads any tenant variable (the bridge process owns tenant binding, D44/D46). A bridge that
 * is not running is not a startup failure: token requests yield `reauth_required`.
 *
 * @throws TokenError `auth_misconfigured` on an invalid selector or invalid endpoint/manual config.
 *   Messages name the variable and never echo its value.
 */
export function selectTokenProvider(options: SelectTokenProviderOptions = {}): TokenProvider {
  const env = options.env ?? process.env;
  const source = env[TOKEN_SOURCE_ENV];

  if (source === undefined || source === "manual") {
    return new ManualTokenProvider({ env });
  }
  if (source === "bridge") {
    try {
      const endpoint = resolveBridgeEndpoint({
        env,
        ...(options.platform !== undefined ? { platform: options.platform } : {}),
        ...(options.homedir !== undefined ? { homedir: options.homedir } : {}),
      });
      return new BridgeTokenProvider({
        socketPath: endpoint.socketPath,
        ...(endpoint.allowDevPipe ? { allowDevPipe: true } : {}),
        ...(options.platform !== undefined ? { platform: options.platform } : {}),
      });
    } catch (error) {
      if (error instanceof BridgeConfigError) {
        throw new TokenError(
          "auth_misconfigured",
          `${TOKEN_SOURCE_ENV}=bridge but the bridge endpoint configuration is invalid: ${error.message}`,
        );
      }
      throw error;
    }
  }
  throw new TokenError(
    "auth_misconfigured",
    `${TOKEN_SOURCE_ENV} must be exactly "manual" or "bridge" (or unset for "manual").`,
  );
}
