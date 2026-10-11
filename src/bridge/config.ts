import os from "node:os";
import path from "node:path";
import { GUID_RE } from "../auth/jwt-claims.js";
import { assertIpcPath, DEV_PIPE_PREFIX, IpcPathError } from "./ipc-path.js";
import type { TokenGateConfig } from "./token-gate.js";

export const TENANT_ID_ENV = "M365_GRAPH_TENANT_ID";
export const CLIENT_PIN_ENV = "M365_GRAPH_CLIENT_ID_PIN";
export const CLIENT_PIN_CLAIM_ENV = "M365_GRAPH_CLIENT_ID_PIN_CLAIM";
export const BROWSER_PATH_ENV = "M365_GRAPH_BROWSER_PATH";
export const BRIDGE_DIR_ENV = "M365_GRAPH_BRIDGE_DIR";
/** Windows development/test only: the broker listens on this `\\.\pipe\<name>` instead of a socket. */
export const BRIDGE_DEV_PIPE_ENV = "M365_GRAPH_BRIDGE_DEV_PIPE";
/** SM-15 acceptance diagnostics switch (D48). Read only by the bridge, never by the MCP server. */
export const ACCEPTANCE_DIAGNOSTICS_ENV = "M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS";
const DEV_PIPE_NAME_RE = /^[A-Za-z0-9._-]{1,64}$/;

/** Short app directory name: keeps the macOS socket path well below sun_path. */
export const APP_DIR_NAME = "m365-bridge";
export const SOCKET_FILE_NAME = "b.sock";

export class BridgeConfigError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "BridgeConfigError";
  }
}

export interface BridgeConfig {
  gate: TokenGateConfig;
  /** Optional absolute browser executable path (consumed by SM-15 part B). */
  browserPath?: string;
  runtimeDir: string;
  socketPath: string;
  /** True only when `socketPath` is a validated Windows dev named pipe (win32 development only). */
  allowDevPipe?: boolean;
  /** Present (true) only when SM-15 acceptance diagnostics were explicitly enabled (D48). */
  acceptanceDiagnostics?: true;
}

export interface BridgeConfigOptions {
  env?: Readonly<Record<string, string | undefined>>;
  platform?: NodeJS.Platform;
  homedir?: string;
}

function clean(v: string | undefined): string | undefined {
  const t = v?.trim();
  return t ? t : undefined;
}

/**
 * D48 enablement: unset or exactly "" is off, exactly "1" is on, anything else is a configuration
 * error. Deliberately NOT passed through `clean()`: no trimming, coercion or fuzzy booleans. The
 * message names only the variable and never echoes the value.
 */
function parseAcceptanceDiagnostics(raw: string | undefined): boolean {
  if (raw === undefined || raw === "") return false;
  if (raw === "1") return true;
  throw new BridgeConfigError(`${ACCEPTANCE_DIAGNOSTICS_ENV} must be unset, empty, or exactly 1.`);
}

/** Broker endpoint shared by the bridge process and the MCP-side `BridgeTokenProvider`. */
export interface BridgeEndpoint {
  runtimeDir: string;
  socketPath: string;
  /** True only when `socketPath` is a validated Windows dev named pipe (win32 development only). */
  allowDevPipe?: boolean;
}

/**
 * Resolves the broker endpoint (runtime directory and socket path) from `M365_GRAPH_BRIDGE_DIR`,
 * else the macOS default, with the Windows-only `M365_GRAPH_BRIDGE_DEV_PIPE` override. Never TCP.
 * This is the single implementation used by both `m365-graph-bridge` (via `parseBridgeConfig`) and
 * the MCP server's bridge token source (D45). It reads only the two endpoint env vars and
 * deliberately neither requires nor reads `M365_GRAPH_TENANT_ID`. Messages never echo values.
 * @throws BridgeConfigError on any invalid endpoint setting.
 */
export function resolveBridgeEndpoint(options: BridgeConfigOptions = {}): BridgeEndpoint {
  const env = options.env ?? process.env;
  const platform = options.platform ?? process.platform;
  const home = options.homedir ?? os.homedir();

  const override = clean(env[BRIDGE_DIR_ENV]);
  let runtimeDir: string;
  if (override !== undefined) {
    if (!path.isAbsolute(override)) {
      throw new BridgeConfigError(`${BRIDGE_DIR_ENV} must be an absolute path.`);
    }
    runtimeDir = override;
  } else if (platform === "darwin") {
    runtimeDir = path.join(home, "Library", "Application Support", APP_DIR_NAME);
  } else {
    throw new BridgeConfigError(`${BRIDGE_DIR_ENV} is required on this platform.`);
  }

  let socketPath = path.join(runtimeDir, SOCKET_FILE_NAME);
  const devPipe = clean(env[BRIDGE_DEV_PIPE_ENV]);
  if (devPipe !== undefined) {
    if (platform !== "win32") {
      throw new BridgeConfigError(`${BRIDGE_DEV_PIPE_ENV} is honored only on Windows (dev/test).`);
    }
    if (
      !devPipe.startsWith(DEV_PIPE_PREFIX) ||
      !DEV_PIPE_NAME_RE.test(devPipe.slice(DEV_PIPE_PREFIX.length))
    ) {
      throw new BridgeConfigError(
        `${BRIDGE_DEV_PIPE_ENV} must be a \\\\.\\pipe\\<name> pipe name.`,
      );
    }
    socketPath = devPipe;
  }
  try {
    assertIpcPath(socketPath, { platform, allowDevPipe: devPipe !== undefined });
  } catch (error) {
    if (error instanceof IpcPathError) throw new BridgeConfigError(error.message);
    throw error;
  }

  return {
    runtimeDir,
    socketPath,
    ...(devPipe !== undefined ? { allowDevPipe: true } : {}),
  };
}

/**
 * Parses bridge configuration. Messages name only the env var, never its value.
 * @throws BridgeConfigError on any missing/invalid setting (refuse to start).
 */
export function parseBridgeConfig(options: BridgeConfigOptions = {}): BridgeConfig {
  const env = options.env ?? process.env;

  const tenant = clean(env[TENANT_ID_ENV]);
  if (tenant === undefined || !GUID_RE.test(tenant)) {
    throw new BridgeConfigError(`${TENANT_ID_ENV} is required and must be a tenant GUID.`);
  }

  let clientPin: TokenGateConfig["clientPin"];
  const pinGuid = clean(env[CLIENT_PIN_ENV]);
  const pinClaim = clean(env[CLIENT_PIN_CLAIM_ENV]);
  if (pinGuid !== undefined) {
    if (!GUID_RE.test(pinGuid)) {
      throw new BridgeConfigError(`${CLIENT_PIN_ENV} must be a client GUID when set.`);
    }
    const claim = pinClaim ?? "auto";
    if (claim !== "auto" && claim !== "appid" && claim !== "azp") {
      throw new BridgeConfigError(`${CLIENT_PIN_CLAIM_ENV} must be one of auto, appid, azp.`);
    }
    clientPin = { guid: pinGuid.toLowerCase(), claim };
  } else if (pinClaim !== undefined) {
    throw new BridgeConfigError(`${CLIENT_PIN_CLAIM_ENV} requires ${CLIENT_PIN_ENV}.`);
  }

  const browserPath = clean(env[BROWSER_PATH_ENV]);
  if (browserPath !== undefined && (!path.isAbsolute(browserPath) || browserPath.includes("\0"))) {
    throw new BridgeConfigError(`${BROWSER_PATH_ENV} must be an absolute path.`);
  }

  const acceptanceDiagnostics = parseAcceptanceDiagnostics(env[ACCEPTANCE_DIAGNOSTICS_ENV]);

  const endpoint = resolveBridgeEndpoint(options);

  return {
    gate: { tenantId: tenant.toLowerCase(), ...(clientPin ? { clientPin } : {}) },
    ...(browserPath !== undefined ? { browserPath } : {}),
    ...endpoint,
    ...(acceptanceDiagnostics ? { acceptanceDiagnostics: true as const } : {}),
  };
}
