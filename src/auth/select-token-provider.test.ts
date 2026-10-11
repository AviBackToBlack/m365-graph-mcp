import fs from "node:fs";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { BridgeBroker, InMemoryTokenSource } from "../bridge/broker.js";
import { parseBridgeConfig, resolveBridgeEndpoint } from "../bridge/config.js";
import { BridgeTokenProvider } from "./bridge-token-provider.js";
import { ManualTokenProvider } from "./manual-token-provider.js";
import { selectTokenProvider, TOKEN_SOURCE_ENV } from "./select-token-provider.js";
import { TokenError } from "./token-provider.js";
import type { TokenProvider } from "./token-provider.js";

const TENANT_A = "11111111-2222-3333-4444-555555555555";
const TENANT_B = "99999999-8888-7777-6666-555555555555";
// Header-safe, JWT-shaped but not decodable (opaque to the manual provider).
const OPAQUE_TOKEN = "TOKEN-CANARY-opaque-bearer-value";
const PIPE = "\\\\.\\pipe\\m365-select-test";

type Env = Record<string, string | undefined>;

/** Env whose reads are recorded, to pin which variables each mode may look at. */
function recordingEnv(values: Env): { env: Env; reads: Set<string> } {
  const reads = new Set<string>();
  const env = new Proxy(values, {
    get(target, prop) {
      if (typeof prop === "string") reads.add(prop);
      return target[prop as string];
    },
  });
  return { env, reads };
}

function startupError(fn: () => unknown): TokenError {
  try {
    fn();
  } catch (e) {
    expect(e).toBeInstanceOf(TokenError);
    return e as TokenError;
  }
  throw new Error("expected startup failure");
}

async function tokenFailure(provider: TokenProvider): Promise<TokenError> {
  try {
    await provider.getAccessToken();
  } catch (e) {
    expect(e).toBeInstanceOf(TokenError);
    return e as TokenError;
  }
  throw new Error("expected failure");
}

/** A bridge-mode env/platform pair with a valid but unused endpoint, on the real host platform. */
function unavailableBridge(): { env: Env; platform: NodeJS.Platform } {
  if (process.platform === "win32") {
    return {
      env: {
        [TOKEN_SOURCE_ENV]: "bridge",
        // Host-absolute dir: path.isAbsolute follows the host, not the simulated platform.
        M365_GRAPH_BRIDGE_DIR: os.tmpdir(),
        M365_GRAPH_BRIDGE_DEV_PIPE: `${PIPE}-absent`,
      },
      platform: "win32",
    };
  }
  const dir = path.join(os.tmpdir(), `m365-select-absent-${process.pid}`);
  return {
    env: { [TOKEN_SOURCE_ENV]: "bridge", M365_GRAPH_BRIDGE_DIR: dir },
    platform: process.platform,
  };
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("selectTokenProvider: selection (D45)", () => {
  it("defaults to manual when unset", () => {
    expect(selectTokenProvider({ env: {} })).toBeInstanceOf(ManualTokenProvider);
  });

  it("selects manual and bridge only on the exact values", () => {
    expect(selectTokenProvider({ env: { [TOKEN_SOURCE_ENV]: "manual" } })).toBeInstanceOf(
      ManualTokenProvider,
    );
    expect(
      selectTokenProvider({
        env: { [TOKEN_SOURCE_ENV]: "bridge", M365_GRAPH_BRIDGE_DIR: "/tmp/mb" },
        platform: "linux",
      }),
    ).toBeInstanceOf(BridgeTokenProvider);
  });

  it.each([
    "",
    " ",
    "Bridge",
    "BRIDGE",
    "Manual",
    " bridge",
    "bridge ",
    "manual\n",
    "auto",
    "both",
    "none",
    "true",
    "1",
    "SECRET-CANARY",
  ])("refuses to start for %j and names only the variable", (value) => {
    const err = startupError(() =>
      selectTokenProvider({
        env: { [TOKEN_SOURCE_ENV]: value, M365_GRAPH_BRIDGE_DIR: "/tmp/mb" },
        platform: "linux",
      }),
    );
    expect(err.code).toBe("auth_misconfigured");
    expect(err.message).toContain(TOKEN_SOURCE_ENV);
    if (value.trim()) expect(err.message).not.toContain(value);
  });

  it("never infers the provider from ambient state", () => {
    // A manual token, a bridge dir and a dev pipe being present must not change an unset selector.
    const env = { M365_GRAPH_BRIDGE_DIR: "/tmp/mb" };
    expect(selectTokenProvider({ env, platform: "darwin" })).toBeInstanceOf(ManualTokenProvider);
    const withToken = {
      ...env,
      M365_GRAPH_ACCESS_TOKEN: OPAQUE_TOKEN,
      [TOKEN_SOURCE_ENV]: "bridge",
    };
    expect(selectTokenProvider({ env: withToken, platform: "linux" })).toBeInstanceOf(
      BridgeTokenProvider,
    );
  });
});

describe("selectTokenProvider: bridge-only diagnostics switch (D48)", () => {
  it("never reads M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS in either mode, valid or not", () => {
    const sw = "M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS";
    const manual = recordingEnv({ [sw]: "CANARY-not-a-valid-value" });
    expect(selectTokenProvider({ env: manual.env, platform: "linux" })).toBeInstanceOf(
      ManualTokenProvider,
    );
    expect(manual.reads.has(sw)).toBe(false);
    const bridge = recordingEnv({
      [TOKEN_SOURCE_ENV]: "bridge",
      M365_GRAPH_BRIDGE_DIR: "/tmp/mb",
      [sw]: "1",
    });
    expect(selectTokenProvider({ env: bridge.env, platform: "linux" })).toBeInstanceOf(
      BridgeTokenProvider,
    );
    expect(bridge.reads.has(sw)).toBe(false);
  });
});

describe("selectTokenProvider: manual mode", () => {
  it("reads no bridge configuration and never contacts the bridge", async () => {
    const connect = vi.spyOn(net, "createConnection");
    const connectAlias = vi.spyOn(net, "connect");
    const { env, reads } = recordingEnv({
      // Invalid bridge settings must be inert in manual mode.
      M365_GRAPH_BRIDGE_DIR: "relative/dir",
      M365_GRAPH_BRIDGE_DEV_PIPE: "not-a-pipe",
      M365_GRAPH_TENANT_ID: "SECRET-CANARY-not-a-guid",
      M365_GRAPH_ACCESS_TOKEN: OPAQUE_TOKEN,
    });
    const provider = selectTokenProvider({ env, platform: "linux" });
    expect(provider).toBeInstanceOf(ManualTokenProvider);
    expect(await provider.getAccessToken()).toBe(OPAQUE_TOKEN);
    expect(connect).not.toHaveBeenCalled();
    expect(connectAlias).not.toHaveBeenCalled();
    expect([...reads].sort()).toEqual(
      [TOKEN_SOURCE_ENV, "M365_GRAPH_ACCESS_TOKEN", "M365_TENANT_ID"].sort(),
    );
  });

  it("keeps the manual tenant pin behavior and rejects an invalid M365_TENANT_ID", async () => {
    const err = startupError(() =>
      selectTokenProvider({ env: { [TOKEN_SOURCE_ENV]: "manual", M365_TENANT_ID: "nope" } }),
    );
    expect(err.code).toBe("auth_misconfigured");
    // A pin is in effect: an opaque token cannot satisfy it (existing ManualTokenProvider rule).
    const pinned = selectTokenProvider({
      env: { M365_GRAPH_ACCESS_TOKEN: OPAQUE_TOKEN, M365_TENANT_ID: TENANT_A },
    });
    expect((await tokenFailure(pinned)).code).toBe("reauth_required");
  });

  it("ignores M365_GRAPH_TENANT_ID entirely, with no cross-provider conflict logic", async () => {
    const unpinned = selectTokenProvider({
      env: { M365_GRAPH_ACCESS_TOKEN: OPAQUE_TOKEN, M365_GRAPH_TENANT_ID: TENANT_A },
    });
    expect(await unpinned.getAccessToken()).toBe(OPAQUE_TOKEN); // not treated as a pin
    // Both set and different: no aliasing/precedence/conflict; M365_TENANT_ID alone governs.
    const both = selectTokenProvider({
      env: {
        M365_GRAPH_ACCESS_TOKEN: OPAQUE_TOKEN,
        M365_TENANT_ID: TENANT_A,
        M365_GRAPH_TENANT_ID: TENANT_B,
      },
    });
    expect((await tokenFailure(both)).message).toContain("tenant pin");
    // An invalid M365_GRAPH_TENANT_ID is not validated by manual mode.
    expect(() =>
      selectTokenProvider({ env: { M365_GRAPH_TENANT_ID: "not-a-guid" } }),
    ).not.toThrow();
  });
});

describe("selectTokenProvider: bridge mode", () => {
  it("starts without any tenant variable and interprets none of them", () => {
    const { env, reads } = recordingEnv({
      [TOKEN_SOURCE_ENV]: "bridge",
      M365_GRAPH_BRIDGE_DIR: "/tmp/mb",
      // Would each be a startup error (or a pin) if read.
      M365_TENANT_ID: "not-a-guid",
      M365_GRAPH_TENANT_ID: "not-a-guid-either",
    });
    expect(selectTokenProvider({ env, platform: "linux" })).toBeInstanceOf(BridgeTokenProvider);
    expect([...reads].sort()).toEqual(
      [TOKEN_SOURCE_ENV, "M365_GRAPH_BRIDGE_DIR", "M365_GRAPH_BRIDGE_DEV_PIPE"].sort(),
    );
  });

  it("starts with a stopped bridge and answers reauth_required, never the manual bearer", async () => {
    const { env, platform } = unavailableBridge();
    const provider = selectTokenProvider({
      env: { ...env, M365_GRAPH_ACCESS_TOKEN: OPAQUE_TOKEN },
      platform,
    });
    expect(provider).toBeInstanceOf(BridgeTokenProvider);
    const err = await tokenFailure(provider);
    expect(err.code).toBe("reauth_required");
    expect(err.message).not.toContain(OPAQUE_TOKEN);
  });

  it("does not read M365_GRAPH_ACCESS_TOKEN at all", async () => {
    const { env: base, platform } = unavailableBridge();
    const { env, reads } = recordingEnv({ ...base, M365_GRAPH_ACCESS_TOKEN: OPAQUE_TOKEN });
    await tokenFailure(selectTokenProvider({ env, platform }));
    expect(reads.has("M365_GRAPH_ACCESS_TOKEN")).toBe(false);
  });

  it.each([
    ["relative directory", { M365_GRAPH_BRIDGE_DIR: "relative/dir" }, "linux"],
    ["no directory off macOS", {}, "linux"],
    ["no directory on win32 without a pipe", {}, "win32"],
    [
      "dev pipe on linux",
      { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", M365_GRAPH_BRIDGE_DEV_PIPE: PIPE },
      "linux",
    ],
    [
      "dev pipe on darwin",
      { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", M365_GRAPH_BRIDGE_DEV_PIPE: PIPE },
      "darwin",
    ],
    [
      "bad dev pipe name on win32",
      { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", M365_GRAPH_BRIDGE_DEV_PIPE: "\\\\.\\pipe\\a\\b" },
      "win32",
    ],
    [
      "TCP-looking dev pipe on win32",
      { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", M365_GRAPH_BRIDGE_DEV_PIPE: "127.0.0.1:8080" },
      "win32",
    ],
    ["over-long socket path", { M365_GRAPH_BRIDGE_DIR: `/${"d".repeat(120)}` }, "linux"],
  ] as const)("is a startup failure for %s", (_name, extra, platform) => {
    const err = startupError(() =>
      selectTokenProvider({ env: { [TOKEN_SOURCE_ENV]: "bridge", ...extra }, platform }),
    );
    expect(err.code).toBe("auth_misconfigured");
    expect(err.message).toContain(TOKEN_SOURCE_ENV);
  });

  it("does not echo endpoint values in startup errors", () => {
    const err = startupError(() =>
      selectTokenProvider({
        env: { [TOKEN_SOURCE_ENV]: "bridge", M365_GRAPH_BRIDGE_DIR: "SECRET-CANARY/rel" },
        platform: "linux",
      }),
    );
    expect(err.message).not.toContain("SECRET-CANARY");
  });

  it("accepts the Windows dev pipe only on win32", () => {
    const env = {
      [TOKEN_SOURCE_ENV]: "bridge",
      M365_GRAPH_BRIDGE_DIR: "/tmp/mb",
      M365_GRAPH_BRIDGE_DEV_PIPE: PIPE,
    };
    expect(selectTokenProvider({ env, platform: "win32" })).toBeInstanceOf(BridgeTokenProvider);
  });

  describe.skipIf(process.platform === "win32")("with a running broker", () => {
    it("serves a token end to end through the shared endpoint resolution", async () => {
      const root = fs.mkdtempSync(path.join(os.tmpdir(), "sel-"));
      const dir = path.join(root, "d");
      const source = new InMemoryTokenSource();
      const now = 1_800_000_000_000;
      source.set({ state: "ok", token: "BRIDGE-TOKEN-CANARY.p.s", expMs: now + 600_000 });
      const broker = new BridgeBroker({
        socketPath: path.join(dir, "b.sock"),
        source,
        now: () => now,
      });
      try {
        await broker.start();
        const provider = selectTokenProvider({
          env: { [TOKEN_SOURCE_ENV]: "bridge", M365_GRAPH_BRIDGE_DIR: dir },
        });
        expect(await provider.getAccessToken()).toBe("BRIDGE-TOKEN-CANARY.p.s");
      } finally {
        await broker.stop();
        fs.rmSync(root, { recursive: true, force: true });
      }
    });
  });
});

describe("endpoint resolution is shared (D45)", () => {
  const matrix: ReadonlyArray<{ env: Env; platform: NodeJS.Platform; homedir?: string }> = [
    { env: {}, platform: "darwin", homedir: "/Users/someone" },
    { env: { M365_GRAPH_BRIDGE_DIR: "/tmp/mb" }, platform: "darwin", homedir: "/Users/someone" },
    { env: { M365_GRAPH_BRIDGE_DIR: "  /tmp/mb  " }, platform: "linux" },
    { env: { M365_GRAPH_BRIDGE_DIR: "/tmp/mb" }, platform: "linux" },
    {
      env: { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", M365_GRAPH_BRIDGE_DEV_PIPE: PIPE },
      platform: "win32",
    },
    // The dev pipe alone is rejected everywhere: the runtime dir is still required.
    { env: { M365_GRAPH_BRIDGE_DEV_PIPE: PIPE }, platform: "win32" },
    {
      env: { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", M365_GRAPH_BRIDGE_DEV_PIPE: PIPE },
      platform: "darwin",
    },
    { env: { M365_GRAPH_BRIDGE_DIR: "relative" }, platform: "linux" },
    { env: {}, platform: "linux" },
    { env: {}, platform: "win32" },
  ];

  it.each(matrix.map((m, i) => [i, m] as const))(
    "bridge config, endpoint resolver and MCP provider agree (case %i)",
    (_i, input) => {
      const { env, platform, homedir } = input;
      const opts = { env, platform, ...(homedir !== undefined ? { homedir } : {}) };
      const bridgeCfg = (() => {
        try {
          return parseBridgeConfig({ ...opts, env: { ...env, M365_GRAPH_TENANT_ID: TENANT_A } });
        } catch {
          return undefined;
        }
      })();
      const endpoint = (() => {
        try {
          return resolveBridgeEndpoint(opts);
        } catch {
          return undefined;
        }
      })();
      const provider = (() => {
        try {
          return selectTokenProvider({
            ...opts,
            env: { ...env, [TOKEN_SOURCE_ENV]: "bridge" },
          });
        } catch {
          return undefined;
        }
      })();
      // Same accept/reject decision everywhere.
      expect(endpoint === undefined).toBe(bridgeCfg === undefined);
      expect(provider === undefined).toBe(bridgeCfg === undefined);
      if (bridgeCfg && endpoint && provider) {
        expect(endpoint.socketPath).toBe(bridgeCfg.socketPath);
        expect(endpoint.runtimeDir).toBe(bridgeCfg.runtimeDir);
        expect(endpoint.allowDevPipe).toBe(bridgeCfg.allowDevPipe);
        expect((provider as unknown as { socketPath: string }).socketPath).toBe(
          bridgeCfg.socketPath,
        );
      }
    },
  );
});
