import path from "node:path";
import { describe, expect, it } from "vitest";
import { BridgeConfigError, parseBridgeConfig, resolveBridgeEndpoint } from "./config.js";
import { assertIpcPath, IpcPathError, MAX_SOCKET_PATH_BYTES } from "./ipc-path.js";

const TENANT = "11111111-2222-3333-4444-555555555555";
const CLIENT = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
const base = { M365_GRAPH_TENANT_ID: TENANT, M365_GRAPH_BRIDGE_DIR: "/tmp/mb" };

describe("bridge config", () => {
  it("refuses to start without a valid tenant GUID", () => {
    for (const t of [undefined, "", "  ", "contoso.com", "not-a-guid", `${TENANT}x`]) {
      expect(() =>
        parseBridgeConfig({ env: { ...base, M365_GRAPH_TENANT_ID: t }, platform: "linux" }),
      ).toThrow(BridgeConfigError);
    }
  });

  it("does not echo bad values in errors", () => {
    expect(() =>
      parseBridgeConfig({
        env: { ...base, M365_GRAPH_TENANT_ID: "SECRET-CANARY" },
        platform: "linux",
      }),
    ).toThrow(/^(?!.*SECRET-CANARY).*$/s);
  });

  it("parses tenant, optional pin and browser path", () => {
    const c = parseBridgeConfig({
      env: {
        ...base,
        M365_GRAPH_TENANT_ID: TENANT.toUpperCase(),
        M365_GRAPH_CLIENT_ID_PIN: CLIENT,
        M365_GRAPH_CLIENT_ID_PIN_CLAIM: "azp",
        M365_GRAPH_BROWSER_PATH: "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
      },
      platform: "linux",
    });
    expect(c.gate).toEqual({ tenantId: TENANT, clientPin: { guid: CLIENT, claim: "azp" } });
    expect(c.browserPath).toContain("Chrome");
    expect(parseBridgeConfig({ env: base, platform: "linux" }).gate.clientPin).toBeUndefined();
  });

  it("rejects bad pin, pin claim, and relative browser or dir paths", () => {
    const bad = [
      { M365_GRAPH_CLIENT_ID_PIN: "nope" },
      { M365_GRAPH_CLIENT_ID_PIN: CLIENT, M365_GRAPH_CLIENT_ID_PIN_CLAIM: "oid" },
      { M365_GRAPH_CLIENT_ID_PIN_CLAIM: "azp" },
      { M365_GRAPH_BROWSER_PATH: "chrome" },
      { M365_GRAPH_BRIDGE_DIR: "relative/dir" },
    ];
    for (const extra of bad) {
      expect(() => parseBridgeConfig({ env: { ...base, ...extra }, platform: "linux" })).toThrow(
        BridgeConfigError,
      );
    }
  });

  it("uses the macOS Application Support default and requires an override elsewhere", () => {
    const c = parseBridgeConfig({
      env: { M365_GRAPH_TENANT_ID: TENANT },
      platform: "darwin",
      homedir: "/Users/someone",
    });
    expect(c.runtimeDir).toBe(
      path.posix.join("/Users/someone", "Library", "Application Support", "m365-bridge"),
    );
    expect(c.socketPath.endsWith("b.sock")).toBe(true);
    expect(Buffer.byteLength(c.socketPath)).toBeLessThanOrEqual(MAX_SOCKET_PATH_BYTES);
    expect(() =>
      parseBridgeConfig({ env: { M365_GRAPH_TENANT_ID: TENANT }, platform: "linux" }),
    ).toThrow(BridgeConfigError);
  });

  it("fails explicitly when the socket path exceeds the sun_path limit", () => {
    expect(MAX_SOCKET_PATH_BYTES).toBe(103);
    const dir = "/" + "d".repeat(MAX_SOCKET_PATH_BYTES - "/b.sock".length - 1);
    expect(Buffer.byteLength(`${dir}/b.sock`)).toBe(MAX_SOCKET_PATH_BYTES);
    expect(
      parseBridgeConfig({ env: { ...base, M365_GRAPH_BRIDGE_DIR: dir }, platform: "linux" })
        .socketPath,
    ).toBe(`${dir}/b.sock`);
    expect(() =>
      parseBridgeConfig({ env: { ...base, M365_GRAPH_BRIDGE_DIR: `${dir}d` }, platform: "linux" }),
    ).toThrow(/exceeds 103 bytes/);
  });

  it("long macOS home directories fail rather than truncate", () => {
    expect(() =>
      parseBridgeConfig({
        env: { M365_GRAPH_TENANT_ID: TENANT },
        platform: "darwin",
        homedir: "/Users/" + "u".repeat(80),
      }),
    ).toThrow(BridgeConfigError);
  });
});

describe("assertIpcPath", () => {
  it.each([3000, "3000", "127.0.0.1:8080", "relative.sock", "", {}, undefined, "\\\\.\\pipe\\x"])(
    "refuses non-path endpoint %j",
    (v) => {
      expect(() => assertIpcPath(v, { platform: "linux" })).toThrow(IpcPathError);
    },
  );

  it("permits dev pipes only on win32 behind the explicit flag", () => {
    const pipe = "\\\\.\\pipe\\m365-dev";
    expect(() => assertIpcPath(pipe, { platform: "win32" })).toThrow(IpcPathError);
    expect(() => assertIpcPath(pipe, { platform: "linux", allowDevPipe: true })).toThrow(
      IpcPathError,
    );
    expect(() => assertIpcPath(pipe, { platform: "win32", allowDevPipe: true })).not.toThrow();
  });

  it("refuses every non-dev-pipe endpoint on win32, where listen(path) binds a named pipe", () => {
    for (const p of ["C:\\bridge\\b.sock", "/tmp/b.sock", "\\\\.\\pipe\\x"]) {
      expect(() => assertIpcPath(p, { platform: "win32" }), p).toThrow(IpcPathError);
      if (!p.startsWith("\\\\.\\pipe\\")) {
        expect(() => assertIpcPath(p, { platform: "win32", allowDevPipe: true }), p).toThrow(
          IpcPathError,
        );
      }
    }
  });
});

describe("resolveBridgeEndpoint (shared with the MCP bridge token source, D45)", () => {
  it("does not require M365_GRAPH_TENANT_ID or read any tenant variable", () => {
    const e = resolveBridgeEndpoint({
      env: { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", M365_TENANT_ID: "ignored" },
      platform: "linux",
    });
    expect(e.runtimeDir).toBe("/tmp/mb");
    expect(e.socketPath).toBe(path.posix.join("/tmp/mb", "b.sock"));
    expect(e.allowDevPipe).toBeUndefined();
  });

  it("is the endpoint parseBridgeConfig returns, and the bridge still requires its own tenant", () => {
    const env = { ...base };
    const cfg = parseBridgeConfig({ env, platform: "linux" });
    expect(resolveBridgeEndpoint({ env, platform: "linux" })).toEqual({
      runtimeDir: cfg.runtimeDir,
      socketPath: cfg.socketPath,
    });
    // M365_TENANT_ID (the manual-mode pin) is never an alias for M365_GRAPH_TENANT_ID.
    expect(() =>
      parseBridgeConfig({
        env: { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", M365_TENANT_ID: TENANT },
        platform: "linux",
      }),
    ).toThrow(BridgeConfigError);
  });

  it("applies the same macOS default, dev-pipe and path rules", () => {
    expect(
      resolveBridgeEndpoint({ env: {}, platform: "darwin", homedir: "/Users/someone" }).runtimeDir,
    ).toBe(path.posix.join("/Users/someone", "Library", "Application Support", "m365-bridge"));
    expect(() => resolveBridgeEndpoint({ env: {}, platform: "linux" })).toThrow(BridgeConfigError);
    expect(() =>
      resolveBridgeEndpoint({
        env: { M365_GRAPH_BRIDGE_DEV_PIPE: "\\\\.\\pipe\\x", M365_GRAPH_BRIDGE_DIR: "/tmp/mb" },
        platform: "darwin",
      }),
    ).toThrow(BridgeConfigError);
    const pipe = resolveBridgeEndpoint({
      env: { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", M365_GRAPH_BRIDGE_DEV_PIPE: "\\\\.\\pipe\\x" },
      platform: "win32",
    });
    expect(pipe).toMatchObject({ socketPath: "\\\\.\\pipe\\x", allowDevPipe: true });
  });
});

describe("acceptance diagnostics enablement (D48)", () => {
  const VAR = "M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS";
  const parse = (extra: Record<string, string | undefined>) =>
    parseBridgeConfig({ env: { ...base, ...extra }, platform: "linux" });

  it("is off when unset or exactly empty, and adds no field to the config", () => {
    expect(parse({}).acceptanceDiagnostics).toBeUndefined();
    expect(parse({ [VAR]: "" }).acceptanceDiagnostics).toBeUndefined();
    expect(parse({ [VAR]: undefined })).toEqual(parse({}));
    expect("acceptanceDiagnostics" in parse({ [VAR]: "" })).toBe(false);
  });

  it("is on only for exactly 1", () => {
    expect(parse({ [VAR]: "1" }).acceptanceDiagnostics).toBe(true);
  });

  it.each(["true", "TRUE", "yes", "on", "0", "01", "11", " 1", "1 ", "1\n", " ", "\t", "2", "-1"])(
    "rejects %j as a configuration error without echoing it",
    (value) => {
      let message = "";
      try {
        parse({ [VAR]: value });
      } catch (e) {
        expect(e).toBeInstanceOf(BridgeConfigError);
        message = (e as Error).message;
      }
      // The fixed message names only the variable; it never embeds the offending value.
      expect(message).toBe(`${VAR} must be unset, empty, or exactly 1.`);
    },
  );

  it("does not echo an arbitrary canary value", () => {
    expect(() => parse({ [VAR]: "CANARY-VALUE-7Q" })).toThrow(/^(?!.*CANARY-VALUE-7Q).*$/s);
  });

  it("is not read by the shared endpoint resolution used by the MCP side", () => {
    const reads: string[] = [];
    const env = new Proxy({ M365_GRAPH_BRIDGE_DIR: "/tmp/mb" } as Record<string, string>, {
      get(t, p) {
        if (typeof p === "string") reads.push(p);
        return t[p as string];
      },
    });
    resolveBridgeEndpoint({ env, platform: "linux" });
    expect(reads).not.toContain(VAR);
    // An invalid value cannot break endpoint resolution (the MCP server never sees this switch).
    expect(() =>
      resolveBridgeEndpoint({
        env: { M365_GRAPH_BRIDGE_DIR: "/tmp/mb", [VAR]: "CANARY" },
        platform: "linux",
      }),
    ).not.toThrow();
  });
});
