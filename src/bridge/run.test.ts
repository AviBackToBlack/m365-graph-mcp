import { EventEmitter } from "node:events";
import fs from "node:fs";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import type { ChildProcess } from "node:child_process";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { SpawnFn } from "./browser.js";
import { FakeCdpPeer } from "./fake-cdp.js";
import { BridgeStartError, CDP_CLOSE_GRACE_MS, EXIT, startBridge } from "./run.js";
import type { BridgeHandle } from "./run.js";
import { BRIDGE_DEV_PIPE_ENV, parseBridgeConfig } from "./config.js";
import { PassThrough } from "node:stream";

const TENANT = "11111111-2222-3333-4444-555555555555";
const GE = "https://developer.microsoft.com/graph/graph-explorer/";
const GRAPH = "https://graph.microsoft.com/v1.0/me";

function jwt(over: Record<string, unknown> = {}): string {
  const enc = (o: unknown): string => Buffer.from(JSON.stringify(o)).toString("base64url");
  return `${enc({ alg: "RS256" })}.${enc({
    aud: "https://graph.microsoft.com",
    tid: TENANT,
    scp: "Mail.Read",
    exp: Math.floor(Date.now() / 1000) + 3600,
    ...over,
  })}.c2ln`;
}

function ask(socketPath: string, line: string): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    const s = net.createConnection({ path: socketPath });
    let out = "";
    s.on("data", (d) => (out += d.toString("utf8")));
    s.on("close", () => resolve(JSON.parse(out) as Record<string, unknown>));
    s.on("error", reject);
    s.on("connect", () => s.write(`${line}\n`));
  });
}
const flush = async (): Promise<void> => {
  for (let i = 0; i < 6; i++) await new Promise((r) => setImmediate(r));
};

interface Harness {
  spawnFn: SpawnFn;
  peer: FakeCdpPeer;
  child: EventEmitter & { stdio: unknown[]; kill: ReturnType<typeof vi.fn> };
  spawned: { command: string; args: string[] }[];
}

function harness(): Harness {
  const peer = new FakeCdpPeer();
  peer.on("Target.setDiscoverTargets", () => ({}));
  peer.on("Target.getTargets", () => ({
    targetInfos: [{ targetId: "T1", type: "page", url: GE }],
  }));
  peer.on("Target.attachToTarget", () => ({ sessionId: "S1" }));
  peer.on("Network.enable", () => ({}));
  const child = Object.assign(new EventEmitter(), {
    stdio: [null, null, new PassThrough(), peer.input, peer.output] as unknown[],
    kill: vi.fn(),
  });
  peer.on("Browser.close", () => {
    setImmediate(() => child.emit("exit", 0, null));
    return {};
  });
  const spawned: Harness["spawned"] = [];
  const spawnFn = ((command: string, args: string[]) => {
    spawned.push({ command, args });
    return child as unknown as ChildProcess;
  }) as unknown as SpawnFn;
  return { spawnFn, peer, child, spawned };
}

describe.skipIf(process.platform === "win32")("bridge run (fake browser, real unix socket)", () => {
  let root: string;
  let handle: BridgeHandle | undefined;
  const logs: string[] = [];
  const env = (extra: Record<string, string> = {}) => {
    const browser = path.join(root, "fakechrome");
    fs.writeFileSync(browser, "#!/bin/sh\n", { mode: 0o755 });
    return {
      M365_GRAPH_TENANT_ID: TENANT,
      M365_GRAPH_BRIDGE_DIR: path.join(root, "r"),
      M365_GRAPH_BROWSER_PATH: browser,
      ...extra,
    };
  };

  beforeEach(() => {
    root = fs.mkdtempSync(path.join(os.tmpdir(), "mr-"));
    logs.length = 0;
  });
  afterEach(async () => {
    await handle?.stop();
    handle = undefined;
    fs.rmSync(root, { recursive: true, force: true });
  });

  it("serves a captured token end-to-end over the broker and shuts down cleanly", async () => {
    const h = harness();
    handle = await startBridge({
      env: env(),
      platform: "linux",
      spawnFn: h.spawnFn,
      log: (l) => logs.push(l),
    });
    const sock = path.join(root, "r", "b.sock");
    // Launch contract.
    expect(h.spawned[0]?.args).toContain("--remote-debugging-pipe");
    expect(h.spawned[0]?.args.find((a) => a.startsWith("--user-data-dir="))).toBe(
      `--user-data-dir=${path.join(root, "r", "profile")}`,
    );
    expect((fs.statSync(path.join(root, "r", "profile")).mode & 0o077) === 0).toBe(true);

    // Before a token is observed the broker reports reauth_required.
    expect(await ask(sock, '{"op":"getToken"}')).toMatchObject({ status: "reauth_required" });

    const token = jwt();
    h.peer.event(
      "Network.requestWillBeSent",
      { requestId: "1", request: { url: GRAPH, headers: { Authorization: `Bearer ${token}` } } },
      "S1",
    );
    h.peer.event(
      "Network.responseReceived",
      { requestId: "1", response: { status: 200, url: GRAPH } },
      "S1",
    );
    await flush();
    expect(await ask(sock, '{"op":"getToken"}')).toMatchObject({ status: "ok", token });
    expect(await ask(sock, '{"op":"status"}')).toMatchObject({ state: "ok" });

    // Nothing but the allowlisted methods was written to the pipe.
    const methods = new Set(h.peer.received.map((m) => m.method));
    expect([...methods].sort()).toEqual(
      [
        "Network.enable",
        "Target.attachToTarget",
        "Target.getTargets",
        "Target.setDiscoverTargets",
      ].sort(),
    );

    await handle.stop();
    expect(h.peer.received.at(-1)?.method).toBe("Browser.close");
    await expect(handle.done).resolves.toBe(EXIT.ok);
    expect(fs.existsSync(sock)).toBe(false);
    // Diagnostics are fixed categories: no token content.
    expect(logs.join("\n")).not.toContain(token);
  });

  it("browser exit makes the source unavailable and stops the broker (exit code 6)", async () => {
    const h = harness();
    let t = Date.now();
    handle = await startBridge({
      env: env(),
      platform: "linux",
      spawnFn: h.spawnFn,
      now: () => t,
    });
    t += 60_000; // well past the early-exit window
    h.child.emit("exit", 0, null);
    await expect(handle.done).resolves.toBe(EXIT.browserExited);
    expect(fs.existsSync(path.join(root, "r", "b.sock"))).toBe(false);
  });

  it("an immediate browser exit is reported as profile busy (exit code 4)", async () => {
    const h = harness();
    handle = await startBridge({ env: env(), platform: "linux", spawnFn: h.spawnFn });
    h.child.emit("exit", 0, null);
    await expect(handle.done).resolves.toBe(EXIT.profileBusy);
  });

  it("a broken CDP pipe stops the bridge with exit code 8 and kills the browser", async () => {
    const h = harness();
    h.child.kill.mockImplementation(() => {
      setImmediate(() => h.child.emit("exit", null, "SIGTERM"));
      return true;
    });
    let t = Date.now();
    handle = await startBridge({
      env: env(),
      platform: "linux",
      spawnFn: h.spawnFn,
      now: () => t,
      log: (l) => logs.push(l),
    });
    t += 60_000;
    h.peer.eof();
    await expect(handle.done).resolves.toBe(EXIT.cdpClosed);
    expect(EXIT.cdpClosed).toBe(8);
    expect(h.child.kill).toHaveBeenCalledWith("SIGTERM");
    expect(fs.existsSync(path.join(root, "r", "b.sock"))).toBe(false);
    expect(logs).toContain("cdp pipe closed");
  });

  it("a deliberate stop does not double-trigger the CDP close path", async () => {
    const h = harness();
    handle = await startBridge({ env: env(), platform: "linux", spawnFn: h.spawnFn });
    await handle.stop();
    await expect(handle.done).resolves.toBe(EXIT.ok);
    await new Promise((r) => setTimeout(r, CDP_CLOSE_GRACE_MS + 100));
    await expect(handle.done).resolves.toBe(EXIT.ok);
  });

  it("an asynchronous spawn error is a browser failure (exit 3), not profile busy", async () => {
    const h = harness();
    handle = await startBridge({
      env: env(),
      platform: "linux",
      spawnFn: h.spawnFn,
      log: (l) => logs.push(l),
    });
    h.child.emit("error", new Error("ENOENT SECRET-CANARY")); // immediately: within the early window
    await expect(handle.done).resolves.toBe(EXIT.browser);
    expect(logs.join(String.fromCharCode(10))).not.toContain("SECRET-CANARY");
    expect(fs.existsSync(path.join(root, "r", "b.sock"))).toBe(false);
  });

  it("a signal before the browser is launched stops the broker and never spawns", async () => {
    const h = harness();
    const ac = new AbortController();
    ac.abort();
    handle = await startBridge({
      env: env(),
      platform: "linux",
      spawnFn: h.spawnFn,
      signal: ac.signal,
    });
    await expect(handle.done).resolves.toBe(EXIT.ok);
    expect(h.spawned).toHaveLength(0);
    expect(fs.existsSync(path.join(root, "r", "b.sock"))).toBe(false);
  });

  it("a signal during CDP startup tears down broker and browser and exits clean", async () => {
    const h = harness();
    const ac = new AbortController();
    h.peer.on("Target.getTargets", () => {
      ac.abort(); // SIGINT while startup is still in flight
      return { targetInfos: [{ targetId: "T1", type: "page", url: GE }] };
    });
    handle = await startBridge({
      env: env(),
      platform: "linux",
      spawnFn: h.spawnFn,
      signal: ac.signal,
    });
    await expect(handle.done).resolves.toBe(EXIT.ok);
    expect(h.peer.received.map((m) => m.method)).toContain("Browser.close");
    expect(h.peer.received.map((m) => m.method)).not.toContain("Target.attachToTarget");
    expect(fs.existsSync(path.join(root, "r", "b.sock"))).toBe(false);
    const closes = h.peer.received.filter((m) => m.method === "Browser.close");
    expect(closes).toHaveLength(1); // single shutdown path
  });

  it("a signal after startup shuts the running bridge down cleanly", async () => {
    const h = harness();
    const ac = new AbortController();
    handle = await startBridge({
      env: env(),
      platform: "linux",
      spawnFn: h.spawnFn,
      signal: ac.signal,
    });
    ac.abort();
    await handle.stop(); // concurrent stop shares the same shutdown
    await expect(handle.done).resolves.toBe(EXIT.ok);
    expect(h.peer.received.filter((m) => m.method === "Browser.close")).toHaveLength(1);
    expect(fs.existsSync(path.join(root, "r", "b.sock"))).toBe(false);
  });

  it("refuses to start with a live lock on the profile and does not spawn", async () => {
    const h = harness();
    const e = env();
    const profile = path.join(root, "r", "profile");
    fs.mkdirSync(profile, { recursive: true, mode: 0o700 });
    fs.symlinkSync(`otherhost-${process.ppid}`, path.join(profile, "SingletonLock"));
    const err = await startBridge({ env: e, platform: "linux", spawnFn: h.spawnFn }).catch(
      (x: unknown) => x,
    );
    expect(err).toBeInstanceOf(BridgeStartError);
    expect((err as BridgeStartError).exitCode).toBe(EXIT.profileBusy);
    expect(h.spawned).toHaveLength(0);
    // The lock is left in place.
    expect(fs.lstatSync(path.join(profile, "SingletonLock")).isSymbolicLink()).toBe(true);
  });

  it("a browser that dies during CDP setup is reported by its exit (profile busy), not as a CDP failure", async () => {
    const h = harness();
    h.peer.on("Target.getTargets", (m) => {
      h.child.emit("exit", 0, null); // within the early-exit window
      h.peer.writeFrame({ id: m.id, error: { code: -32000, message: "gone" } });
      return {};
    });
    const err = await startBridge({ env: env(), platform: "linux", spawnFn: h.spawnFn }).catch(
      (x: unknown) => x as BridgeStartError,
    );
    expect(err).toBeInstanceOf(BridgeStartError);
    expect((err as BridgeStartError).exitCode).toBe(EXIT.profileBusy);
    expect(fs.existsSync(path.join(root, "r", "b.sock"))).toBe(false);
  });

  it("fails with browser/config exit codes and fixed categories", async () => {
    const h = harness();
    const missing = await startBridge({
      env: env({ M365_GRAPH_BROWSER_PATH: "/nonexistent/SECRET-CANARY" }),
      platform: "linux",
      spawnFn: h.spawnFn,
    }).catch((x: unknown) => x as BridgeStartError);
    expect((missing as BridgeStartError).exitCode).toBe(EXIT.browser);
    expect((missing as BridgeStartError).message).not.toContain("SECRET-CANARY");
    const noTenant = await startBridge({
      env: { ...env(), M365_GRAPH_TENANT_ID: "bad" },
      platform: "linux",
      spawnFn: h.spawnFn,
    }).catch((x: unknown) => x as BridgeStartError);
    expect((noTenant as BridgeStartError).exitCode).toBe(EXIT.config);
    expect(h.spawned).toHaveLength(0);
  });

  it("does not launch a second browser when a bridge is already running", async () => {
    const h = harness();
    handle = await startBridge({ env: env(), platform: "linux", spawnFn: h.spawnFn });
    const h2 = harness();
    const err = await startBridge({ env: env(), platform: "linux", spawnFn: h2.spawnFn }).catch(
      (x: unknown) => x as BridgeStartError,
    );
    expect((err as BridgeStartError).exitCode).toBe(EXIT.broker);
    expect(h2.spawned).toHaveLength(0);
  });

  describe("acceptance diagnostics (D48)", () => {
    const FIXED_NOW = 1_800_000_000_000;
    const SWITCH = "M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS";
    const CANARY_PAGE = "https://developer.microsoft.com/de-de/graph/graph-explorer?q=PAGEQ7";
    const CANARY_GRAPH = "https://graph.microsoft.com/v1.0/me/messages?$search=GRAPHQ7";

    function fixedJwt(): string {
      const enc = (o: unknown): string => Buffer.from(JSON.stringify(o)).toString("base64url");
      return `${enc({ alg: "RS256" })}.${enc({
        aud: "https://graph.microsoft.com",
        tid: TENANT,
        scp: "Canary.Scope.Read",
        ver: "1.0",
        appid: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
        exp: FIXED_NOW / 1000 + 3600,
      })}.SIGCANARY8M`;
    }

    /** The scripted session: attach, one accepted Graph request, token and status via the broker. */
    async function scenario(extra: Record<string, string>) {
      const h = harness();
      h.peer.on("Target.getTargets", () => ({
        targetInfos: [{ targetId: "TGTCANARY9Z", type: "page", url: CANARY_PAGE }],
      }));
      h.peer.on("Target.attachToTarget", () => ({ sessionId: "SESSCANARY7Q" }));
      const local: string[] = [];
      handle = await startBridge({
        env: env(extra),
        platform: "linux",
        spawnFn: h.spawnFn,
        now: () => FIXED_NOW,
        log: (l) => local.push(l),
      });
      const token = fixedJwt();
      h.peer.event(
        "Network.requestWillBeSent",
        {
          requestId: "REQCANARY3X",
          request: { url: CANARY_GRAPH, headers: { Authorization: `Bearer ${token}` } },
        },
        "SESSCANARY7Q",
      );
      h.peer.event(
        "Network.responseReceived",
        { requestId: "REQCANARY3X", response: { status: 200, url: CANARY_GRAPH } },
        "SESSCANARY7Q",
      );
      await flush();
      const sock = path.join(root, "r", "b.sock");
      const got = await ask(sock, '{"op":"getToken"}');
      const status = await ask(sock, '{"op":"status"}');
      const methods = h.peer.received.map((m) => m.method);
      await handle.stop();
      handle = undefined;
      return { logs: local, got, status, token, methods };
    }

    const ALLOWED_ACCEPTANCE =
      /^acceptance: (mode=enabled|token_shape audience=\w+ client_claim=\w+ token_version=\w+ scp=\w+ roles=\w+ tenant=\w+ remaining_minutes=\d+|graph_request_exact_origin_match=\w+ graph_request_scheme=\w+ graph_request_host=\w+ graph_request_port=\w+|graph_explorer_page_shape=(canonical|localized locale_segment_shape=xx-xx)|authorization_observation=\w+|extra_info_authorization=(present|not_observed|present_after_response))$/;

    it("is silent, with the exact pre-existing log lines, when unset or empty", async () => {
      const expected = [
        "target attached",
        "bridge ready (client pin off)",
        "token accepted (remaining 3600s)",
      ];
      const unset = await scenario({});
      const empty = await scenario({ [SWITCH]: "" });
      expect(unset.logs).toEqual(expected);
      expect(empty.logs).toEqual(expected);
      expect(unset.logs.some((l) => l.includes("acceptance"))).toBe(false);
    });

    it("when exactly 1, adds only allowlisted shape lines and leaves every other line, the broker answers and the CDP traffic unchanged", async () => {
      const off = await scenario({});
      const on = await scenario({ [SWITCH]: "1" });
      const acceptance = on.logs.filter((l) => l.startsWith("acceptance: "));
      expect(on.logs.filter((l) => !l.startsWith("acceptance: "))).toEqual(off.logs);
      expect(acceptance.length).toBeGreaterThan(3);
      for (const l of acceptance) expect(l).toMatch(ALLOWED_ACCEPTANCE);
      // The mode marker is announced before capture starts, so it is always the first line.
      expect(acceptance[0]).toBe("acceptance: mode=enabled");
      expect(acceptance).toContain(
        "acceptance: token_shape audience=graph_url client_claim=appid token_version=v1 scp=present roles=absent tenant=matched remaining_minutes=60",
      );
      expect(acceptance).toContain(
        "acceptance: graph_explorer_page_shape=localized locale_segment_shape=xx-xx",
      );
      expect(acceptance).toContain("acceptance: authorization_observation=main_event");

      // Nothing through the broker: identical responses and no acceptance field anywhere.
      expect(on.got).toEqual(off.got);
      expect(on.status).toEqual(off.status);
      expect(JSON.stringify([on.got, on.status])).not.toContain("acceptance");
      expect(on.got).toMatchObject({ status: "ok", token: on.token });
      // No extra CDP command is issued by the diagnostics.
      expect(on.methods).toEqual(off.methods);

      // Canary material never reaches stderr text, in any form.
      const dump = on.logs.join("\n");
      for (const needle of [
        on.token,
        ...on.token.split("."),
        TENANT,
        "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
        "Canary.Scope.Read",
        "SESSCANARY7Q",
        "TGTCANARY9Z",
        "REQCANARY3X",
        "PAGEQ7",
        "GRAPHQ7",
        "de-de",
        "microsoft.com",
        "Bearer",
      ]) {
        expect(dump, needle).not.toContain(needle);
      }
    });
  });
});

describe("dev pipe config", () => {
  const base = { M365_GRAPH_TENANT_ID: TENANT, M365_GRAPH_BRIDGE_DIR: "/tmp/mb" };
  it("is honored only on win32 and only for a plain pipe name", () => {
    const pipe = "\\\\.\\pipe\\m365-bridge-test";
    const c = parseBridgeConfig({
      env: { ...base, [BRIDGE_DEV_PIPE_ENV]: pipe },
      platform: "win32",
    });
    expect(c.socketPath).toBe(pipe);
    expect(c.allowDevPipe).toBe(true);
    for (const bad of [
      "\\\\.\\pipe\\a\\b",
      "\\\\.\\pipe\\",
      "\\\\.\\pipe\\x y",
      "\\\\server\\pipe\\x",
      "127.0.0.1:8080",
      "/tmp/x.sock",
    ]) {
      expect(() =>
        parseBridgeConfig({ env: { ...base, [BRIDGE_DEV_PIPE_ENV]: bad }, platform: "win32" }),
      ).toThrow();
    }
    expect(() =>
      parseBridgeConfig({
        env: { ...base, M365_GRAPH_BRIDGE_DIR: "/tmp/mb", [BRIDGE_DEV_PIPE_ENV]: pipe },
        platform: "darwin",
      }),
    ).toThrow();
    expect(() =>
      parseBridgeConfig({
        env: { ...base, M365_GRAPH_BRIDGE_DIR: "/tmp/mb", [BRIDGE_DEV_PIPE_ENV]: pipe },
        platform: "linux",
      }),
    ).toThrow();
  });

  it("win32 without the dev pipe still has no usable endpoint", () => {
    expect(() => parseBridgeConfig({ env: base, platform: "win32" })).toThrow();
  });
});

describe("acceptance diagnostics enablement at startup (D48)", () => {
  it.each(["true", "yes", "01", " 1", "1 ", " ", "0", "CANARY-VALUE-7Q"])(
    "exits with the config code for %j without echoing the value",
    async (value) => {
      const h = harness();
      const err = await startBridge({
        env: {
          M365_GRAPH_TENANT_ID: TENANT,
          M365_GRAPH_BRIDGE_DIR: "/tmp/mb",
          M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS: value,
        },
        platform: "linux",
        spawnFn: h.spawnFn,
      }).catch((x: unknown) => x);
      expect(err).toBeInstanceOf(BridgeStartError);
      expect((err as BridgeStartError).exitCode).toBe(EXIT.config);
      expect((err as BridgeStartError).message).toContain(
        "M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS",
      );
      if (value.trim().length > 1) expect((err as BridgeStartError).message).not.toContain(value);
      expect(h.spawned).toHaveLength(0);
    },
  );
});
