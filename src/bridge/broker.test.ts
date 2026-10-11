import { spawn } from "node:child_process";
import fs from "node:fs";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { BridgeBroker, BrokerError, InMemoryTokenSource, MAX_REQUEST_BYTES } from "./broker.js";
import { IpcPathError } from "./ipc-path.js";

const TOKEN = "TOKEN-CANARY.payload.sig";
const NOW = 1_800_000_000_000;

function ask(socketPath: string, raw: string, endAfter = true): Promise<string> {
  return new Promise((resolve, reject) => {
    const s = net.createConnection({ path: socketPath });
    let out = "";
    s.on("data", (d) => (out += d.toString("utf8")));
    s.on("close", () => resolve(out));
    s.on("error", reject);
    s.on("connect", () => {
      s.write(raw);
      if (!endAfter) return;
    });
  });
}

describe.skipIf(process.platform === "win32")("bridge broker (unix domain socket)", () => {
  let root: string;
  let dir: string;
  let sock: string;
  let source: InMemoryTokenSource;
  let broker: BridgeBroker | undefined;

  const make = (over: Partial<ConstructorParameters<typeof BridgeBroker>[0]> = {}) =>
    new BridgeBroker({ socketPath: sock, source, now: () => NOW, ...over });

  beforeEach(() => {
    root = fs.mkdtempSync(path.join(os.tmpdir(), "mb-"));
    dir = path.join(root, "r");
    sock = path.join(dir, "b.sock");
    source = new InMemoryTokenSource(true);
  });
  afterEach(async () => {
    await broker?.stop();
    broker = undefined;
    vi.restoreAllMocks();
    fs.rmSync(root, { recursive: true, force: true });
  });

  it("creates the runtime dir 0700 and the socket 0600", async () => {
    broker = make();
    await broker.start();
    expect(fs.lstatSync(dir).mode & 0o777).toBe(0o700);
    const st = fs.lstatSync(sock);
    expect(st.isSocket()).toBe(true);
    expect(st.mode & 0o777).toBe(0o600);
  });

  it("refuses a group/world-accessible runtime dir", async () => {
    fs.mkdirSync(dir, { mode: 0o755 });
    fs.chmodSync(dir, 0o755);
    broker = make();
    await expect(broker.start()).rejects.toBeInstanceOf(BrokerError);
    expect(fs.existsSync(sock)).toBe(false);
    fs.chmodSync(dir, 0o770);
    await expect(make().start()).rejects.toThrow(/group\/others/);
  });

  it("refuses a symlinked runtime dir", async () => {
    const real = path.join(root, "real");
    fs.mkdirSync(real, { mode: 0o700 });
    fs.symlinkSync(real, dir);
    await expect(make().start()).rejects.toThrow(/plain directory/);
  });

  it("refuses a symlinked socket path and a non-socket file at the path", async () => {
    fs.mkdirSync(dir, { mode: 0o700 });
    const target = path.join(root, "target");
    fs.writeFileSync(target, "x");
    fs.symlinkSync(target, sock);
    await expect(make().start()).rejects.toThrow(/not a plain socket/);
    fs.unlinkSync(sock);
    fs.writeFileSync(sock, "x");
    await expect(make().start()).rejects.toThrow(/not a plain socket/);
    expect(fs.readFileSync(sock, "utf8")).toBe("x");
  });

  it("removes a stale socket and starts", async () => {
    fs.mkdirSync(dir, { mode: 0o700 });
    // A killed owner leaves its socket file behind (a graceful close would unlink it).
    const child = spawn(
      process.execPath,
      [
        "-e",
        'require("net").createServer().listen(process.argv[1],()=>console.log("ready"))',
        sock,
      ],
      { stdio: ["ignore", "pipe", "ignore"] },
    );
    await new Promise<void>((r) => child.stdout.once("data", () => r()));
    child.kill("SIGKILL");
    await new Promise<void>((r) => child.once("exit", () => r()));
    expect(fs.lstatSync(sock).isSocket()).toBe(true);
    broker = make();
    await broker.start();
    source.set({ state: "unavailable" });
    expect(JSON.parse(await ask(sock, '{"op":"status"}\n')).status).toBe("ok");
  });

  it("refuses to replace a live bridge", async () => {
    broker = make();
    await broker.start();
    const second = make();
    await expect(second.start()).rejects.toMatchObject({ code: "already_running" });
    // The original keeps working.
    expect(JSON.parse(await ask(sock, '{"op":"status"}\n')).status).toBe("ok");
  });

  it("serves getToken with remaining lifetime", async () => {
    source.set({ state: "ok", token: TOKEN, expMs: NOW + 600_000 });
    broker = make();
    await broker.start();
    expect(JSON.parse(await ask(sock, '{"op":"getToken"}\n'))).toEqual({
      status: "ok",
      token: TOKEN,
      expiresInSeconds: 600,
    });
  });

  it("does not serve a token that is too close to expiry", async () => {
    source.set({ state: "ok", token: TOKEN, expMs: NOW + 10_000 });
    broker = make();
    await broker.start();
    const out = await ask(sock, '{"op":"getToken"}\n');
    expect(out).not.toContain(TOKEN);
    expect(JSON.parse(out)).toEqual({ status: "reauth_required", reason: "lifetime" });
  });

  it.each(["reauth_required", "unavailable", "unsupported_token_format"] as const)(
    "reports %s without token material",
    async (state) => {
      source.set({ state });
      broker = make();
      await broker.start();
      const out = JSON.parse(await ask(sock, '{"op":"getToken"}\n'));
      expect(out.status).toBe(state);
      expect(out.token).toBeUndefined();
    },
  );

  it("status is token-free metadata", async () => {
    source.set({ state: "ok", token: TOKEN, expMs: NOW + 600_000 });
    broker = make();
    await broker.start();
    const raw = await ask(sock, '{"op":"status"}\n');
    expect(raw).not.toContain(TOKEN);
    expect(JSON.parse(raw)).toEqual({
      status: "ok",
      state: "ok",
      remainingSeconds: 600,
      clientPinEnabled: true,
    });
  });

  it("only exposes the defined operations", async () => {
    source.set({ state: "ok", token: TOKEN, expMs: NOW + 600_000 });
    broker = make();
    await broker.start();
    for (const req of [
      '{"op":"browser"}',
      '{"op":"cdp"}',
      '{"op":"GETTOKEN"}',
      '{"op":"__proto__"}',
    ]) {
      const out = await ask(sock, `${req}\n`);
      expect(JSON.parse(out)).toEqual({ status: "error", reason: "unknown_op" });
      expect(out).not.toContain(TOKEN);
    }
    for (const req of ["garbage", "[]", "null", '{"op":1}', '{"op":"status","x":1}', "{}"]) {
      expect(JSON.parse(await ask(sock, `${req}\n`))).toEqual({
        status: "error",
        reason: "bad_request",
      });
    }
  });

  it("rejects oversized requests", async () => {
    source.set({ state: "ok", token: TOKEN, expMs: NOW + 600_000 });
    broker = make();
    await broker.start();
    const out = await ask(sock, `{"op":"${"a".repeat(MAX_REQUEST_BYTES + 10)}"}\n`);
    expect(JSON.parse(out)).toEqual({ status: "error", reason: "request_too_large" });
    const noNewline = await ask(sock, "a".repeat(MAX_REQUEST_BYTES + 10));
    expect(JSON.parse(noNewline).reason).toBe("request_too_large");
  });

  it("closes idle connections after the timeout", async () => {
    broker = make({ connectionTimeoutMs: 50 });
    await broker.start();
    const out = await ask(sock, '{"op":"stat'); // never completes the line
    expect(out).toBe("");
  });

  it("listens on a filesystem path and never on TCP", async () => {
    const spy = vi.spyOn(net.Server.prototype, "listen");
    broker = make();
    await broker.start();
    const calls = spy.mock.calls.filter((c) => typeof c[0] !== "undefined");
    expect(calls).toHaveLength(1);
    expect(calls[0]?.[0]).toBe(sock);
    expect(typeof calls[0]?.[0]).toBe("string");
    const server = (broker as unknown as { server: net.Server }).server;
    expect(server.address()).toBe(sock); // a string address means a UDS, not {port}
  });

  it.each([8080, "8080", "127.0.0.1:8080", "relative.sock"])(
    "refuses non-path endpoint %j before listening",
    async (endpoint) => {
      const spy = vi.spyOn(net.Server.prototype, "listen");
      const b = new BridgeBroker({ socketPath: endpoint as string, source });
      await expect(b.start()).rejects.toBeInstanceOf(IpcPathError);
      expect(spy).not.toHaveBeenCalled();
    },
  );

  it("refuses the Windows dev pipe without the flag, and on non-win32 even with it", async () => {
    const pipe = "\\\\.\\pipe\\m365-dev";
    await expect(new BridgeBroker({ socketPath: pipe, source }).start()).rejects.toBeInstanceOf(
      IpcPathError,
    );
    await expect(
      new BridgeBroker({ socketPath: pipe, source, allowDevPipe: true }).start(),
    ).rejects.toBeInstanceOf(IpcPathError);
  });

  it("removes the socket on stop", async () => {
    broker = make();
    await broker.start();
    await broker.stop();
    expect(fs.existsSync(sock)).toBe(false);
  });

  it("re-applies the five-minute floor at serve time as an accepted token ages", async () => {
    let now = NOW;
    source.set({ state: "ok", token: TOKEN, expMs: NOW + 6 * 60_000 });
    broker = make({ now: () => now });
    await broker.start();
    expect(JSON.parse(await ask(sock, '{"op":"getToken"}\n')).status).toBe("ok");
    now += 2 * 60_000; // 4 minutes left: below the D44 serving floor
    const out = await ask(sock, '{"op":"getToken"}\n');
    expect(out).not.toContain(TOKEN);
    expect(JSON.parse(out)).toEqual({ status: "reauth_required", reason: "lifetime" });
    now = NOW + 60_000; // exactly 5 minutes left is still served
    expect(JSON.parse(await ask(sock, '{"op":"getToken"}\n')).status).toBe("ok");
  });

  it("never buffers more than the request limit, even with a huge suffix in one chunk", async () => {
    source.set({ state: "ok", token: TOKEN, expMs: NOW + 600_000 });
    broker = make();
    await broker.start();
    const concat = vi.spyOn(Buffer, "concat");
    const out = await ask(sock, `{"op":"status"}\n${"x".repeat(4 * 1024 * 1024)}`);
    expect(JSON.parse(out).status).toBe("ok");
    for (const call of concat.mock.calls) {
      const total = (call[0] as Buffer[]).reduce((n, b) => n + b.length, 0);
      expect(total).toBeLessThanOrEqual(MAX_REQUEST_BYTES);
    }
  });

  it("stop() never unlinks a socket that is no longer the one it bound", async () => {
    broker = make();
    await broker.start();
    const realLstat = fs.lstatSync;
    const foreign = { ...realLstat(sock), isSocket: () => true, ino: realLstat(sock).ino + 1 };
    vi.spyOn(fs, "lstatSync").mockImplementation(((p: fs.PathLike, o?: unknown) =>
      p === sock ? foreign : realLstat(p, o as never)) as typeof fs.lstatSync);
    const unlink = vi.spyOn(fs, "unlinkSync");
    await broker.stop();
    broker = undefined;
    expect(unlink.mock.calls.some(([p]) => p === sock)).toBe(false);
  });

  it.each([0o777, 0o775])(
    "refuses a runtime dir whose parent is group/other-writable without sticky (%o)",
    async (mode) => {
      const parent = path.join(root, "p");
      fs.mkdirSync(parent);
      fs.chmodSync(parent, mode);
      const d = path.join(parent, "r");
      broker = make({ socketPath: path.join(d, "b.sock") });
      await expect(broker.start()).rejects.toBeInstanceOf(BrokerError);
      broker = undefined;
    },
  );

  it("accepts a sticky world-writable parent such as /tmp", async () => {
    const parent = path.join(root, "t");
    fs.mkdirSync(parent);
    fs.chmodSync(parent, 0o1777);
    broker = make({ socketPath: path.join(parent, "r", "b.sock") });
    await broker.start();
  });

  it("a throwing token source yields a fixed unavailable response and the bridge keeps serving", async () => {
    let fail = true;
    const throwing = {
      clientPinEnabled: false,
      current: () => {
        if (fail) throw new Error(`boom ${TOKEN}`);
        return { state: "ok" as const, token: TOKEN, expMs: NOW + 600_000 };
      },
    };
    broker = make({ source: throwing });
    await broker.start();
    const out = await ask(sock, '{"op":"getToken"}\n');
    expect(JSON.parse(out)).toEqual({ status: "unavailable" });
    expect(out).not.toContain("boom");
    fail = false;
    expect(JSON.parse(await ask(sock, '{"op":"getToken"}\n')).status).toBe("ok");
  });

  it.each([
    ["a BigInt token", { state: "ok", token: 1n, expMs: NOW + 600_000 }],
    ["an object token", { state: "ok", token: { t: TOKEN }, expMs: NOW + 600_000 }],
    ["a string expiry", { state: "ok", token: TOKEN, expMs: String(NOW + 600_000) }],
  ])("a source returning %s is reported unavailable and never served", async (_n, snap) => {
    const bad = { clientPinEnabled: false, current: () => snap as never };
    broker = make({ source: bad });
    await broker.start();
    const out = await ask(sock, '{"op":"getToken"}\n');
    expect(JSON.parse(out)).toEqual({ status: "unavailable" });
    expect(out).not.toContain(TOKEN);
    expect(JSON.parse(await ask(sock, '{"op":"status"}\n')).state).toBe("unavailable");
  });
});
