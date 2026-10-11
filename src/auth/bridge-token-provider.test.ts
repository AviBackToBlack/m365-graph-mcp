import fs from "node:fs";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { BridgeBroker, InMemoryTokenSource } from "../bridge/broker.js";
import { BridgeTokenProvider, MAX_RESPONSE_BYTES } from "./bridge-token-provider.js";
import { TokenError } from "./token-provider.js";

const TOKEN = "TOKEN-CANARY.payload.sig";
const NOW = 1_800_000_000_000;

describe.skipIf(process.platform === "win32")("BridgeTokenProvider", () => {
  let root: string;
  let sock: string;
  let fake: net.Server | undefined;
  const conns = new Set<net.Socket>();
  let broker: BridgeBroker | undefined;

  beforeEach(() => {
    root = fs.mkdtempSync(path.join(os.tmpdir(), "mp-"));
    sock = path.join(root, "b.sock");
  });
  afterEach(async () => {
    await broker?.stop();
    broker = undefined;
    await closeFake();
    fs.rmSync(root, { recursive: true, force: true });
  });

  async function closeFake(): Promise<void> {
    for (const c of conns) c.destroy();
    conns.clear();
    const f = fake;
    fake = undefined;
    if (f) await new Promise<void>((r) => f.close(() => r()));
  }

  async function fakeServer(onConn: (s: net.Socket) => void): Promise<void> {
    fake = net.createServer((s) => {
      conns.add(s);
      onConn(s);
    });
    await new Promise<void>((r) => fake!.listen(sock, r));
  }

  async function failure(provider: BridgeTokenProvider): Promise<TokenError> {
    try {
      await provider.getAccessToken();
    } catch (e) {
      expect(e).toBeInstanceOf(TokenError);
      return e as TokenError;
    }
    throw new Error("expected failure");
  }

  it("returns the token from a real broker", async () => {
    const source = new InMemoryTokenSource();
    source.set({ state: "ok", token: TOKEN, expMs: NOW + 600_000 });
    // Broker requires its own 0700 dir.
    const dir = path.join(root, "d");
    sock = path.join(dir, "b.sock");
    broker = new BridgeBroker({ socketPath: sock, source, now: () => NOW });
    await broker.start();
    expect(await new BridgeTokenProvider({ socketPath: sock }).getAccessToken()).toBe(TOKEN);
  });

  it.each(["reauth_required", "unavailable", "unsupported_token_format"])(
    "maps broker state %s to reauth_required without leaking",
    async (state) => {
      await fakeServer((s) => s.end(`${JSON.stringify({ status: state, reason: "tenant" })}\n`));
      const err = await failure(new BridgeTokenProvider({ socketPath: sock }));
      expect(err.code).toBe("reauth_required");
      expect(err.message).toContain(state);
      expect(err.message).not.toContain("tenant");
    },
  );

  it("maps a missing socket and a refused connection to reauth_required", async () => {
    const err = await failure(new BridgeTokenProvider({ socketPath: sock }));
    expect(err.code).toBe("reauth_required");
    expect(err.message).toContain("connect_failed");
  });

  it("does not echo server-supplied status text or the token on bad responses", async () => {
    for (const reply of [
      `${JSON.stringify({ status: `evil-${TOKEN}`, token: TOKEN })}\n`,
      `not json ${TOKEN}\n`,
      `${JSON.stringify({ status: "ok" })}\n`,
      `${JSON.stringify({ status: "ok", token: "has space" })}\n`,
      `${JSON.stringify({ status: "ok", token: 5 })}\n`,
      `[1]\n`,
      "",
    ]) {
      await fakeServer((s) => s.end(reply));
      const err = await failure(new BridgeTokenProvider({ socketPath: sock }));
      expect(err.code).toBe("reauth_required");
      expect(err.message).not.toContain(TOKEN);
      expect(JSON.stringify(err)).not.toContain(TOKEN);
      await closeFake();
      fs.rmSync(sock, { force: true });
    }
  });

  it("rejects oversized responses", async () => {
    await fakeServer((s) => s.write("a".repeat(MAX_RESPONSE_BYTES + 100)));
    const err = await failure(new BridgeTokenProvider({ socketPath: sock }));
    expect(err.message).toContain("response_too_large");
  });

  it("times out on a silent server", async () => {
    await fakeServer(() => undefined);
    const err = await failure(new BridgeTokenProvider({ socketPath: sock, timeoutMs: 50 }));
    expect(err.code).toBe("reauth_required");
    expect(err.message).toContain("timeout");
  });

  it("refuses non-path endpoints at construction", () => {
    for (const p of ["8080", "127.0.0.1:80", "rel.sock"]) {
      expect(() => new BridgeTokenProvider({ socketPath: p })).toThrow(TokenError);
    }
  });

  it("never connects to a socket in a group/other-accessible directory (unsafe_endpoint)", async () => {
    const dir = path.join(root, "open");
    fs.mkdirSync(dir);
    fs.chmodSync(dir, 0o755);
    sock = path.join(dir, "b.sock");
    let connected = false;
    await fakeServer((s) => {
      connected = true;
      s.end(`${JSON.stringify({ status: "ok", token: TOKEN, expiresInSeconds: 600 })}\n`);
    });
    const err = await failure(new BridgeTokenProvider({ socketPath: sock }));
    expect(err.code).toBe("reauth_required");
    expect(err.message).toContain("unsafe_endpoint");
    expect(err.message).not.toContain(TOKEN);
    expect(connected).toBe(false);
  });

  it("refuses a non-socket file at the socket path (unsafe_endpoint)", async () => {
    fs.writeFileSync(sock, "x");
    const err = await failure(new BridgeTokenProvider({ socketPath: sock }));
    expect(err.message).toContain("unsafe_endpoint");
  });

  it("refuses a socket whose directory has an unsafe ancestor (unsafe_endpoint)", async () => {
    const parent = path.join(root, "w");
    const dir = path.join(parent, "d");
    fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
    fs.chmodSync(dir, 0o700);
    fs.chmodSync(parent, 0o777);
    sock = path.join(dir, "b.sock");
    await fakeServer((s) => s.end("{}\n"));
    const err = await failure(new BridgeTokenProvider({ socketPath: sock }));
    expect(err.message).toContain("unsafe_endpoint");
  });
});
