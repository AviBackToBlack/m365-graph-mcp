import { afterEach, describe, expect, it, vi } from "vitest";
import { CDP_ALLOWED_METHODS, CdpError, CdpPipe, MAX_IN_FLIGHT_COMMANDS } from "./cdp-pipe.js";
import type { CdpCloseReason, CdpEvent } from "./cdp-pipe.js";
import { FakeCdpPeer } from "./fake-cdp.js";

const flush = (): Promise<void> => new Promise((r) => setImmediate(r));

function setup(opts: { max?: number; timeout?: number } = {}) {
  const peer = new FakeCdpPeer();
  const pipe = new CdpPipe({
    input: peer.input,
    output: peer.output,
    ...(opts.max ? { maxMessageBytes: opts.max } : {}),
    ...(opts.timeout ? { commandTimeoutMs: opts.timeout } : {}),
  });
  const events: CdpEvent[] = [];
  pipe.onEvent((e) => events.push(e));
  const closes: CdpCloseReason[] = [];
  pipe.onClose((r) => closes.push(r));
  return { peer, pipe, events, closes };
}

afterEach(() => vi.useRealTimers());

describe("cdp pipe: allowlist", () => {
  it("allows exactly the fixed method set", () => {
    expect([...CDP_ALLOWED_METHODS].sort()).toEqual(
      [
        "Browser.close",
        "Network.enable",
        "Target.attachToTarget",
        "Target.detachFromTarget",
        "Target.getTargets",
        "Target.setDiscoverTargets",
      ].sort(),
    );
  });

  it("rejects forbidden methods locally and writes nothing", async () => {
    const { peer, pipe } = setup();
    for (const m of [
      "Network.getCookies",
      "Network.getAllCookies",
      "Storage.getCookies",
      "Runtime.evaluate",
      "DOM.getDocument",
      "Page.navigate",
      "Network.getResponseBody",
      "Target.createTarget",
      "Target.sendMessageToTarget",
      "Browser.getVersion",
      "network.enable",
      "",
      "__proto__",
    ]) {
      await expect(pipe.send(m, {}, "s")).rejects.toMatchObject({ code: "forbidden_method" });
    }
    await flush();
    expect(peer.rawBytes).toBe(0);
  });

  it("enforces session scoping of allowed methods", async () => {
    const { peer, pipe } = setup();
    await expect(pipe.send("Network.enable")).rejects.toMatchObject({ code: "bad_arguments" });
    await expect(pipe.send("Network.enable", { x: 1 }, "s")).rejects.toMatchObject({
      code: "bad_arguments",
    });
    await expect(pipe.send("Browser.close", undefined, "s")).rejects.toMatchObject({
      code: "bad_arguments",
    });
    await flush();
    expect(peer.rawBytes).toBe(0);
  });
});

describe("cdp pipe: framing and correlation", () => {
  it("writes NUL-terminated JSON with increasing ids and correlates responses", async () => {
    const { peer, pipe } = setup();
    peer.on("Target.getTargets", () => ({ targetInfos: [] }));
    peer.on("Network.enable", () => ({}));
    const a = await pipe.send("Target.getTargets");
    const b = await pipe.send("Network.enable", undefined, "sess");
    expect(a).toEqual({ targetInfos: [] });
    expect(b).toEqual({});
    expect(peer.received.map((m) => m.id)).toEqual([1, 2]);
    expect(peer.received[1]).toMatchObject({ method: "Network.enable", sessionId: "sess" });
    expect(peer.rawBytes).toBeGreaterThan(0);
  });

  it("reassembles frames split across chunks and splits merged frames", async () => {
    const { peer, pipe, events } = setup();
    const p = pipe.send("Target.getTargets");
    await flush();
    const resp = Buffer.from(`${JSON.stringify({ id: 1, result: { ok: true } })}\0`);
    const ev1 = Buffer.from(
      `${JSON.stringify({ method: "Target.targetCreated", params: { n: 1 } })}\0`,
    );
    const ev2 = Buffer.from(
      `${JSON.stringify({ method: "Target.targetDestroyed", params: { n: 2 } })}\0`,
    );
    const all = Buffer.concat([ev1, resp, ev2]);
    for (let i = 0; i < all.length; i += 7) {
      peer.writeRaw(all.subarray(i, i + 7)); // 7-byte dribble, frames straddle chunks
      await flush();
    }
    await expect(p).resolves.toEqual({ ok: true });
    expect(events.map((e) => e.method)).toEqual(["Target.targetCreated", "Target.targetDestroyed"]);
  });

  it("ignores unknown ids, empty frames, malformed JSON and non-object JSON", async () => {
    const { peer, pipe, events } = setup();
    const p = pipe.send("Target.getTargets");
    await flush();
    peer.writeRaw("\0\0");
    peer.writeRaw("{not json SECRET-CANARY\0");
    peer.writeRaw("[1,2]\0");
    peer.writeRaw('"str"\0');
    peer.writeFrame({ id: 999, result: {} });
    peer.writeFrame({ id: 1, result: { a: 1 } });
    await expect(p).resolves.toEqual({ a: 1 });
    expect(events).toEqual([]);
    expect(pipe.isClosed).toBe(false);
  });

  it("drops events outside the consumed set without dispatch", async () => {
    const { peer, events } = setup();
    peer.event("Network.dataReceived", { requestId: "1" });
    peer.event("Runtime.consoleAPICalled", { args: [] });
    peer.event("Network.responseReceived", { requestId: "1" }, "s");
    await flush();
    expect(events.map((e) => e.method)).toEqual(["Network.responseReceived"]);
    expect(events[0]?.sessionId).toBe("s");
  });

  it("rejects command errors with a code only, no peer message text", async () => {
    const { peer, pipe } = setup();
    const p = pipe.send("Target.getTargets");
    await flush();
    peer.writeFrame({ id: 1, error: { code: -32000, message: "SECRET-CANARY detail" } });
    const err = await p.catch((e: unknown) => e);
    expect(err).toBeInstanceOf(CdpError);
    expect((err as CdpError).code).toBe("command_failed");
    expect((err as CdpError).message).not.toContain("SECRET-CANARY");
  });

  it("bounds in-flight commands", async () => {
    const { pipe } = setup();
    const ps: Promise<unknown>[] = [];
    for (let i = 0; i < MAX_IN_FLIGHT_COMMANDS; i++) ps.push(pipe.send("Target.getTargets"));
    await expect(pipe.send("Target.getTargets")).rejects.toMatchObject({
      code: "too_many_in_flight",
    });
    pipe.close();
    await Promise.allSettled(ps);
  });
});

describe("cdp pipe: timeouts, oversize, EOF", () => {
  it("times out a command without a response", async () => {
    vi.useFakeTimers();
    const { pipe } = setup({ timeout: 500 });
    const p = pipe.send("Target.getTargets");
    const assertion = expect(p).rejects.toMatchObject({ code: "timeout" });
    await vi.advanceTimersByTimeAsync(501);
    await assertion;
  });

  it("closes on an oversize frame and rejects pending, never delivering it", async () => {
    const { peer, pipe, events, closes } = setup({ max: 1024 });
    const p = pipe.send("Target.getTargets");
    const assertion = expect(p).rejects.toMatchObject({ code: "closed" });
    await flush();
    peer.writeRaw(Buffer.alloc(900, 0x61));
    peer.writeRaw(Buffer.alloc(900, 0x61)); // no NUL yet; accumulated > max
    await assertion;
    expect(closes).toEqual(["oversize"]);
    expect(events).toEqual([]);
    await expect(pipe.send("Target.getTargets")).rejects.toMatchObject({ code: "closed" });
  });

  it("closes on a single oversize frame within one chunk", async () => {
    const { peer, closes } = setup({ max: 64 });
    peer.writeRaw(Buffer.concat([Buffer.alloc(200, 0x61), Buffer.from([0])]));
    await flush();
    expect(closes).toEqual(["oversize"]);
  });

  it("rejects pending commands on EOF and reports eof", async () => {
    const { peer, pipe, closes } = setup();
    const p = pipe.send("Target.getTargets");
    const assertion = expect(p).rejects.toMatchObject({ code: "closed" });
    await flush();
    peer.eof();
    await assertion;
    expect(closes).toEqual(["eof"]);
    expect(pipe.isClosed).toBe(true);
  });

  it("rejects pending commands on a stream error", async () => {
    const { peer, pipe, closes } = setup();
    const p = pipe.send("Target.getTargets");
    const assertion = expect(p).rejects.toMatchObject({ code: "closed" });
    await flush();
    peer.output.destroy(new Error("boom SECRET-CANARY"));
    await assertion;
    expect(closes).toEqual(["stream_error"]);
  });
});
