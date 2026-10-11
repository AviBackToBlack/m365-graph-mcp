import { describe, expect, it } from "vitest";
import { CdpPipe } from "./cdp-pipe.js";
import { GraphExplorerCapture, extractBearer, isGraphExplorerUrl } from "./capture.js";
import type { CaptureDiagnostic } from "./capture.js";
import { FakeCdpPeer } from "./fake-cdp.js";
import { MIN_REMAINING_MS } from "./token-gate.js";
import type { TokenGateConfig } from "./token-gate.js";

const TENANT = "11111111-2222-3333-4444-555555555555";
const OTHER_TENANT = "99999999-8888-7777-6666-555555555555";
const CLIENT = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
const NOW = 1_800_000_000_000;
const GE = "https://developer.microsoft.com/en-us/graph/graph-explorer";
const GRAPH = "https://graph.microsoft.com/v1.0/me";
const flush = async (): Promise<void> => {
  for (let i = 0; i < 4; i++) await new Promise((r) => setImmediate(r));
};

function jwt(over: Record<string, unknown> = {}, lifeMs = 3_600_000): string {
  const enc = (o: unknown): string => Buffer.from(JSON.stringify(o)).toString("base64url");
  return `${enc({ alg: "RS256" })}.${enc({
    aud: "https://graph.microsoft.com",
    tid: TENANT,
    scp: "Mail.Read",
    exp: Math.floor((NOW + lifeMs) / 1000),
    ver: "1.0",
    appid: CLIENT,
    ...over,
  })}.c2ln`;
}

interface Env {
  peer: FakeCdpPeer;
  cap: GraphExplorerCapture;
  diags: CaptureDiagnostic[];
  clock: { now: number };
}

async function setup(
  opts: {
    gate?: TokenGateConfig;
    targets?: unknown[];
    attachFails?: boolean;
    attachRetryMs?: number;
  } = {},
): Promise<Env> {
  const peer = new FakeCdpPeer();
  peer.on("Target.setDiscoverTargets", () => ({}));
  peer.on("Target.getTargets", () => ({
    targetInfos: opts.targets ?? [{ targetId: "T1", type: "page", url: GE }],
  }));
  peer.on("Target.attachToTarget", () => (opts.attachFails ? undefined : { sessionId: "S1" }));
  peer.on("Network.enable", () => ({}));
  const cdp = new CdpPipe({ input: peer.input, output: peer.output, commandTimeoutMs: 1000 });
  const diags: CaptureDiagnostic[] = [];
  const clock = { now: NOW };
  const cap = new GraphExplorerCapture({
    cdp,
    gate: opts.gate ?? { tenantId: TENANT },
    now: () => clock.now,
    onDiagnostic: (d) => diags.push(d),
    ...(opts.attachRetryMs !== undefined ? { attachRetryMs: opts.attachRetryMs } : {}),
  });
  await cap.start();
  await flush();
  return { peer, cap, diags, clock };
}

function request(
  e: Env,
  id: string,
  token: string | undefined,
  over: { url?: string; session?: string; headers?: Record<string, string> } = {},
): void {
  const headers = over.headers ?? (token ? { Authorization: `Bearer ${token}` } : {});
  e.peer.event(
    "Network.requestWillBeSent",
    { requestId: id, request: { url: over.url ?? GRAPH, headers } },
    over.session ?? "S1",
  );
}
const respond = (e: Env, id: string, status: number, session = "S1", url = GRAPH): void =>
  e.peer.event("Network.responseReceived", { requestId: id, response: { status, url } }, session);

const navigate = (e: Env, targetId: string, url: string): void =>
  e.peer.event("Target.targetInfoChanged", { targetInfo: { targetId, type: "page", url } });
const extra = (e: Env, id: string, token: string, session = "S1"): void =>
  e.peer.event(
    "Network.requestWillBeSentExtraInfo",
    { requestId: id, headers: { Authorization: `Bearer ${token}`, Cookie: "SECRET-CANARY" } },
    session,
  );
const sent = (e: Env, method: string): { params?: unknown; sessionId?: string }[] =>
  e.peer.received.filter((m) => m.method === method);

describe("graph explorer url matching", () => {
  it("accepts only the exact https origin and exact path", () => {
    for (const u of [
      GE,
      `${GE}/`,
      `${GE}?x=1#y`,
      "https://developer.microsoft.com/graph/graph-explorer/",
      "https://developer.microsoft.com/de-de/graph/graph-explorer/",
      "https://developer.microsoft.com/ja-JP/graph/graph-explorer?x=1",
      "https://developer.microsoft.com/fr-fr/graph/graph-explorer",
    ])
      expect(isGraphExplorerUrl(u), u).toBe(true);
    for (const u of [
      "http://developer.microsoft.com/en-us/graph/graph-explorer",
      "https://developer.microsoft.com.evil.com/en-us/graph/graph-explorer",
      "https://evil.developer.microsoft.com/en-us/graph/graph-explorer",
      "https://developer.microsoft.com:8443/en-us/graph/graph-explorer",
      "https://user@developer.microsoft.com/en-us/graph/graph-explorer",
      "https://developer.microsoft.com/en-us/graph/graph-explorer-evil",
      "https://developer.microsoft.com/en-us/graph/graph-explorer/../x",
      "https://developer.microsoft.com/evil/en-us/graph/graph-explorer",
      "https://developer.microsoft.com/xx/graph/graph-explorer",
      "https://developer.microsoft.com/en-us/foo/graph/graph-explorer",
      "https://developer.microsoft.com/en-us/graph/graph-explorer/extra",
      "https://developer.microsoft.com/en-us/en-us/graph/graph-explorer",
      "https://developer.microsoft.com/eng-us/graph/graph-explorer",
      "https://developer.microsoft.com/e1-us/graph/graph-explorer",
      "https://developer.microsoft.com/en_us/graph/graph-explorer",
      "https://developer.microsoft.com/en-us/graph",
      "https://developer.microsoft.com/graph/graph-explorer//",
      "https://other.microsoft.com/de-de/graph/graph-explorer",
      "http://developer.microsoft.com/de-de/graph/graph-explorer",
      "https://login.microsoftonline.com/en-us/graph/graph-explorer",
      "not a url",
      undefined,
    ])
      expect(isGraphExplorerUrl(u), String(u)).toBe(false);
  });
});

describe("header extraction", () => {
  it("extracts only a single Bearer value, header name case-insensitive", () => {
    expect(extractBearer({ authorization: "Bearer abc.def" })).toBe("abc.def");
    expect(extractBearer({ AUTHORIZATION: "bearer abc" })).toBe("abc");
    expect(extractBearer({ Authorization: "Basic abc" })).toBeUndefined();
    expect(extractBearer({ Authorization: "Bearer a b" })).toBeUndefined();
    expect(extractBearer({ Authorization: "Bearer " })).toBeUndefined();
    expect(extractBearer({ Authorization: "Bearer a", authorization: "Bearer b" })).toBeUndefined();
    expect(extractBearer({ Authorization: ["Bearer a"] })).toBeUndefined();
    expect(extractBearer({ Cookie: "Bearer a" })).toBeUndefined();
    expect(extractBearer(null)).toBeUndefined();
    expect(extractBearer("Bearer a")).toBeUndefined();
  });
});

describe("capture: target selection and session isolation", () => {
  it("attaches (flatten) only to the Graph Explorer page and enables Network on that session", async () => {
    const e = await setup({
      targets: [
        { targetId: "X1", type: "service_worker", url: GE },
        { targetId: "X2", type: "page", url: "https://example.com/" },
        { targetId: "X3", type: "page", url: "https://developer.microsoft.com/evil" },
        { targetId: "T1", type: "page", url: GE },
      ],
    });
    const attach = e.peer.received.filter((m) => m.method === "Target.attachToTarget");
    expect(attach).toHaveLength(1);
    expect(attach[0]?.params).toEqual({ targetId: "T1", flatten: true });
    const enable = e.peer.received.filter((m) => m.method === "Network.enable");
    expect(enable.map((m) => m.sessionId)).toEqual(["S1"]);
    expect(e.cap.current()).toEqual({ state: "reauth_required" });
  });

  it("is unavailable until a page is attached; later page creation attaches", async () => {
    const e = await setup({ targets: [{ targetId: "A", type: "page", url: "about:blank" }] });
    expect(e.cap.current()).toEqual({ state: "unavailable" });
    e.peer.event("Target.targetCreated", {
      targetInfo: { targetId: "T9", type: "page", url: GE },
    });
    await flush();
    expect(e.peer.received.some((m) => m.method === "Target.attachToTarget")).toBe(true);
    expect(e.cap.current().state).toBe("reauth_required");
  });

  it("ignores Network events from other sessions and browser-level events", async () => {
    const e = await setup();
    const t = jwt();
    request(e, "r1", t, { session: "OTHER" });
    respond(e, "r1", 200, "OTHER");
    request(e, "r2", t, { session: "" });
    respond(e, "r2", 200, "");
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
  });

  it("stops treating requests as eligible after the page navigates away", async () => {
    const e = await setup();
    const away = "https://login.microsoftonline.com/x";
    e.peer.on("Target.getTargets", () => ({
      targetInfos: [{ targetId: "T1", type: "page", url: away }],
    }));
    navigate(e, "T1", away);
    await flush();
    request(e, "r1", jwt());
    respond(e, "r1", 200);
    await flush();
    expect(e.diags.filter((d) => d.kind === "token_accepted")).toEqual([]);
    expect(e.cap.current().state).toBe("reauth_required");
  });

  it("forgets the session when the target is destroyed", async () => {
    const e = await setup();
    e.peer.on("Target.getTargets", () => ({ targetInfos: [] }));
    e.peer.event("Target.targetDestroyed", { targetId: "T1" });
    await flush();
    request(e, "r1", jwt());
    respond(e, "r1", 200);
    await flush();
    expect(e.cap.current().state).not.toBe("ok");
  });
});

describe("capture: request/response correlation and the gate", () => {
  it("accepts a good Graph token after a 200 response", async () => {
    const e = await setup();
    const t = jwt();
    request(e, "r1", t);
    await flush();
    expect(e.cap.current().state).toBe("reauth_required"); // not until the response is seen
    respond(e, "r1", 200);
    await flush();
    expect(e.cap.current()).toEqual({
      state: "ok",
      token: t,
      expMs: Math.floor((NOW + 3_600_000) / 1000) * 1000,
    });
  });

  it("accepts on 403 but rejects on 401", async () => {
    const e = await setup();
    request(e, "a", jwt({ scp: "Mail.Read" }));
    respond(e, "a", 401);
    await flush();
    expect(e.cap.current()).toEqual({ state: "reauth_required", reason: "http_401" });
    request(e, "b", jwt());
    respond(e, "b", 403);
    await flush();
    expect(e.cap.current().state).toBe("ok");
  });

  it("supports response arriving before ExtraInfo carries the header", async () => {
    const e = await setup();
    const t = jwt();
    request(e, "r1", undefined); // no Authorization in the main event
    respond(e, "r1", 200);
    e.peer.event(
      "Network.requestWillBeSentExtraInfo",
      { requestId: "r1", headers: { Authorization: `Bearer ${t}`, Cookie: "SECRET-CANARY" } },
      "S1",
    );
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t });
  });

  it("ignores ExtraInfo for requests not known to target Graph", async () => {
    const e = await setup();
    request(e, "r1", undefined, { url: "https://evil.example/x" });
    respond(e, "r1", 200, "S1", "https://evil.example/x");
    e.peer.event(
      "Network.requestWillBeSentExtraInfo",
      { requestId: "r1", headers: { Authorization: `Bearer ${jwt()}` } },
      "S1",
    );
    e.peer.event(
      "Network.requestWillBeSentExtraInfo",
      { requestId: "never-seen", headers: { Authorization: `Bearer ${jwt()}` } },
      "S1",
    );
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
  });

  it("never captures bearers sent to non-Graph or look-alike origins", async () => {
    const e = await setup();
    const urls = [
      "https://graph.microsoft.com.evil.com/v1.0/me",
      "https://evil.graph.microsoft.com/v1.0/me",
      "http://graph.microsoft.com/v1.0/me",
      "https://graph.microsoft.com:8443/v1.0/me",
      "https://login.microsoftonline.com/token",
      GE,
    ];
    urls.forEach((u, i) => {
      request(e, `r${i}`, jwt(), { url: u });
      respond(e, `r${i}`, 200, "S1", u);
    });
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
    expect(e.diags.filter((d) => d.kind === "token_accepted")).toEqual([]);
  });

  it("ignores multiple, non-Bearer or missing Authorization headers", async () => {
    const e = await setup();
    request(e, "a", undefined, {
      headers: { Authorization: `Bearer ${jwt()}`, authorization: `Bearer ${jwt()}` },
    });
    request(e, "b", undefined, { headers: { Authorization: `Basic ${jwt()}` } });
    request(e, "c", undefined, { headers: {} });
    for (const id of ["a", "b", "c"]) respond(e, id, 200);
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
  });

  it("discards entries on loadingFailed and when the response never arrives", async () => {
    const e = await setup();
    request(e, "a", jwt());
    e.peer.event("Network.loadingFailed", { requestId: "a" }, "S1");
    respond(e, "a", 200);
    request(e, "b", jwt());
    await flush();
    e.clock.now += 16_000; // beyond the pending TTL
    respond(e, "b", 200);
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
  });

  it("bounds the pending map", async () => {
    const e = await setup();
    for (let i = 0; i < 200; i++) request(e, `r${i}`, jwt());
    await flush();
    respond(e, "r0", 200); // evicted long ago
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
    respond(e, "r199", 200);
    await flush();
    expect(e.cap.current().state).toBe("ok");
  });

  it("applies tenant pin, unsupported format and client pin through the gate", async () => {
    let e = await setup();
    request(e, "a", jwt({ tid: OTHER_TENANT }));
    respond(e, "a", 200);
    await flush();
    expect(e.cap.current()).toEqual({ state: "reauth_required", reason: "tenant" });
    request(e, "b", "opaque-not-a-jwt");
    respond(e, "b", 200);
    await flush();
    expect(e.cap.current()).toEqual({
      state: "unsupported_token_format",
      reason: "unsupported_token_format",
    });
    request(e, "c", jwt());
    respond(e, "c", 200);
    await flush();
    expect(e.cap.current().state).toBe("ok");

    e = await setup({ gate: { tenantId: TENANT, clientPin: { guid: CLIENT, claim: "auto" } } });
    expect(e.cap.clientPinEnabled).toBe(true);
    request(e, "d", jwt({ appid: OTHER_TENANT }));
    respond(e, "d", 200);
    await flush();
    expect(e.cap.current()).toEqual({ state: "reauth_required", reason: "client_pin" });
  });

  it("keeps a still-valid token on later rejections or earlier-expiring tokens; replaces with later", async () => {
    const e = await setup();
    const t1 = jwt({}, 3_600_000);
    request(e, "a", t1);
    respond(e, "a", 200);
    await flush();
    request(e, "b", "opaque");
    respond(e, "b", 200);
    request(e, "c", jwt({ tid: OTHER_TENANT }));
    respond(e, "c", 200);
    request(e, "d", jwt({ scp: "x" }, 1_800_000)); // earlier expiry
    respond(e, "d", 200);
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t1 });
    const t2 = jwt({ scp: "y" }, 7_200_000);
    request(e, "e", t2);
    respond(e, "e", 200);
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t2 });
  });

  it("drops a held token when Graph answers it with 401", async () => {
    const e = await setup();
    const t = jwt();
    request(e, "a", t);
    respond(e, "a", 200);
    await flush();
    request(e, "b", t);
    respond(e, "b", 401);
    await flush();
    expect(e.cap.current()).toEqual({ state: "reauth_required", reason: "http_401" });
  });

  it("reports reauth_required once the held token falls below the serve floor", async () => {
    const e = await setup();
    request(e, "a", jwt());
    respond(e, "a", 200);
    await flush();
    e.clock.now += 3_600_000 - MIN_REMAINING_MS + 1000;
    expect(e.cap.current()).toEqual({ state: "reauth_required", reason: "lifetime" });
  });

  it("markUnavailable (browser exit) drops the token", async () => {
    const e = await setup();
    request(e, "a", jwt());
    respond(e, "a", 200);
    await flush();
    e.cap.markUnavailable();
    expect(e.cap.current()).toEqual({ state: "unavailable" });
  });

  it("does not leak token, headers or claims through diagnostics (canary)", async () => {
    const e = await setup();
    const canary = jwt({ tid: OTHER_TENANT, scp: "CANARY-SCOPE", sub: "CANARY-SUB" });
    request(e, "a", canary, { headers: { Authorization: `Bearer ${canary}`, Cookie: "CANARY-C" } });
    respond(e, "a", 200);
    request(e, "b", "CANARY-OPAQUE");
    respond(e, "b", 200);
    await flush();
    const dump = JSON.stringify(e.diags);
    for (const s of [canary, "CANARY", OTHER_TENANT, "Bearer"]) expect(dump).not.toContain(s);
    // Failure snapshots carry only fixed reason categories.
    expect(JSON.stringify(e.cap.current())).not.toContain("CANARY");
  });

  it("reports attach failure as a diagnostic without throwing or leaking", async () => {
    const e = await setup({ attachFails: true });
    expect(e.diags.map((d) => d.kind)).toContain("attach_failed");
    expect(e.cap.current().state).toBe("unavailable");
  });
});

describe("capture: claim release, re-attachment and attach race", () => {
  const twoPages = (e: Env, targets: () => unknown[]): void => {
    e.peer.on("Target.getTargets", () => ({ targetInfos: targets() }));
    e.peer.on("Target.attachToTarget", (m) => ({
      sessionId: (m.params as { targetId: string }).targetId === "T1" ? "S1" : "S2",
    }));
  };
  const attachedTargets = (e: Env): string[] =>
    sent(e, "Target.attachToTarget").map((m) => (m.params as { targetId: string }).targetId);

  it("retries a failed attach after the delay, so an idle signed-in page is still claimed", async () => {
    const e = await setup({ attachFails: true, attachRetryMs: 5 });
    expect(e.diags.some((d) => d.kind === "attach_failed")).toBe(true);
    expect(attachedTargets(e)).toEqual(["T1"]);
    // The page stays idle (no further target events); the retry rescans and attaches.
    e.peer.on("Target.attachToTarget", () => ({ sessionId: "S1" }));
    await new Promise((r) => setTimeout(r, 40));
    await flush();
    expect(attachedTargets(e)).toEqual(["T1", "T1"]);
    const t = jwt();
    request(e, "a", t);
    respond(e, "a", 200);
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t });
  });

  it("retries a rescan whose Target.getTargets failed", async () => {
    const e = await setup({ attachRetryMs: 5 });
    let failOnce = true;
    e.peer.on("Target.getTargets", (m) => {
      if (failOnce) {
        failOnce = false;
        e.peer.writeFrame({ id: m.id, error: { code: -32000, message: "busy" } });
        return {};
      }
      return { targetInfos: [{ targetId: "T2", type: "page", url: GE }] };
    });
    e.peer.on("Target.attachToTarget", () => ({ sessionId: "S2" }));
    navigate(e, "T1", "https://login.microsoftonline.com/x"); // release -> rescan fails
    await flush();
    expect(attachedTargets(e)).toEqual(["T1"]);
    await new Promise((r) => setTimeout(r, 40));
    await flush();
    expect(attachedTargets(e)).toEqual(["T1", "T2"]);
  });

  it("does not retry after stop()", async () => {
    const e = await setup({ attachFails: true, attachRetryMs: 5 });
    e.cap.stop();
    e.peer.on("Target.attachToTarget", () => ({ sessionId: "S1" }));
    await new Promise((r) => setTimeout(r, 40));
    await flush();
    expect(attachedTargets(e)).toEqual(["T1"]);
  });

  it("releases the claim when the page navigates away and attaches another Graph Explorer page", async () => {
    const e = await setup();
    const away = "https://login.microsoftonline.com/x";
    twoPages(e, () => [
      { targetId: "T1", type: "page", url: away },
      { targetId: "T2", type: "page", url: GE },
    ]);
    navigate(e, "T1", away);
    await flush();
    expect(sent(e, "Target.detachFromTarget").map((m) => m.params)).toEqual([{ sessionId: "S1" }]);
    expect(attachedTargets(e)).toEqual(["T1", "T2"]);
    // The old session is dead; the new one works.
    const t = jwt();
    request(e, "old", t, { session: "S1" });
    respond(e, "old", 200, "S1");
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
    request(e, "new", t, { session: "S2" });
    respond(e, "new", 200, "S2");
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t });
  });

  it("drops pending entries of the released session, keeps a valid held token, re-attaches on return", async () => {
    const e = await setup();
    const t = jwt();
    request(e, "a", t);
    respond(e, "a", 200);
    await flush();
    request(e, "pending", jwt({ scp: "y" }, 7_200_000));
    navigate(e, "T1", "https://login.microsoftonline.com/x");
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t });
    e.peer.on("Target.attachToTarget", () => ({ sessionId: "S3" }));
    navigate(e, "T1", GE);
    await flush();
    expect(attachedTargets(e)).toEqual(["T1", "T1"]);
    respond(e, "pending", 200, "S3"); // the old pending request did not survive
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t });
  });

  it("rescans after the claimed target is destroyed or lost", async () => {
    let e = await setup();
    twoPages(e, () => [{ targetId: "T2", type: "page", url: GE }]);
    e.peer.event("Target.targetDestroyed", { targetId: "T1" });
    await flush();
    expect(attachedTargets(e)).toEqual(["T1", "T2"]);
    request(e, "n", jwt(), { session: "S2" });
    respond(e, "n", 200, "S2");
    await flush();
    expect(e.cap.current().state).toBe("ok");

    e = await setup();
    twoPages(e, () => [{ targetId: "T2", type: "page", url: GE }]);
    e.peer.event("Target.detachedFromTarget", { sessionId: "S1" });
    await flush();
    expect(attachedTargets(e)).toEqual(["T1", "T2"]);
  });

  it("does not claim a page that navigated away while the attach was in flight", async () => {
    const e = await setup({ targets: [{ targetId: "A", type: "page", url: "about:blank" }] });
    e.peer.on("Network.enable", () => {
      navigate(e, "T9", "https://login.microsoftonline.com/x");
      return {};
    });
    e.peer.event("Target.targetCreated", { targetInfo: { targetId: "T9", type: "page", url: GE } });
    await flush();
    expect(sent(e, "Target.detachFromTarget").map((m) => m.params)).toEqual([{ sessionId: "S1" }]);
    request(e, "r", jwt());
    respond(e, "r", 200);
    await flush();
    expect(e.cap.current().state).toBe("unavailable");
    expect(e.diags.map((d) => d.kind)).not.toContain("target_attached");
  });

  it("claims the page if it returned to Graph Explorer before the attach committed", async () => {
    const e = await setup({ targets: [{ targetId: "A", type: "page", url: "about:blank" }] });
    e.peer.on("Network.enable", () => {
      navigate(e, "T9", "https://login.microsoftonline.com/x");
      navigate(e, "T9", GE);
      return {};
    });
    e.peer.event("Target.targetCreated", { targetInfo: { targetId: "T9", type: "page", url: GE } });
    await flush();
    expect(e.diags.map((d) => d.kind)).toContain("target_attached");
  });
});

describe("capture: ExtraInfo ordering, redirects and renewal", () => {
  const redirect = (e: Env, id: string, headers: Record<string, string>): void =>
    e.peer.event(
      "Network.requestWillBeSent",
      { requestId: id, redirectResponse: { status: 302 }, request: { url: GRAPH, headers } },
      "S1",
    );

  it("merges an ExtraInfo bearer that arrives before requestWillBeSent", async () => {
    const e = await setup();
    const t = jwt();
    extra(e, "r1", t);
    request(e, "r1", undefined);
    respond(e, "r1", 200);
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t });
  });

  it("never lets a buffered ExtraInfo bearer reach a non-Graph request or a later id reuse", async () => {
    const e = await setup();
    extra(e, "r1", jwt());
    request(e, "r1", undefined, { url: "https://evil.example/x" });
    respond(e, "r1", 200, "S1", "https://evil.example/x");
    request(e, "r1", undefined); // id reuse: the buffer must be gone
    respond(e, "r1", 200);
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
  });

  it("drops buffered ExtraInfo on loadingFailed, TTL and bounds", async () => {
    const e = await setup();
    extra(e, "a", jwt());
    e.peer.event("Network.loadingFailed", { requestId: "a" }, "S1");
    request(e, "a", undefined);
    respond(e, "a", 200);
    extra(e, "b", jwt());
    await flush();
    e.clock.now += 16_000;
    request(e, "b", undefined);
    respond(e, "b", 200);
    for (let i = 0; i < 200; i++) extra(e, `x${i}`, jwt());
    request(e, "x0", undefined); // evicted by the bound
    respond(e, "x0", 200);
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
    request(e, "x199", undefined); // newest survives
    respond(e, "x199", 200);
    await flush();
    expect(e.cap.current().state).toBe("ok");
  });

  it("does not attach an earlier hop's bearer to a redirected request", async () => {
    const e = await setup();
    extra(e, "r1", jwt()); // hop 1 (non-Graph) ExtraInfo, buffered
    request(e, "r1", undefined, { url: "https://evil.example/x" });
    redirect(e, "r1", {});
    extra(e, "r1", jwt({ scp: "y" })); // late ExtraInfo cannot be tied to the final hop
    respond(e, "r1", 200);
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");

    extra(e, "r2", jwt()); // buffered before a redirect hop of the same id
    redirect(e, "r2", {});
    respond(e, "r2", 200);
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");
  });

  it("resets a Graph bearer from a prior hop and accepts only the final hop's own header", async () => {
    const e = await setup();
    request(e, "r1", jwt());
    redirect(e, "r1", {});
    respond(e, "r1", 200);
    await flush();
    expect(e.cap.current().state).toBe("reauth_required");

    const b = jwt({ scp: "y" });
    redirect(e, "r2", { Authorization: `Bearer ${b}` });
    respond(e, "r2", 200);
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: b });
  });

  it("replaces the held token with a different bearer of equal expiry, never churns on the same one", async () => {
    const e = await setup();
    const t1 = jwt({ scp: "a" });
    const t2 = jwt({ scp: "b" }); // same exp
    request(e, "a", t1);
    respond(e, "a", 200);
    await flush();
    request(e, "b", t2);
    respond(e, "b", 200);
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t2 });
    request(e, "c", t1);
    respond(e, "c", 200);
    await flush();
    expect(e.cap.current()).toMatchObject({ state: "ok", token: t1 }); // different again: replaced
  });
});
