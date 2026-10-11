import fs from "node:fs";
import { describe, expect, it, vi } from "vitest";
import {
  ACCEPTANCE_PREFIX,
  DISABLED_ACCEPTANCE_DIAGNOSTICS,
  MAX_REMAINING_MINUTES,
  createAcceptanceDiagnostics,
} from "./acceptance-diagnostics.js";
import type {
  AcceptanceDiagnostics,
  AudienceCategory,
  AuthorizationObservation,
  ClientClaimKind,
  GraphExplorerPageShape,
  HostClass,
  PortClass,
  Presence,
  SchemeClass,
  TenantCheck,
  TokenVersionCategory,
} from "./acceptance-diagnostics.js";
import { CdpPipe } from "./cdp-pipe.js";
import { GraphExplorerCapture, classifyGraphExplorerPage } from "./capture.js";
import type { CaptureDiagnostic } from "./capture.js";
import { FakeCdpPeer } from "./fake-cdp.js";
import { classifyGraphOrigin, classifyTokenShape, evaluateCandidate } from "./token-gate.js";
import type { TokenCandidate, TokenGateConfig } from "./token-gate.js";

const TENANT = "11111111-2222-3333-4444-555555555555";
const OTHER_TENANT = "99999999-8888-7777-6666-555555555555";
const CLIENT = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
const OTHER_CLIENT = "cccccccc-dddd-eeee-ffff-000000000000";
const NOW = 1_800_000_000_000;
const SCOPE = "Canary.Scope.Read";
const SESSION = "SESSCANARY7Q";
const TARGET = "TGTCANARY9Z";
const REQ = "REQCANARY3X";
const PAGE_CANONICAL = "https://developer.microsoft.com/graph/graph-explorer/";
const PAGE_LOCALIZED = "https://developer.microsoft.com/de-de/graph/graph-explorer?canary=PAGEQ7";
const GRAPH = "https://graph.microsoft.com/v1.0/me/messages?$search=GRAPHQ7";

/** The closed allowlist of every line the acceptance layer can ever emit. */
const ALLOWED: RegExp[] = [
  /^acceptance: mode=enabled$/,
  /^acceptance: token_shape audience=(graph_url|graph_url_trailing_slash|graph_application_id) client_claim=(appid|azp|both|none) token_version=(v1|v2|other_or_absent) scp=(present|absent) roles=(present|absent) tenant=(matched|mismatch) remaining_minutes=\d{1,5}$/,
  /^acceptance: graph_request_exact_origin_match=(yes|no) graph_request_scheme=(https|other) graph_request_host=(expected_graph_host|other) graph_request_port=(default_https|other)$/,
  /^acceptance: graph_explorer_page_shape=canonical$/,
  /^acceptance: graph_explorer_page_shape=localized locale_segment_shape=xx-xx$/,
  /^acceptance: authorization_observation=(main_event|extra_info_before_main|extra_info_after_main)$/,
  /^acceptance: extra_info_authorization=(present|not_observed|present_after_response)$/,
  /^acceptance: redirect_seen=yes$/,
  /^acceptance: redirect_final_hop_authorization=main_event$/,
  /^acceptance: redirect_extra_info_discarded=yes$/,
];
const MAX_LINE_LENGTH = 200;

function expectAllowed(lines: string[]): void {
  for (const line of lines) {
    expect(line.length, line).toBeLessThanOrEqual(MAX_LINE_LENGTH);
    expect(
      ALLOWED.some((re) => re.test(line)),
      line,
    ).toBe(true);
  }
}

function collector(): { lines: string[]; acc: AcceptanceDiagnostics } {
  const lines: string[] = [];
  return { lines, acc: createAcceptanceDiagnostics((l) => lines.push(l)) };
}

describe("acceptance diagnostics: typed events and closed output", () => {
  it("emits only lines from the closed allowlist across every enum combination", () => {
    const { lines, acc } = collector();
    const audiences: AudienceCategory[] = [
      "graph_url",
      "graph_url_trailing_slash",
      "graph_application_id",
    ];
    const claims: ClientClaimKind[] = ["appid", "azp", "both", "none"];
    const versions: TokenVersionCategory[] = ["v1", "v2", "other_or_absent"];
    const presence: Presence[] = ["present", "absent"];
    const tenants: TenantCheck[] = ["matched", "mismatch"];
    const schemes: SchemeClass[] = ["https", "other"];
    const hosts: HostClass[] = ["expected_graph_host", "other"];
    const ports: PortClass[] = ["default_https", "other"];
    const pages: GraphExplorerPageShape[] = ["canonical", "localized"];
    const auths: AuthorizationObservation[] = [
      "main_event",
      "extra_info_before_main",
      "extra_info_after_main",
    ];
    acc.modeEnabled();
    for (const audience of audiences)
      for (const clientClaim of claims)
        for (const version of versions)
          for (const scp of presence)
            for (const roles of presence)
              for (const tenant of tenants)
                for (const remainingMinutes of [0, 1, 1439, MAX_REMAINING_MINUTES, 10 ** 9, -5])
                  acc.tokenShape({
                    audience,
                    clientClaim,
                    version,
                    scp,
                    roles,
                    tenant,
                    remainingMinutes,
                  });
    for (const exactOriginMatch of [true, false])
      for (const scheme of schemes)
        for (const host of hosts)
          for (const port of ports)
            acc.graphRequestOrigin({ exactOriginMatch, scheme, host, port });
    for (const p of pages) acc.graphExplorerPage(p);
    for (const a of auths) acc.authorizationObservation(a);
    acc.extraInfoAuthorization("present");
    acc.extraInfoAuthorization("not_observed");
    acc.extraInfoAuthorization("present_after_response");
    acc.redirectSeen();
    acc.redirectFinalHopAuthorizationFromMainEvent();
    acc.redirectExtraInfoDiscarded();

    expect(lines.length).toBeGreaterThan(100);
    expectAllowed(lines);
    expect(lines.every((l) => l.startsWith(ACCEPTANCE_PREFIX))).toBe(true);
    // Lifetime is clamped, never raw.
    expect(lines.some((l) => l.endsWith(`remaining_minutes=${MAX_REMAINING_MINUTES}`))).toBe(true);
    expect(lines.some((l) => l.endsWith("remaining_minutes=0"))).toBe(true);
    expect(lines.join("\n")).not.toContain("1000000000");
    expect(lines.join("\n")).not.toContain("remaining_minutes=-");
  });

  it("drops anything outside the closed sets, even when typing is bypassed", () => {
    const { lines, acc } = collector();
    const evil = "CANARY-https://x.example/?q=1";
    acc.tokenShape({
      audience: evil,
      clientClaim: "appid",
      version: "v1",
      scp: "present",
      roles: "absent",
      tenant: "matched",
      remainingMinutes: 5,
    } as never);
    acc.tokenShape({
      audience: "graph_url",
      clientClaim: "appid",
      version: "v1",
      scp: "present",
      roles: "absent",
      tenant: "matched",
      remainingMinutes: Number.NaN,
    });
    acc.tokenShape({ remainingMinutes: 5 } as never);
    acc.tokenShape(undefined as never);
    acc.graphRequestOrigin({
      exactOriginMatch: "yes",
      scheme: "https",
      host: evil,
      port: "default_https",
    } as never);
    acc.graphRequestOrigin(undefined as never);
    acc.graphExplorerPage(evil as never);
    acc.graphExplorerPage("xx-xx" as never);
    acc.authorizationObservation(evil as never);
    acc.authorizationObservation(undefined as never);
    acc.extraInfoAuthorization(evil as never);
    acc.extraInfoAuthorization("absent" as never);
    expect(lines).toEqual([]);
  });

  it("a disabled instance is inert and its methods do nothing", () => {
    const d = DISABLED_ACCEPTANCE_DIAGNOSTICS;
    expect(d.enabled).toBe(false);
    expect(Object.isFrozen(d)).toBe(true);
    expect(() => {
      d.modeEnabled();
      d.redirectSeen();
      d.redirectExtraInfoDiscarded();
      d.redirectFinalHopAuthorizationFromMainEvent();
      d.authorizationObservation("main_event");
      d.graphExplorerPage("canonical");
    }).not.toThrow();
  });

  it("imports nothing: no file system, network or process access in the module", () => {
    const source = fs.readFileSync(new URL("./acceptance-diagnostics.ts", import.meta.url), "utf8");
    expect(source).not.toMatch(/^\s*import\s/m);
    expect(source).not.toMatch(/\bfrom\s+["']/);
    expect(source).not.toMatch(/require\(|import\(|node:|\bprocess\b|\bconsole\b|\bfs\b/);
  });
});

describe("acceptance diagnostics: classifiers", () => {
  const claims = (over: Record<string, unknown> = {}) => ({
    aud: "https://graph.microsoft.com",
    tid: TENANT,
    scp: SCOPE,
    ver: "1.0",
    appid: CLIENT,
    ...over,
  });

  it("classifies audience, client claim, version, scp, roles, tenant and lifetime", () => {
    const shape = (over: Record<string, unknown>, ms = 90 * 60_000 + 59_999) => {
      const s = classifyTokenShape(claims(over), TENANT, ms);
      if (s === undefined) throw new Error("expected a shape");
      return s;
    };
    expect(shape({})).toEqual({
      audience: "graph_url",
      clientClaim: "appid",
      version: "v1",
      scp: "present",
      roles: "absent",
      tenant: "matched",
      remainingMinutes: 90,
    });
    expect(shape({ aud: "https://graph.microsoft.com/" }).audience).toBe(
      "graph_url_trailing_slash",
    );
    expect(shape({ aud: "00000003-0000-0000-c000-000000000000" }).audience).toBe(
      "graph_application_id",
    );
    expect(shape({ appid: undefined, azp: OTHER_CLIENT, ver: "2.0" })).toMatchObject({
      clientClaim: "azp",
      version: "v2",
    });
    expect(shape({ azp: OTHER_CLIENT }).clientClaim).toBe("both");
    expect(shape({ appid: undefined }).clientClaim).toBe("none");
    expect(shape({ ver: undefined }).version).toBe("other_or_absent");
    expect(shape({ ver: "3.0" }).version).toBe("other_or_absent");
    expect(shape({ ver: 1 }).version).toBe("other_or_absent");
    expect(shape({ roles: [] }).roles).toBe("present");
    expect(shape({ scp: "  " }).scp).toBe("absent");
    expect(classifyTokenShape(claims(), OTHER_TENANT, 0)?.tenant).toBe("mismatch");
    expect(classifyTokenShape(claims({ tid: TENANT.toUpperCase() }), TENANT, 0)?.tenant).toBe(
      "matched",
    );
  });

  it("returns no shape (rather than mislabeling) for an audience outside the mapped set", () => {
    for (const aud of ["https://graph.microsoft.us", "api://other", undefined, 7]) {
      expect(classifyTokenShape(claims({ aud }), TENANT, 0)).toBeUndefined();
    }
  });

  it("classifies the Graph origin with the same exact rule as the gate", () => {
    const yes = { exactOriginMatch: true };
    expect(classifyGraphOrigin(GRAPH)).toEqual({
      ...yes,
      scheme: "https",
      host: "expected_graph_host",
      port: "default_https",
    });
    expect(classifyGraphOrigin("https://GRAPH.microsoft.com:443/v1.0/me")).toMatchObject(yes);
    expect(classifyGraphOrigin("http://graph.microsoft.com/v1.0/me")).toEqual({
      exactOriginMatch: false,
      scheme: "other",
      host: "expected_graph_host",
      port: "other",
    });
    expect(classifyGraphOrigin("https://graph.microsoft.com:8443/x")).toEqual({
      exactOriginMatch: false,
      scheme: "https",
      host: "expected_graph_host",
      port: "other",
    });
    expect(classifyGraphOrigin("https://graph.microsoft.com.evil.example/x")).toMatchObject({
      exactOriginMatch: false,
      host: "other",
    });
    expect(classifyGraphOrigin("https://user@graph.microsoft.com/x").exactOriginMatch).toBe(false);
    expect(classifyGraphOrigin("not a url")).toEqual({
      exactOriginMatch: false,
      scheme: "other",
      host: "other",
      port: "other",
    });
  });

  it("classifies the Graph Explorer page shape with the single URL matcher", () => {
    expect(classifyGraphExplorerPage(PAGE_CANONICAL)).toBe("canonical");
    expect(classifyGraphExplorerPage("https://developer.microsoft.com/graph/graph-explorer")).toBe(
      "canonical",
    );
    expect(classifyGraphExplorerPage(PAGE_LOCALIZED)).toBe("localized");
    expect(
      classifyGraphExplorerPage("https://developer.microsoft.com/EN-us/graph/graph-explorer/"),
    ).toBe("localized");
    for (const bad of [
      "https://developer.microsoft.com/xx/graph/graph-explorer",
      "https://developer.microsoft.com/en-us/en-us/graph/graph-explorer",
      "https://developer.microsoft.com.evil.example/graph/graph-explorer",
      "http://developer.microsoft.com/graph/graph-explorer",
      undefined,
      42,
    ]) {
      expect(classifyGraphExplorerPage(bad)).toBeUndefined();
    }
  });
});

function b64(o: unknown): string {
  return Buffer.from(JSON.stringify(o)).toString("base64url");
}
function jwt(over: Record<string, unknown> = {}, lifeMs = 3_600_000): string {
  return `${b64({ alg: "RS256", kid: "KIDCANARY2P" })}.${b64({
    aud: "https://graph.microsoft.com",
    tid: TENANT,
    scp: SCOPE,
    sub: "SUBCANARY5K",
    exp: Math.floor((NOW + lifeMs) / 1000),
    ver: "1.0",
    appid: CLIENT,
    ...over,
  })}.SIGCANARY8M`;
}

describe("acceptance diagnostics: gate integration", () => {
  const gate: TokenGateConfig = { tenantId: TENANT };
  const cand = (bearer: string, over: Partial<TokenCandidate> = {}): TokenCandidate => ({
    bearer,
    requestUrl: GRAPH,
    responseStatus: 200,
    sourceIsBridgeTarget: true,
    ...over,
  });

  it("emits one token_shape line only for an accepted token, and the result is unchanged", () => {
    const { lines, acc } = collector();
    const token = jwt();
    const on = evaluateCandidate(cand(token), gate, NOW, acc);
    const off = evaluateCandidate(cand(token), gate, NOW);
    expect(on).toEqual(off);
    expect(lines).toEqual([
      "acceptance: token_shape audience=graph_url client_claim=appid token_version=v1 scp=present roles=absent tenant=matched remaining_minutes=60",
    ]);
  });

  it("emits no shape for any rejection, including tenant and client-pin failures", () => {
    const { lines, acc } = collector();
    const pinned: TokenGateConfig = {
      tenantId: TENANT,
      clientPin: { guid: CLIENT, claim: "auto" },
    };
    const cases: [TokenCandidate, TokenGateConfig, string][] = [
      [cand(jwt({ tid: OTHER_TENANT })), gate, "tenant"],
      [cand(jwt({ appid: OTHER_CLIENT })), pinned, "client_pin"],
      [cand(jwt({ ver: "3.0" })), pinned, "client_pin"],
      [cand(jwt({ aud: "https://other.example" })), gate, "audience"],
      [cand(jwt({ scp: undefined })), gate, "not_delegated"],
      [cand(jwt({ roles: ["Mail.Read"] })), gate, "not_delegated"],
      [cand(jwt({}, 60_000)), gate, "lifetime"],
      [cand(jwt({ nbf: Math.floor(NOW / 1000) + 3600 })), gate, "not_yet_valid"],
      [cand(jwt(), { responseStatus: 401 }), gate, "http_401"],
      [cand(jwt(), { responseStatus: undefined }), gate, "no_response"],
      [cand(jwt(), { requestUrl: "https://evil.example/x" }), gate, "wrong_origin"],
      [cand(jwt(), { sourceIsBridgeTarget: false }), gate, "wrong_source"],
      [cand("opaque-token"), gate, "unsupported_token_format"],
    ];
    for (const [c, g, reason] of cases) {
      expect(evaluateCandidate(c, g, NOW, acc)).toEqual({ ok: false, reason });
    }
    expect(lines).toEqual([]);
  });

  it("does no classification work when disabled", () => {
    const boom = (): never => {
      throw new Error("diagnostics invoked while disabled");
    };
    const trap: AcceptanceDiagnostics = {
      enabled: false,
      modeEnabled: boom,
      tokenShape: boom,
      graphRequestOrigin: boom,
      graphExplorerPage: boom,
      authorizationObservation: boom,
      extraInfoAuthorization: boom,
      redirectSeen: boom,
      redirectFinalHopAuthorizationFromMainEvent: boom,
      redirectExtraInfoDiscarded: boom,
    };
    expect(evaluateCandidate(cand(jwt()), gate, NOW, trap).ok).toBe(true);
  });
});

// ---- full fake-CDP flows ----

const flush = async (): Promise<void> => {
  for (let i = 0; i < 4; i++) await new Promise((r) => setImmediate(r));
};

interface Flow {
  /** The URL the fake browser currently reports for the page target. */
  page: { url: string };
  peer: FakeCdpPeer;
  cap: GraphExplorerCapture;
  diags: CaptureDiagnostic[];
}

async function startFlow(acc?: AcceptanceDiagnostics, pageUrl = PAGE_LOCALIZED): Promise<Flow> {
  const peer = new FakeCdpPeer();
  const page = { url: pageUrl };
  peer.on("Target.setDiscoverTargets", () => ({}));
  peer.on("Target.getTargets", () => ({
    targetInfos: [{ targetId: TARGET, type: "page", url: page.url }],
  }));
  peer.on("Target.attachToTarget", () => ({ sessionId: SESSION }));
  peer.on("Network.enable", () => ({}));
  peer.on("Target.detachFromTarget", () => ({}));
  const cdp = new CdpPipe({ input: peer.input, output: peer.output, commandTimeoutMs: 1000 });
  const diags: CaptureDiagnostic[] = [];
  const cap = new GraphExplorerCapture({
    cdp,
    gate: { tenantId: TENANT },
    now: () => NOW,
    onDiagnostic: (d) => diags.push(d),
    ...(acc ? { acceptanceDiagnostics: acc } : {}),
  });
  await cap.start();
  await flush();
  return { page, peer, cap, diags };
}

const send = (f: Flow, id: string, token: string | undefined, over: Record<string, unknown> = {}) =>
  f.peer.event(
    "Network.requestWillBeSent",
    {
      requestId: id,
      request: {
        url: GRAPH,
        headers: token ? { Authorization: `Bearer ${token}`, Cookie: "COOKIECANARY" } : {},
      },
      ...over,
    },
    SESSION,
  );
const extraInfo = (f: Flow, id: string, token: string) =>
  f.peer.event(
    "Network.requestWillBeSentExtraInfo",
    { requestId: id, headers: { Authorization: `Bearer ${token}`, Cookie: "COOKIECANARY" } },
    SESSION,
  );
const answer = (f: Flow, id: string, status = 200) =>
  f.peer.event(
    "Network.responseReceived",
    { requestId: id, response: { status, url: GRAPH } },
    SESSION,
  );
const navigate = (f: Flow, url: string): void => {
  f.page.url = url;
  f.peer.event("Target.targetInfoChanged", { targetInfo: { targetId: TARGET, type: "page", url } });
};

const T1 = jwt({}); // v1, appid, aud url
const T2 = jwt(
  { aud: "https://graph.microsoft.com/", appid: undefined, azp: CLIENT, ver: "2.0" },
  7_200_000,
);
const T3 = jwt(
  { aud: "00000003-0000-0000-c000-000000000000", azp: CLIENT, ver: undefined },
  3_700_000,
);
const T4 = jwt({ scp: "Canary.Scope.Two" }, 3_650_000);
const T_BAD = jwt({ tid: OTHER_TENANT });

/** One scripted session covering every observation kind. */
async function runScript(f: Flow): Promise<void> {
  send(f, `${REQ}-main`, T1); // bearer in the main event
  extraInfo(f, `${REQ}-main`, T1); // ...and also in a later ExtraInfo (checklist item 4)
  answer(f, `${REQ}-main`);
  extraInfo(f, `${REQ}-before`, T2); // ExtraInfo before the main event
  send(f, `${REQ}-before`, undefined);
  answer(f, `${REQ}-before`);
  send(f, `${REQ}-after`, undefined); // ExtraInfo after the main event
  extraInfo(f, `${REQ}-after`, T3);
  answer(f, `${REQ}-after`);
  // Redirect: the final hop carries its own bearer; later ExtraInfo for that id is discarded.
  send(f, `${REQ}-redir`, T4, { redirectResponse: { status: 302, url: "https://evil.example/r" } });
  extraInfo(f, `${REQ}-redir`, T4);
  answer(f, `${REQ}-redir`);
  // Rejected candidate (wrong tenant): ordinary D44 reason only.
  send(f, `${REQ}-bad`, T_BAD);
  answer(f, `${REQ}-bad`);
  navigate(f, PAGE_CANONICAL);
  navigate(f, PAGE_CANONICAL); // same shape again: no duplicate line
  navigate(f, PAGE_LOCALIZED);
  await flush();
}

function lineAfter(lines: string[], prefix: string): string[] {
  return lines.filter((l) => l.startsWith(`${ACCEPTANCE_PREFIX}${prefix}`));
}

describe("acceptance diagnostics: full fake-CDP flow (enabled)", () => {
  async function enabledRun(): Promise<{ lines: string[]; f: Flow }> {
    const { lines, acc } = collector();
    const f = await startFlow(acc);
    await runScript(f);
    return { lines, f };
  }

  it("emits only allowlisted lines and none of the canary material", async () => {
    const { lines } = await enabledRun();
    expect(lines.length).toBeGreaterThan(10);
    expectAllowed(lines);

    const dump = lines.join("\n");
    const forbidden = [
      TENANT,
      OTHER_TENANT,
      CLIENT,
      OTHER_CLIENT,
      SCOPE,
      "Canary.Scope.Two",
      SESSION,
      TARGET,
      REQ,
      "SUBCANARY5K",
      "KIDCANARY2P",
      "SIGCANARY8M",
      "COOKIECANARY",
      "PAGEQ7",
      "GRAPHQ7",
      "de-de",
      "evil.example",
      "microsoft.com",
      "://",
      "Bearer",
      "Authorization",
      "eyJ",
      "/v1.0",
      "$search",
    ];
    for (const needle of forbidden) expect(dump, needle).not.toContain(needle);
    for (const token of [T1, T2, T3, T4, T_BAD]) {
      expect(dump).not.toContain(token);
      for (const segment of token.split(".")) expect(dump).not.toContain(segment);
    }
  });

  it("makes audience, client claim, version, ordering, redirects and page shape observable", async () => {
    const { lines } = await enabledRun();
    const shapes = lineAfter(lines, "token_shape");
    expect(shapes).toEqual([
      "acceptance: token_shape audience=graph_url client_claim=appid token_version=v1 scp=present roles=absent tenant=matched remaining_minutes=60",
      "acceptance: token_shape audience=graph_url_trailing_slash client_claim=azp token_version=v2 scp=present roles=absent tenant=matched remaining_minutes=120",
      "acceptance: token_shape audience=graph_application_id client_claim=both token_version=other_or_absent scp=present roles=absent tenant=matched remaining_minutes=61",
      "acceptance: token_shape audience=graph_url client_claim=appid token_version=v1 scp=present roles=absent tenant=matched remaining_minutes=60",
    ]);
    // The wrong-tenant candidate produced no shape: 5 candidates, 4 shapes.
    expect(lineAfter(lines, "graph_request_exact_origin_match")).toHaveLength(5);
    expect(new Set(lineAfter(lines, "graph_request_exact_origin_match"))).toEqual(
      new Set([
        "acceptance: graph_request_exact_origin_match=yes graph_request_scheme=https graph_request_host=expected_graph_host graph_request_port=default_https",
      ]),
    );
    expect(lineAfter(lines, "authorization_observation")).toEqual([
      "acceptance: authorization_observation=main_event",
      "acceptance: authorization_observation=extra_info_before_main",
      "acceptance: authorization_observation=extra_info_after_main",
      "acceptance: authorization_observation=main_event", // redirect final hop
      "acceptance: authorization_observation=main_event", // wrong-tenant candidate
    ]);
    // Item 4: ExtraInfo carrying Authorization is visible even when the main event had it first.
    expect(lineAfter(lines, "extra_info_authorization")).toEqual([
      "acceptance: extra_info_authorization=present", // main + later ExtraInfo
      "acceptance: extra_info_authorization=present", // ExtraInfo before main
      "acceptance: extra_info_authorization=present", // ExtraInfo after main
      "acceptance: extra_info_authorization=not_observed", // redirected: ExtraInfo discarded
      "acceptance: extra_info_authorization=not_observed", // wrong-tenant: no ExtraInfo sent
    ]);
    expect(lineAfter(lines, "redirect_")).toEqual([
      "acceptance: redirect_seen=yes",
      "acceptance: redirect_extra_info_discarded=yes",
      "acceptance: redirect_final_hop_authorization=main_event", // emitted once the response proves it final
    ]);
    expect(lineAfter(lines, "graph_explorer_page_shape")).toEqual([
      "acceptance: graph_explorer_page_shape=localized locale_segment_shape=xx-xx", // attach
      "acceptance: graph_explorer_page_shape=canonical",
      "acceptance: graph_explorer_page_shape=localized locale_segment_shape=xx-xx",
    ]);
  });

  it("reports a redirect that leaves the Graph origin without its destination", async () => {
    const { lines, acc } = collector();
    const f = await startFlow(acc);
    send(f, `${REQ}-away`, T1);
    f.peer.event(
      "Network.requestWillBeSent",
      {
        requestId: `${REQ}-away`,
        redirectResponse: { status: 302 },
        request: { url: "https://login.example.invalid/DESTCANARY", headers: {} },
      },
      SESSION,
    );
    await flush();
    expect(lines).toContain("acceptance: redirect_seen=yes");
    expect(lines.join("\n")).not.toContain("DESTCANARY");
    expect(lines.join("\n")).not.toContain("invalid");
  });

  it("re-reports the page shape after navigating away and back (checklist 4b)", async () => {
    const { lines, acc } = collector();
    const f = await startFlow(acc, PAGE_CANONICAL);
    navigate(f, "https://www.microsoft.com/");
    await flush();
    navigate(f, PAGE_LOCALIZED);
    await flush();
    expect(lineAfter(lines, "graph_explorer_page_shape")).toEqual([
      "acceptance: graph_explorer_page_shape=canonical",
      "acceptance: graph_explorer_page_shape=localized locale_segment_shape=xx-xx",
    ]);
    expect(lines.join("\n")).not.toContain("www.microsoft.com");
  });

  it("writes nothing to the file system", async () => {
    const spies = [
      vi.spyOn(fs, "writeFileSync"),
      vi.spyOn(fs, "appendFileSync"),
      vi.spyOn(fs, "writeSync"),
      vi.spyOn(fs, "createWriteStream"),
      vi.spyOn(fs, "openSync"),
      vi.spyOn(fs.promises, "writeFile"),
      vi.spyOn(fs.promises, "appendFile"),
    ];
    try {
      await enabledRun();
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      for (const spy of spies) spy.mockRestore();
    }
  });
});

describe("acceptance diagnostics: ordering and redirect edge cases (PR #44 review)", () => {
  it("labels ExtraInfo-before-main by arrival order even when the main event also has a bearer", async () => {
    const { lines, acc } = collector();
    const f = await startFlow(acc);
    extraInfo(f, `${REQ}-both`, T1); // ExtraInfo first
    send(f, `${REQ}-both`, T1); // main event carries the bearer too
    answer(f, `${REQ}-both`);
    await flush();
    expect(lineAfter(lines, "authorization_observation")).toEqual([
      "acceptance: authorization_observation=extra_info_before_main",
    ]);
    expect(lineAfter(lines, "extra_info_authorization")).toEqual([
      "acceptance: extra_info_authorization=present",
    ]);
    expect(f.cap.current().state).toBe("ok"); // capture behavior unchanged
  });

  it("reports an ExtraInfo bearer that arrives only after the response", async () => {
    const { lines, acc } = collector();
    const f = await startFlow(acc);
    send(f, `${REQ}-late`, T1);
    answer(f, `${REQ}-late`); // evaluated here
    extraInfo(f, `${REQ}-late`, T1); // late ExtraInfo
    extraInfo(f, `${REQ}-late`, T1); // reported once only
    extraInfo(f, `${REQ}-never-seen`, T1); // unknown id: buffered, no line
    await flush();
    expect(lineAfter(lines, "extra_info_authorization")).toEqual([
      "acceptance: extra_info_authorization=not_observed",
      "acceptance: extra_info_authorization=present_after_response",
    ]);
    expectAllowed(lines);
  });

  it("forgets evaluated ids when the target is lost, so a late ExtraInfo is not reported", async () => {
    const { lines, acc } = collector();
    const f = await startFlow(acc);
    send(f, `${REQ}-gone`, T1);
    answer(f, `${REQ}-gone`);
    navigate(f, "https://www.microsoft.com/"); // target lost: request state is cleared
    await flush();
    navigate(f, PAGE_CANONICAL); // re-attached (the fake peer reuses the session id)
    await flush();
    extraInfo(f, `${REQ}-gone`, T1); // same request id in the new attachment
    await flush();
    expect(lineAfter(lines, "extra_info_authorization=present_after_response")).toEqual([]);
  });

  it("claims a final redirect hop only once its response arrives", async () => {
    const { lines, acc } = collector();
    const f = await startFlow(acc);
    const redirect = { redirectResponse: { status: 302, url: "https://graph.microsoft.com/x" } };
    send(f, `${REQ}-multi`, T1, redirect); // hop B: redirected, own bearer
    send(f, `${REQ}-multi`, undefined, redirect); // hop C: redirected again, no bearer
    answer(f, `${REQ}-multi`); // final response, but the final hop had no bearer
    await flush();
    expect(lineAfter(lines, "redirect_seen")).toHaveLength(2);
    expect(lineAfter(lines, "redirect_final_hop_authorization")).toEqual([]);

    send(f, `${REQ}-single`, T1, redirect); // a final hop that carries its own bearer
    answer(f, `${REQ}-single`);
    await flush();
    expect(lineAfter(lines, "redirect_final_hop_authorization")).toEqual([
      "acceptance: redirect_final_hop_authorization=main_event",
    ]);
  });
});

describe("acceptance diagnostics: observation only (disabled equals enabled behavior)", () => {
  const trap: AcceptanceDiagnostics = (() => {
    const boom = (): never => {
      throw new Error("diagnostics invoked while disabled");
    };
    return {
      enabled: false,
      modeEnabled: boom,
      tokenShape: boom,
      graphRequestOrigin: boom,
      graphExplorerPage: boom,
      authorizationObservation: boom,
      extraInfoAuthorization: boom,
      redirectSeen: boom,
      redirectFinalHopAuthorizationFromMainEvent: boom,
      redirectExtraInfoDiscarded: boom,
    };
  })();

  async function observe(acc: AcceptanceDiagnostics | undefined) {
    const f = await startFlow(acc);
    await runScript(f);
    return {
      diags: f.diags,
      current: f.cap.current(),
      cdp: f.peer.received.map(
        (m) => `${m.method}|${m.sessionId ?? ""}|${JSON.stringify(m.params)}`,
      ),
    };
  }

  it("capture behavior, diagnostics and CDP traffic are identical with the mode off or on", async () => {
    const off = await observe(undefined);
    const noop = await observe(DISABLED_ACCEPTANCE_DIAGNOSTICS);
    const gated = await observe(trap); // would throw if any observation ran while disabled
    const on = await observe(collector().acc);
    expect(noop).toEqual(off);
    expect(gated).toEqual(off);
    expect(on).toEqual(off);
    expect(off.current).toMatchObject({ state: "ok" });
  });
});
