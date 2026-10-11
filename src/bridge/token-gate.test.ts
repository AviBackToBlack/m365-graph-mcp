import { describe, expect, it } from "vitest";
import {
  evaluateCandidate,
  isExactGraphOrigin,
  MIN_REMAINING_MS,
  NBF_SKEW_MS,
} from "./token-gate.js";
import type { TokenCandidate, TokenGateConfig } from "./token-gate.js";

const TENANT = "11111111-2222-3333-4444-555555555555";
const OTHER_TENANT = "99999999-8888-7777-6666-555555555555";
const CLIENT = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
const NOW = 1_800_000_000_000;
const GRAPH_URL = "https://graph.microsoft.com/v1.0/me";

function jwt(claims: Record<string, unknown>): string {
  const enc = (o: unknown): string => Buffer.from(JSON.stringify(o)).toString("base64url");
  return `${enc({ alg: "RS256", typ: "JWT" })}.${enc(claims)}.c2ln`;
}

function goodClaims(over: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    aud: "https://graph.microsoft.com",
    tid: TENANT,
    scp: "Mail.Read User.Read",
    exp: Math.floor((NOW + 3_600_000) / 1000),
    ver: "1.0",
    appid: CLIENT,
    ...over,
  };
}

function cand(over: Partial<TokenCandidate> = {}, claims = goodClaims()): TokenCandidate {
  return {
    bearer: jwt(claims),
    requestUrl: GRAPH_URL,
    responseStatus: 200,
    sourceIsBridgeTarget: true,
    ...over,
  };
}

const CFG: TokenGateConfig = { tenantId: TENANT };
const reasonOf = (c: TokenCandidate, cfg = CFG, now = NOW): string | "ok" => {
  const r = evaluateCandidate(c, cfg, now);
  return r.ok ? "ok" : r.reason;
};

describe("token gate: source and origin", () => {
  it("accepts a good candidate and reports expiry", () => {
    const r = evaluateCandidate(cand(), CFG, NOW);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.expMs).toBe(Math.floor((NOW + 3_600_000) / 1000) * 1000);
      expect(r.remainingMs).toBe(3_600_000);
    }
  });

  it("rejects non-bridge sources", () => {
    expect(reasonOf(cand({ sourceIsBridgeTarget: false }))).toBe("wrong_source");
  });

  it.each([
    "https://graph.microsoft.com.evil.com/v1.0/me",
    "https://evilgraph.microsoft.com/v1.0/me",
    "https://graph.microsoft.com:8443/v1.0/me",
    "http://graph.microsoft.com/v1.0/me",
    "https://user:pw@graph.microsoft.com/v1.0/me",
    "https://user@graph.microsoft.com/v1.0/me",
    "https://graph.microsoft.com@evil.com/v1.0/me",
    "https://graph.microsoft.com./v1.0/me",
    "https://login.microsoftonline.com/common/oauth2/token",
    "https://developer.microsoft.com/en-us/graph/graph-explorer",
    "https://graph.microsoft.us/v1.0/me",
    "not a url",
    "",
  ])("rejects origin %s", (url) => {
    expect(reasonOf(cand({ requestUrl: url }))).toBe("wrong_origin");
  });

  it.each([
    "https://graph.microsoft.com/v1.0/me",
    "https://GRAPH.Microsoft.COM/v1.0/me",
    "https://graph.microsoft.com:443/v1.0/me",
  ])("accepts the exact origin %s", (url) => {
    expect(isExactGraphOrigin(url)).toBe(true);
  });

  it("rejects 401 but not 403, and requires a response", () => {
    expect(reasonOf(cand({ responseStatus: 401 }))).toBe("http_401");
    expect(reasonOf(cand({ responseStatus: 403 }))).toBe("ok");
    expect(reasonOf(cand({ responseStatus: undefined }))).toBe("no_response");
  });
});

describe("token gate: token format and claims", () => {
  it.each([
    ["opaque", "EwBwA8l6BAAU-opaque-token"],
    ["malformed", "a.b"],
    [
      "non-json payload",
      `${Buffer.from("{}").toString("base64url")}.${Buffer.from("x").toString("base64url")}.c2ln`,
    ],
    ["array payload", `e30.${Buffer.from("[]").toString("base64url")}.c2ln`],
    ["whitespace", `${jwt(goodClaims())} x`],
    ["empty", ""],
    ["too long", "a".repeat(20_000)],
  ])("rejects %s tokens as unsupported", (_n, bearer) => {
    expect(reasonOf(cand({ bearer }))).toBe("unsupported_token_format");
  });

  it.each([
    ["https://graph.microsoft.com", true],
    ["https://graph.microsoft.com/", true],
    ["00000003-0000-0000-c000-000000000000", true],
    ["https://graph.microsoft.com/v1.0", false],
    ["https://graph.microsoft.us", false],
    ["https://outlook.office.com", false],
  ])("audience %s accepted=%s", (aud, ok) => {
    expect(reasonOf(cand({}, goodClaims({ aud })))).toBe(ok ? "ok" : "audience");
  });

  it("rejects array, missing audiences", () => {
    expect(reasonOf(cand({}, goodClaims({ aud: ["https://graph.microsoft.com"] })))).toBe(
      "audience",
    );
    const c = goodClaims();
    delete c["aud"];
    expect(reasonOf(cand({}, c))).toBe("audience");
  });

  it("pins the tenant exactly (case-insensitive GUID), requires tid", () => {
    expect(reasonOf(cand({}, goodClaims({ tid: OTHER_TENANT })))).toBe("tenant");
    expect(reasonOf(cand({}, goodClaims({ tid: TENANT.toUpperCase() })))).toBe("ok");
    const c = goodClaims();
    delete c["tid"];
    expect(reasonOf(cand({}, c))).toBe("tenant");
    expect(reasonOf(cand({}, goodClaims({ tid: undefined, upn: "u@contoso.com" })))).toBe("tenant");
  });

  it("requires a non-empty delegated scp and rejects app roles", () => {
    expect(reasonOf(cand({}, goodClaims({ scp: "" })))).toBe("not_delegated");
    expect(reasonOf(cand({}, goodClaims({ scp: "  " })))).toBe("not_delegated");
    expect(reasonOf(cand({}, goodClaims({ scp: ["Mail.Read"] })))).toBe("not_delegated");
    const c = goodClaims();
    delete c["scp"];
    expect(reasonOf(cand({}, { ...c, roles: ["Mail.Read"] }))).toBe("not_delegated");
    expect(reasonOf(cand({}, goodClaims({ roles: ["Mail.Read"] })))).toBe("not_delegated");
    expect(reasonOf(cand({}, goodClaims({ roles: "x" })))).toBe("not_delegated");
    // D44 delegated-only profile: roles must be absent; even an empty array is refused.
    expect(reasonOf(cand({}, goodClaims({ roles: [] })))).toBe("not_delegated");
  });

  it("enforces the 5 minute minimum remaining lifetime at the boundary", () => {
    const expAt = (ms: number) => goodClaims({ exp: (NOW + ms) / 1000 });
    expect(reasonOf(cand({}, expAt(MIN_REMAINING_MS)))).toBe("ok");
    expect(reasonOf(cand({}, expAt(MIN_REMAINING_MS - 1000)))).toBe("lifetime");
    expect(reasonOf(cand({}, expAt(-1000)))).toBe("lifetime");
    expect(reasonOf(cand({}, goodClaims({ exp: "9999999999" })))).toBe("lifetime");
    expect(reasonOf(cand({}, goodClaims({ exp: 1e300 })))).toBe("lifetime");
    const c = goodClaims();
    delete c["exp"];
    expect(reasonOf(cand({}, c))).toBe("lifetime");
  });

  it("applies the optional nbf check with small skew", () => {
    const nbfAt = (ms: number) => goodClaims({ nbf: (NOW + ms) / 1000 });
    expect(reasonOf(cand({}, nbfAt(NBF_SKEW_MS)))).toBe("ok");
    expect(reasonOf(cand({}, nbfAt(NBF_SKEW_MS + 1000)))).toBe("not_yet_valid");
    expect(reasonOf(cand({}, nbfAt(-10_000)))).toBe("ok");
    expect(reasonOf(cand({}, goodClaims({ nbf: "soon" })))).toBe("lifetime");
  });

  it("rejects an invalid clock", () => {
    expect(reasonOf(cand(), CFG, Number.NaN)).toBe("clock_invalid");
  });
});

describe("token gate: client pin", () => {
  const pinned = (claim: "auto" | "appid" | "azp"): TokenGateConfig => ({
    tenantId: TENANT,
    clientPin: { guid: CLIENT, claim },
  });

  it("is off by default (no appid/azp needed)", () => {
    const c = goodClaims();
    delete c["appid"];
    expect(reasonOf(cand({}, c))).toBe("ok");
  });

  it("auto: v1 uses appid, v2 uses azp", () => {
    expect(reasonOf(cand({}, goodClaims()), pinned("auto"))).toBe("ok");
    const v2 = goodClaims({ ver: "2.0", azp: CLIENT.toUpperCase() });
    delete v2["appid"];
    expect(reasonOf(cand({}, v2), pinned("auto"))).toBe("ok");
    // v1 token carrying only azp, or v2 token carrying only appid, is rejected in auto mode.
    const v1WrongClaim = goodClaims({ azp: CLIENT });
    delete v1WrongClaim["appid"];
    expect(reasonOf(cand({}, v1WrongClaim), pinned("auto"))).toBe("client_pin");
    expect(reasonOf(cand({}, goodClaims({ ver: "2.0" })), pinned("auto"))).toBe("client_pin");
  });

  it("rejects absent claim, mismatch, and unknown ver", () => {
    const absent = goodClaims();
    delete absent["appid"];
    expect(reasonOf(cand({}, absent), pinned("auto"))).toBe("client_pin");
    expect(reasonOf(cand({}, goodClaims({ appid: OTHER_TENANT })), pinned("auto"))).toBe(
      "client_pin",
    );
    expect(reasonOf(cand({}, goodClaims({ ver: undefined })), pinned("auto"))).toBe("client_pin");
    expect(reasonOf(cand({}, goodClaims({ ver: "3.0" })), pinned("auto"))).toBe("client_pin");
  });

  it("explicit claim selection", () => {
    expect(reasonOf(cand({}, goodClaims()), pinned("appid"))).toBe("ok");
    expect(reasonOf(cand({}, goodClaims()), pinned("azp"))).toBe("client_pin");
    expect(reasonOf(cand({}, goodClaims({ azp: CLIENT })), pinned("azp"))).toBe("ok");
  });
});

describe("token gate: no secret material in rejections", () => {
  it("never leaks canaries through any rejection result", () => {
    const canaryTenant = "deadbeef-dead-beef-dead-beefdeadbeef";
    const canary = "CANARY-SCOPE-VALUE";
    const variants: TokenCandidate[] = [
      cand({}, goodClaims({ tid: canaryTenant, scp: canary })),
      cand({}, goodClaims({ aud: `https://${canary}.example`, scp: canary })),
      cand({}, goodClaims({ scp: "", appid: canaryTenant })),
      cand({}, goodClaims({ roles: [canary] })),
      cand({}, goodClaims({ appid: canaryTenant })),
      cand({ requestUrl: `https://${canary}.evil.com/` }),
      cand({ bearer: `opaque-${canary}` }),
    ];
    const cfg: TokenGateConfig = { tenantId: TENANT, clientPin: { guid: CLIENT, claim: "auto" } };
    for (const v of variants) {
      const json = JSON.stringify(evaluateCandidate(v, cfg, NOW));
      expect(json).not.toContain(canary);
      expect(json).not.toContain(canaryTenant);
      expect(json).not.toContain(TENANT);
      expect(json).not.toContain(CLIENT);
      expect(json).not.toContain(v.bearer);
      expect(json).toMatch(/^\{"ok":false,"reason":"[a-z0-9_]+"\}$/);
    }
  });
});
