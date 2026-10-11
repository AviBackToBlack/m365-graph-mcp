import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";
import { sanitizePresence, sanitizeTrending } from "./presence-insights.js";

/**
 * SM-14 part A: get-my-presence, get-user-presence (GUID-only, D39) and list-trending-insights
 * (D40). The shared 401/403/token-leak matrix lives in presence-insights-capability.test.ts.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const MY = "get-my-presence";
const USER = "get-user-presence";
const TREND = "list-trending-insights";
const GUID = "f17e0000-0000-4000-8000-000000000006";
const OTHER_GUID = "f17e0000-0000-4000-8000-000000000019";
const enc = encodeURIComponent;

interface Call {
  url: string;
  method: string;
  headers: Record<string, string>;
}

let catalog: EndpointCatalog;
let mcp: Client;
let calls: Call[];
let responses: Array<() => Response>;

const json = (body: unknown, status = 200): Response =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });

type Env = {
  ok: boolean;
  data?: unknown;
  meta?: Record<string, unknown>;
  error?: { code: string; message: string; details?: Record<string, unknown> };
};
const call = async (name: string, args: Record<string, unknown> = {}) => {
  const r = await mcp.callTool({ name, arguments: args });
  return { r, env: r.structuredContent as Env, text: JSON.stringify(r) };
};
const parsed = (i = 0) => new URL((calls[i] as Call).url);

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  const fetch: FetchLike = (url, init) => {
    calls.push({
      url,
      method: String(init.method),
      headers: { ...(init.headers as Record<string, string>) },
    });
    const next = responses.shift();
    return Promise.resolve(next ? next() : json({ value: [] }));
  };
  const graph = new GraphClient({
    tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
    catalog,
    fetch,
    sleep: () => Promise.resolve(),
  });
  const server = createServer(catalog, createHandlers(graph, catalog));
  const [ct, st] = InMemoryTransport.createLinkedPair();
  mcp = new Client({ name: "t", version: "0" });
  await Promise.all([server.connect(st), mcp.connect(ct)]);
});
afterAll(async () => {
  await mcp.close();
});
beforeEach(() => {
  calls = [];
  responses = [];
});

// --- Fixtures: Graph shapes with extras that must be dropped ----------------------------------

const rawPresence = (over: Record<string, unknown> = {}) => ({
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#presences/$entity",
  "@odata.type": "#microsoft.graph.presence",
  id: GUID,
  availability: "Busy",
  activity: "InAMeeting",
  sequenceNumber: "A0129311063-sequence-secret",
  statusMessage: {
    message: { content: "status-message-secret", contentType: "text" },
    publishedDateTime: "2026-10-07T10:00:00Z",
  },
  outOfOfficeSettings: { message: "ooo-message-secret", isOutOfOffice: false },
  workLocation: {
    workLocationType: "office",
    source: "source-secret",
    placeId: "place-secret",
  },
  unknownThing: "unknown-secret",
  ...over,
});
const projectedPresence = {
  id: GUID,
  availability: "Busy",
  activity: "InAMeeting",
  outOfOfficeSettings: { isOutOfOffice: false },
  workLocation: { workLocationType: "office" },
};
const PRESENCE_SECRETS =
  /secret|sequenceNumber|statusMessage|placeId|source|@odata|publishedDateTime|unknownThing/;

const rawTrending = (over: Record<string, unknown> = {}) => ({
  "@odata.type": "#microsoft.graph.trending",
  id: "AWMiSOpKHlJCpP_trending-id",
  weight: 0.1583399742569597,
  resourceVisualization: {
    title: "LiveCaptions",
    type: "Image",
    mediaType: "application/octet-stream",
    previewImageUrl: "https://contoso.sharepoint.com/preview-secret",
    previewText: "preview-text-secret",
    containerWebUrl: "https://contoso.sharepoint.com/sites/Mark8ProjectTeam/Shared Documents",
    containerDisplayName: "Mark 8 Project Team",
    containerType: "Site",
    extraViz: "viz-secret",
  },
  resourceReference: {
    webUrl: "https://contoso.sharepoint.com/sites/Mark8ProjectTeam/Shared%20Documents/x.gif",
    id: "drives/b!resource-id-secret/items/01H273",
    type: "microsoft.graph.driveItem",
    extraRef: "ref-secret",
  },
  unknownThing: "unknown-secret",
  ...over,
});
const projectedTrending = {
  id: "AWMiSOpKHlJCpP_trending-id",
  weight: 0.1583399742569597,
  resourceVisualization: {
    title: "LiveCaptions",
    type: "Image",
    mediaType: "application/octet-stream",
    containerWebUrl: "https://contoso.sharepoint.com/sites/Mark8ProjectTeam/Shared Documents",
    containerDisplayName: "Mark 8 Project Team",
    containerType: "Site",
  },
  resourceReference: {
    webUrl: "https://contoso.sharepoint.com/sites/Mark8ProjectTeam/Shared%20Documents/x.gif",
    type: "microsoft.graph.driveItem",
  },
};
const TRENDING_SECRETS =
  /secret|preview|resource-id|"id":"drives|@odata|extraViz|extraRef|unknownThing/;

describe("catalog and surface", () => {
  it("matches the catalog entries (v1.0 GET) with the scopes recorded in D39/D40", () => {
    for (const [name, pattern, scopes] of [
      [MY, "/me/presence", ["Presence.Read"]],
      [USER, "/users/{user-id}/presence", ["Presence.Read.All"]],
      [TREND, "/me/insights/trending", ["Sites.Read.All"]],
    ] as const) {
      const e = catalog.find((x) => x.toolName === name);
      expect(e?.method, name).toBe("get");
      expect(e?.pathPattern, name).toBe(pattern);
      expect(e?.["apiVersion"], name).not.toBe("beta");
      expect(e?.["workScopes"], name).toEqual(scopes);
    }
  });

  it("exposes only the typed inputs (no query options; user-id does not advertise a UPN)", async () => {
    const tools = (await mcp.listTools()).tools;
    const schema = (n: string) => tools.find((t) => t.name === n)?.inputSchema;
    const props = (n: string) =>
      Object.keys((schema(n)?.properties ?? {}) as Record<string, unknown>).sort();
    expect(props(MY)).toEqual([]);
    expect(props(USER)).toEqual(["user-id"]);
    expect(schema(USER)?.required).toEqual(["user-id"]);
    expect(props(TREND)).toEqual(["top"]);
    const top = (schema(TREND)?.properties as Record<string, Record<string, unknown>>)["top"];
    expect(top).toMatchObject({ type: "integer", minimum: 1, maximum: 50 });
    expect(JSON.stringify(schema(USER))).not.toMatch(/UPN|email|userPrincipalName/i);
  });

  it("describes the typed contracts, not raw Graph syntax (D39, D40)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    const tip = (n: string) => String(catalog.find((x) => x.toolName === n)?.["llmTip"]);
    const all = [MY, USER, TREND].map((n) => desc(n) + tip(n)).join("\n");
    // The previous trending tip prescribed raw OData; no ordering is advertised anywhere, and
    // nothing claims Microsoft lacks $filter (it documents it; this surface omits it deliberately).
    expect(all).not.toMatch(/\$orderby|orderby|\$filter=|PowerPoint|get-user\b/);
    expect(all).not.toMatch(
      /doesn't support|does not support \$filter|not supported by Microsoft/i,
    );
    for (const n of [MY, USER]) {
      expect(desc(n)).toMatch(
        /Returns only id, availability, activity, outOfOfficeSettings\.isOutOfOffice and workLocation\.workLocationType/,
      );
      expect(desc(n)).toMatch(
        /the status message, the out-of-office message, work-location place and source, sequenceNumber and unknown properties are not exposed/,
      );
      expect(desc(n)).toMatch(/work\/school accounts only/);
      expect(desc(n)).toMatch(/Takes no query options and no caller-supplied headers/);
      expect(desc(n)).toMatch(/normal error mapping/);
      expect(tip(n)).toMatch(
        /Status and out-of-office messages are not returned|status message and out-of-office message are not returned/,
      );
    }
    expect(desc(MY)).toMatch(/Requires Presence\.Read \(delegated\)/);
    expect(desc(USER)).toMatch(/Requires Presence\.Read\.All \(delegated\)/);
    expect(desc(USER)).toMatch(/user-id is that user's Microsoft Entra object ID GUID/);
    expect(desc(USER)).toMatch(
      /normally obtained from list-users or from the id of an already-returned user object/,
    );
    expect(desc(USER)).toMatch(/never constructed or inferred from an email address/);
    expect(desc(USER)).toMatch(/a UPN or email address is rejected locally with no Graph request/);
    expect(tip(USER)).toMatch(/Microsoft Entra object ID GUID, taken from list-users/);
    expect(tip(USER)).toMatch(/a UPN or email is refused locally/);
    expect(desc(MY) + tip(MY)).not.toMatch(/user-id|list-users/);

    expect(desc(TREND)).toMatch(
      /top is a local result limit \(1-50, default 25\): no Graph \$top is sent/,
    );
    expect(desc(TREND)).toMatch(
      /No raw OData options or caller-supplied headers; this tool does not expose a filter/,
    );
    expect(desc(TREND)).toMatch(
      /Each item returns only id, weight \(a relevance score\), resourceVisualization title, type, mediaType, containerDisplayName, containerType and containerWebUrl, and resourceReference webUrl and type/,
    );
    expect(desc(TREND)).toMatch(
      /previewImageUrl, previewText, resourceReference\.id \(a raw Graph resource path that is deliberately omitted and must not be used as a request target\) and unknown properties are not exposed/,
    );
    expect(desc(TREND)).toMatch(/untrusted tenant content/);
    expect(desc(TREND)).toMatch(/global service only \(not national clouds\)/);
    expect(desc(TREND)).toMatch(/work\/school accounts only/);
    expect(desc(TREND)).toMatch(/Requires Sites\.Read\.All \(delegated\)/);
    expect(tip(TREND)).toMatch(/top is a local limit \(1-50, default 25\)/);
    expect(tip(TREND)).toMatch(/No raw OData is accepted/);
    expect(tip(TREND)).toMatch(/raw resource id are intentionally omitted; use webUrl/);
    expect(tip(TREND)).toMatch(/Global service only; work\/school accounts only/);
  });

  it("every catalogued tool, GET and POST, has a handler (search-query: SM-14 part B, copilot-retrieve: part C)", () => {
    const handlers = createHandlers(
      new GraphClient({ tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) }, catalog }),
      catalog,
    );
    const unimplementedGets = catalog
      .filter((e) => e.method === "get" && !handlers.has(e.toolName))
      .map((e) => e.toolName);
    expect(unimplementedGets).toEqual([]);
    const unimplementedPosts = catalog
      .filter((e) => e.method === "post" && !handlers.has(e.toolName))
      .map((e) => e.toolName)
      .sort();
    expect(unimplementedPosts).toEqual([]);
  });
});

describe(MY, () => {
  it("sends exactly GET /v1.0/me/presence with no query and no extra headers, and projects", async () => {
    responses.push(() => json(rawPresence()));
    const { env, text } = await call(MY);
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe("/v1.0/me/presence");
    expect([...parsed().searchParams.keys()]).toEqual([]);
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(env.data).toEqual(projectedPresence);
    expect(text).not.toMatch(PRESENCE_SECRETS);
    expect(text).not.toContain(TOKEN);
  });

  it("rejects any option with zero requests", async () => {
    for (const extra of [
      { select: ["id"] },
      { $select: "id" },
      { top: 1 },
      { filter: "x" },
      { "user-id": GUID },
      { url: "/me/presence" },
    ]) {
      const { env } = await call(MY, extra);
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("fails closed on a non-object 200 body", async () => {
    for (const body of [[], "x", null, 5]) {
      calls = [];
      responses.push(() => json(body));
      const { env } = await call(MY);
      expect(env.ok, JSON.stringify(body)).toBe(false);
      expect(env.error?.code).toBe("graph_error");
    }
  });
});

describe(USER, () => {
  it("requests /v1.0/users/{guid}/presence (either GUID case) and projects like get-my-presence", async () => {
    for (const id of [GUID, GUID.toUpperCase()]) {
      calls = [];
      responses.push(() => json(rawPresence({ id })));
      const { env, text } = await call(USER, { "user-id": id });
      expect(env.ok, id).toBe(true);
      expect(parsed().pathname).toBe(`/v1.0/users/${id}/presence`);
      expect([...parsed().searchParams.keys()]).toEqual([]);
      expect(env.data).toEqual({ ...projectedPresence, id });
      expect(text).not.toMatch(PRESENCE_SECRETS);
    }
  });

  it("rejects a UPN, email or any non-GUID user-id locally with zero Graph requests", async () => {
    for (const v of [
      "alice@contoso.com",
      "alice%40contoso.com",
      "o'brien+tag@sub.contoso.co.uk",
      "bob_contoso.com#EXT#@t.onmicrosoft.com",
      "me",
      `${GUID}x`,
      `{${GUID}}`,
      GUID.replaceAll("-", ""),
      `${GUID}/presence`,
      `${GUID}/../${OTHER_GUID}`,
      enc(`${GUID}/messages`),
      "delta",
      "$count",
      "",
    ]) {
      const { env, text } = await call(USER, { "user-id": v });
      expect(env.ok, v).toBe(false);
      expect(env.error?.code, v).toBe("invalid_input");
      expect(text).not.toContain("contoso.com");
    }
    expect(calls).toHaveLength(0);
  });

  it("requires user-id and rejects options with zero requests", async () => {
    for (const args of [
      {},
      { "user-id": GUID, select: ["id"] },
      { "user-id": GUID, top: 1 },
      { "user-id": GUID, $select: "statusMessage" },
      { "user-id": GUID, "presence-id": GUID },
    ]) {
      const { env } = await call(USER, args);
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("fails closed on a non-object 200 body", async () => {
    responses.push(() => json([]));
    const { env } = await call(USER, { "user-id": GUID });
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("graph_error");
  });
});

describe("sanitizePresence", () => {
  it("keeps only the five allowlisted values and reduces the nested objects", () => {
    expect(sanitizePresence(rawPresence())).toEqual(projectedPresence);
  });

  it("omits a null, scalar or array nested object instead of passing it through", () => {
    for (const bad of [null, "x", 5, [], [{ isOutOfOffice: true }]]) {
      expect(
        sanitizePresence(rawPresence({ outOfOfficeSettings: bad, workLocation: bad })),
        JSON.stringify(bad),
      ).toEqual({ id: GUID, availability: "Busy", activity: "InAMeeting" });
    }
  });

  it("keeps only scalar values for allowlisted keys and nothing for objects in their place", () => {
    expect(
      sanitizePresence({
        id: GUID,
        availability: { nested: "x-secret" },
        activity: ["x-secret"],
        outOfOfficeSettings: { isOutOfOffice: { x: 1 } },
        workLocation: { workLocationType: ["x-secret"] },
      }),
    ).toEqual({ id: GUID, outOfOfficeSettings: {}, workLocation: {} });
  });

  it("returns undefined for non-objects", () => {
    for (const v of [null, undefined, "x", 1, []]) expect(sanitizePresence(v)).toBeUndefined();
  });
});

describe(TREND, () => {
  it("sends exactly GET /v1.0/me/insights/trending (no query, no $top) and projects", async () => {
    responses.push(() => json({ "@odata.context": "x", value: [rawTrending()] }));
    const { env, text } = await call(TREND);
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe("/v1.0/me/insights/trending");
    expect([...parsed().searchParams.keys()]).toEqual([]);
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(env.data).toEqual([projectedTrending]);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
    expect(text).not.toMatch(TRENDING_SECRETS);
    expect(text).not.toContain(TOKEN);
  });

  it("keeps weight whether Graph sends a number or a string (Learn's example is a string)", async () => {
    responses.push(() =>
      json({ value: [rawTrending({ weight: "0.1583399742569597" }), rawTrending({ weight: 2 })] }),
    );
    const { env } = await call(TREND);
    expect((env.data as Array<{ weight: unknown }>).map((i) => i.weight)).toEqual([
      "0.1583399742569597",
      2,
    ]);
  });

  it("top is a local limit: no $top on the wire, extra items clipped, hasMore reported", async () => {
    responses.push(() =>
      json({
        value: [rawTrending({ id: "a" }), rawTrending({ id: "b" }), rawTrending({ id: "c" })],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/insights/trending?$skiptoken=abc",
      }),
    );
    const { env } = await call(TREND, { top: 2 });
    expect(parsed().searchParams.get("$top")).toBeNull();
    expect((env.data as Array<{ id: string }>).map((s) => s.id)).toEqual(["a", "b"]);
    expect(env.meta).toMatchObject({ count: 2, hasMore: true });
    expect(calls).toHaveLength(1);
  });

  it("the default limit is 25 and the maximum is 50", async () => {
    const many = (n: number) =>
      Array.from({ length: n }, (_v, i) => rawTrending({ id: `i${String(i)}` }));
    responses.push(() => json({ value: many(60) }));
    const dflt = await call(TREND);
    expect(dflt.env.data).toHaveLength(25);
    expect(dflt.env.meta).toMatchObject({ count: 25, hasMore: true });
    responses.push(() => json({ value: many(60) }));
    const max = await call(TREND, { top: 50 });
    expect(max.env.data).toHaveLength(50);
    expect(parsed(1).searchParams.get("$top")).toBeNull();
  });

  it("rejects top outside 1-50 or not an integer, and every raw OData option, with zero requests", async () => {
    for (const bad of [
      { top: 0 },
      { top: 51 },
      { top: -1 },
      { top: 1.5 },
      { top: "5" },
      { top: null },
      { $top: 5 },
      { filter: "resourceVisualization/type eq 'PowerPoint'" },
      { $filter: "x" },
      { orderby: [{ property: "weight" }] },
      { $orderby: "weight desc" },
      { select: ["id"] },
      { search: "x" },
      { skip: 1 },
      { expand: ["resource"] },
      { "user-id": GUID },
    ]) {
      const { env } = await call(TREND, bad);
      expect(env.ok, JSON.stringify(bad)).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("follows nextLink pages on the same route and projects every page", async () => {
    responses.push(() =>
      json({
        value: [rawTrending({ id: "a" })],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/insights/trending?$skiptoken=abc",
      }),
    );
    responses.push(() => json({ value: [rawTrending({ id: "b" })] }));
    const { env, text } = await call(TREND, { top: 5 });
    expect((env.data as Array<{ id: string }>).map((s) => s.id)).toEqual(["a", "b"]);
    expect(calls).toHaveLength(2);
    expect(parsed(1).pathname).toBe("/v1.0/me/insights/trending");
    expect(parsed(1).searchParams.get("$skiptoken")).toBe("abc");
    expect(text).not.toMatch(TRENDING_SECRETS);
  });

  it("refuses a nextLink that leaves the route or widens the query (zero further requests)", async () => {
    const base = "https://graph.microsoft.com/v1.0";
    for (const next of [
      `${base}/users/${GUID}/insights/trending?$skiptoken=abc`,
      `${base}/me/insights/trending/abc/resource?$skiptoken=abc`,
      `${base}/me/insights/used?$skiptoken=abc`,
      `${base}/me/insights/trending?$skiptoken=abc&$filter=${enc("resourceVisualization/type eq 'PowerPoint'")}`,
      `${base}/me/insights/trending?$skiptoken=abc&$select=resourceReference`,
      `https://evil.example.com/v1.0/me/insights/trending?$skiptoken=abc`,
      `${base}/me/presence`,
      `${base}/beta/me/insights/trending?$skiptoken=abc`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [rawTrending({ id: "a" })], "@odata.nextLink": next }));
      const { env, text } = await call(TREND, { top: 5 });
      expect(env.ok, next).toBe(false);
      expect(env.error?.code, next).toBe("invalid_input");
      expect(calls, next).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("an empty result is a successful empty list; a missing value array or non-object is a failure", async () => {
    responses.push(() => json({ value: [] }));
    const empty = await call(TREND);
    expect(empty.env.ok).toBe(true);
    expect(empty.env.data).toEqual([]);
    for (const body of [{}, [], "x", { value: "x" }]) {
      responses.push(() => json(body));
      const bad = await call(TREND);
      expect(bad.env.ok, JSON.stringify(body)).toBe(false);
      expect(bad.env.error?.code).toBe("graph_error");
    }
  });

  it("drops non-object items instead of passing them through", async () => {
    responses.push(() => json({ value: [null, "x", 5, [], rawTrending({ id: "ok" })] }));
    const { env } = await call(TREND);
    expect((env.data as Array<{ id: string }>).map((s) => s.id)).toEqual(["ok"]);
  });
});

describe("sanitizeTrending", () => {
  it("keeps exactly the D40 allowlist", () => {
    expect(sanitizeTrending(rawTrending())).toEqual(projectedTrending);
  });

  it("omits null, scalar and array nested objects and non-scalar allowlisted values", () => {
    expect(
      sanitizeTrending(rawTrending({ resourceVisualization: null, resourceReference: [{}] })),
    ).toEqual({ id: projectedTrending.id, weight: projectedTrending.weight });
    expect(
      sanitizeTrending({
        id: { x: "secret" },
        weight: ["secret"],
        resourceVisualization: { title: { x: "secret" }, type: "Image" },
        resourceReference: { webUrl: ["secret"], type: "microsoft.graph.driveItem" },
      }),
    ).toEqual({
      resourceVisualization: { type: "Image" },
      resourceReference: { type: "microsoft.graph.driveItem" },
    });
  });

  it("returns undefined for non-objects", () => {
    for (const v of [null, undefined, "x", 1, []]) expect(sanitizeTrending(v)).toBeUndefined();
  });
});

describe("graph-batch reachability and projection parity", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const ok = (body: unknown) => () => json({ responses: [{ id: "1", status: 200, body }] });
  const bodyOf = (r: { env: Env }) =>
    (r.env.data as { responses: Array<{ body?: unknown }> }).responses[0]?.body;

  it("get-my-presence: batch projects exactly like the single path", async () => {
    responses.push(() => json(rawPresence()));
    const single = (await call(MY)).env.data;
    calls = [];
    responses.push(ok(rawPresence()));
    const b = await batch("/me/presence");
    expect(b.env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(bodyOf(b)).toEqual(single);
    expect(bodyOf(b)).toEqual(projectedPresence);
    expect(b.text).not.toMatch(PRESENCE_SECRETS);
  });

  it("get-user-presence: batch projects exactly like the single path for a GUID", async () => {
    responses.push(() => json(rawPresence()));
    const single = (await call(USER, { "user-id": GUID })).env.data;
    calls = [];
    responses.push(ok(rawPresence()));
    const b = await batch(`/users/${GUID}/presence`);
    expect(b.env.ok).toBe(true);
    expect(bodyOf(b)).toEqual(single);
    expect(b.text).not.toMatch(PRESENCE_SECRETS);
  });

  it("a non-object presence batch body becomes {} (never raw)", async () => {
    responses.push(ok([{ statusMessage: "status-message-secret" }]));
    const b = await batch("/me/presence");
    expect(bodyOf(b)).toEqual({});
    expect(b.text).not.toMatch(/secret/);
  });

  it("list-trending-insights: batch projects items like the single path and drops collection metadata", async () => {
    responses.push(() => json({ value: [rawTrending()] }));
    const single = (await call(TREND)).env.data;
    calls = [];
    responses.push(
      ok({
        "@odata.context": "ctx-secret",
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/insights/trending?$skiptoken=x",
        value: [rawTrending()],
      }),
    );
    const b = await batch("/me/insights/trending");
    expect(b.env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect((bodyOf(b) as { value: unknown }).value).toEqual(single);
    expect((bodyOf(b) as { hasMore?: boolean }).hasMore).toBe(true);
    expect(b.text).not.toMatch(TRENDING_SECRETS);
    expect(b.text).not.toContain("skiptoken");
  });

  it("trending batch bodies are clipped to the default limit with hasMore", async () => {
    responses.push(
      ok({
        value: Array.from({ length: 30 }, (_v, i) => rawTrending({ id: `i${String(i)}` })),
      }),
    );
    const b = await batch("/me/insights/trending");
    const body = bodyOf(b) as { value: unknown[]; hasMore?: boolean };
    expect(body.value).toHaveLength(25);
    expect(body.hasMore).toBe(true);
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const bad: Array<[string, Record<string, string>?]> = [
      // user-id must be a GUID in batch too (parity with the single path, D39)
      ["/users/alice@contoso.com/presence"],
      ["/users/alice%40contoso.com/presence"],
      ["/users('alice@contoso.com')/presence"],
      ["/users/me/presence"],
      [`/users/${GUID}x/presence`],
      [`/users/${GUID}/presence/extra`],
      // presence has no query options
      ["/me/presence?$select=statusMessage"],
      [`/users/${GUID}/presence?$select=statusMessage`],
      ["/me/presence?$expand=x"],
      ["/me/presence?$top=1"],
      ["/me/presence", { ConsistencyLevel: "eventual" }],
      ["/me/presence", { Prefer: 'outlook.timezone="UTC"' }],
      // other presence routes are not catalogued
      [`/communications/presences/${GUID}`],
      ["/communications/presences"],
      ["/communications/getPresencesByUserId"],
      ["/me/presence/setPresence"],
      ["/me/presence/statusMessage"],
      ["/beta/me/presence"],
      // trending: no option at all, no sub-routes, no other user
      ["/me/insights/trending?$top=5"],
      ["/me/insights/trending?$top=51"],
      ["/me/insights/trending?$filter=resourceVisualization/type eq 'PowerPoint'"],
      ["/me/insights/trending?$orderby=weight desc"],
      ["/me/insights/trending?$select=resourceReference"],
      ["/me/insights/trending?$count=true"],
      ["/me/insights/trending?$expand=resource"],
      ["/me/insights/trending?$skip=1"],
      ["/me/insights/trending?$search=x"],
      ["/me/insights/trending", { ConsistencyLevel: "eventual" }],
      ["/me/insights/trending/abc/resource"],
      ["/me/insights/trending/abc"],
      [`/users/${GUID}/insights/trending`],
      ["/me/insights/used"],
      ["/me/insights/shared"],
      ["/beta/me/insights/trending"],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(env.error?.code, url).toBe("invalid_input");
      expect(calls, url).toHaveLength(0);
    }
  });
});
