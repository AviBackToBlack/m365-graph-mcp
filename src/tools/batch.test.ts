import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import type { TokenProvider } from "../auth/token-provider.js";
import { TokenError } from "../auth/token-provider.js";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createBatchTool } from "./batch.js";
import { createHandlers } from "./index.js";
import {
  probePermittedSyntheticBatch,
  probeUnimplementedGet,
} from "./unimplemented-get-fixture.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";

let catalog: EndpointCatalog;
let mcp: Client;
let fetchCalls: { url: string; body: string }[];
let responder: (n: number) => Response;
let provider: TokenProvider;
let graph: GraphClient;

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  fetchCalls = [];
  provider = { getAccessToken: () => Promise.resolve(TOKEN) };
  const fetch: FetchLike = (url, init) => {
    fetchCalls.push({ url, body: String(init.body) });
    return Promise.resolve(responder(fetchCalls.length));
  };
  graph = new GraphClient({
    tokenProvider: { getAccessToken: () => provider.getAccessToken() },
    catalog,
    fetch,
    sleep: () => Promise.resolve(),
  });
  // Tool-level behaviour is tested with every catalog GET permitted; the production
  // wiring (implemented tools only) is covered separately below.
  const permitted = catalog.filter((e) => e.method === "get");
  const server = createServer(
    catalog,
    new Map([
      [
        "graph-batch",
        createBatchTool(
          graph,
          catalog,
          new Map(permitted.map((e) => [e, { check: () => undefined }])),
        ),
      ],
    ]),
  );
  const [ct, st] = InMemoryTransport.createLinkedPair();
  mcp = new Client({ name: "t", version: "0" });
  await Promise.all([server.connect(st), mcp.connect(ct)]);
});
afterAll(async () => {
  await mcp.close();
});

type Env = {
  ok: boolean;
  data?: { responses: { id: string; status: number }[] };
  error?: { code: string };
};
const call = async (args: Record<string, unknown>) => {
  const r = await mcp.callTool({ name: "graph-batch", arguments: args });
  return { r, env: r.structuredContent as Env, text: JSON.stringify(r) };
};
const ok = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

describe("graph-batch tool", () => {
  it("advertises a GET-only schema without write verbs", async () => {
    const tool = (await mcp.listTools()).tools.find((t) => t.name === "graph-batch");
    const schema = JSON.stringify(tool?.inputSchema);
    expect(schema).toContain('"GET"');
    expect(schema).not.toMatch(/POST|PATCH|PUT|DELETE|dependsOn|"body"/);
    expect(tool?.inputSchema.required).toEqual(["requests"]);
  });

  it("returns an envelope with per-request results and no token", async () => {
    fetchCalls.length = 0;
    responder = () =>
      ok({
        responses: [
          { id: "b", status: 200, body: { v: 2 } },
          { id: "a", status: 200, body: { v: 1 } },
        ],
      });
    const { r, env, text } = await call({
      requests: [
        { id: "a", method: "GET", url: "/me/messages?$top=1" },
        { id: "b", method: "GET", url: "/me/mailFolders" },
      ],
    });
    expect(r.isError).toBe(false);
    expect(env.ok).toBe(true);
    expect(env.data?.responses.map((x) => x.id)).toEqual(["a", "b"]);
    expect(text).not.toContain(TOKEN);
    expect(fetchCalls).toHaveLength(1);
  });

  it("rejects hostile input with invalid_input and no network or token use", async () => {
    fetchCalls.length = 0;
    const bad: Record<string, unknown>[] = [
      { requests: [{ id: "a", method: "DELETE", url: "/me/messages" }] },
      { requests: [{ id: "a", method: "GET", url: "https://graph.microsoft.com/v1.0/me" }] },
      { requests: [{ id: "a", method: "GET", url: "/$batch" }] },
      { requests: [{ id: "a", method: "GET", url: "/me/messages", body: {} }] },
      { requests: [{ id: "a", method: "GET", url: "/me/messages", dependsOn: ["a"] }] },
      {
        requests: [
          { id: "a", method: "GET", url: "/me/messages", headers: { Authorization: "Bearer x" } },
        ],
      },
      { requests: [] },
      { requests: [{ id: "a", method: "GET", url: "/me/messages" }], extra: 1 },
      {},
    ];
    for (const args of bad) {
      const { r, env, text } = await call(args);
      expect(r.isError).toBe(true);
      expect(env.error?.code).toBe("invalid_input");
      expect(text).not.toContain(TOKEN);
    }
    // Defense in depth: a loosened schema layer still cannot reach the network.
    expect(fetchCalls).toHaveLength(0);
  });

  it("surfaces graph and token failures as sanitized error results", async () => {
    responder = () => ok({ error: { code: "x", message: `echo ${TOKEN}` } }, 401);
    const a = await call({ requests: [{ id: "a", method: "GET", url: "/me/messages" }] });
    expect(a.env.error?.code).toBe("reauth_required");
    expect(a.text).not.toContain(TOKEN);

    responder = () => ok({ responses: [{ id: "a", status: 200, body: { t: TOKEN } }] });
    const b = await call({ requests: [{ id: "a", method: "GET", url: "/me/messages" }] });
    expect(b.r.isError).toBe(true);
    expect(b.env.error?.code).toBe("graph_error");
    expect(b.text).not.toContain(TOKEN);

    responder = () => ok({ responses: [{ id: "zzz", status: 200 }] });
    const c = await call({ requests: [{ id: "a", method: "GET", url: "/me/messages" }] });
    expect(c.env.error?.code).toBe("graph_error");

    provider = {
      getAccessToken: () => Promise.reject(new TokenError("reauth_required", "expired")),
    };
    const d = await call({ requests: [{ id: "a", method: "GET", url: "/me/messages" }] });
    expect(d.env.error?.code).toBe("reauth_required");
    provider = { getAccessToken: () => Promise.resolve(TOKEN) };
  });
});

describe("graph-batch argument errors", () => {
  it("never echoes a caller-chosen record key (e.g. a token used as a header name)", async () => {
    const { env, text } = await call({
      requests: [{ id: "1", method: "GET", url: "/me/messages", headers: { "opaque-token": 1 } }],
    });
    expect(env.error?.code).toBe("invalid_input");
    expect(text).not.toContain("opaque-token");
    expect(text).toContain("requests.0.headers.[redacted]");
  });
});

describe("graph-batch production wiring", () => {
  it("rejects catalog GETs whose tool is not implemented, with zero network calls", async () => {
    // After SM-14 part A every real catalog GET is implemented, so the property is proven on a
    // synthetic catalogued-but-unimplemented GET appended to a copy of the catalog (see
    // unimplemented-get-fixture.ts). "Catalogued" and "not implemented" are asserted explicitly so
    // that this stays distinct from the "not in the catalog" refusal below.
    const probe = await probeUnimplementedGet(catalog);
    expect(probe.catalogued).toBe(true);
    expect(probe.hasHandler).toBe(false);
    expect(probe.batch.ok).toBe(false);
    expect(probe.batch.error?.code).toBe("invalid_input");
    expect(probe.direct.error?.code).toBe("not_implemented");
    expect(probe.requests).toBe(0);
  });

  it("the same synthetic GET is accepted once explicitly permitted (so 'no handler' was the only reason)", async () => {
    const control = await probePermittedSyntheticBatch(catalog);
    expect(control).toEqual({ ok: true, requests: 1 });
  });

  it("rejects a GET that is not in the catalog at all, with zero network calls", async () => {
    const before = fetchCalls.length;
    const handlers = createHandlers(graph, catalog);
    const server = createServer(catalog, handlers);
    const [ct, st] = InMemoryTransport.createLinkedPair();
    const client = new Client({ name: "w", version: "0" });
    await Promise.all([server.connect(st), client.connect(ct)]);
    for (const url of ["/me/zz-synthetic-unimplemented", "/me/zz-not-in-catalog"]) {
      const r = await client.callTool({
        name: "graph-batch",
        arguments: { requests: [{ id: "1", method: "GET", url }] },
      });
      expect((r.structuredContent as Env).error?.code, url).toBe("invalid_input");
    }
    expect(fetchCalls.length).toBe(before);
    await client.close();
  });
});

describe("graph-batch catalog llmTip", () => {
  const tip = (): string => String(catalog.find((e) => e.toolName === "graph-batch")?.["llmTip"]);

  it("pins the real contract and does not advertise unsupported fields", () => {
    const t = tip();
    expect(t).not.toBe("undefined");
    // Advertised field syntax only; the bare words "body"/"dependsOn" appear in "no ..." wording.
    expect(t).not.toMatch(/\bbody\b\s*\??:/);
    expect(t).not.toMatch(/\bdependsOn\b\s*\??:/i);
    expect(t).not.toMatch(/headers\s*\??:\s*\{/);
    expect(t).not.toMatch(/'GET'\s*\|/);
    expect(t).not.toMatch(/POST|PATCH|DELETE|PUT/);
    expect(t).not.toMatch(/arbitrary order/i);
    expect(t).not.toContain("https://");
    expect(t).toMatch(/\bGET\b/);
    expect(t).toMatch(/relative/i);
    expect(t).toContain("ConsistencyLevel");
    expect(t).toContain("Prefer");
    expect(t).toContain("20");
    expect(t).toMatch(/no request body/i);
    expect(t).toMatch(/no dependsOn/i);
  });
});
