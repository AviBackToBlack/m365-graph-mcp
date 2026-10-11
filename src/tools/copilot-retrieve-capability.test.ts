import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";

/**
 * SM-14 part C permission/capability matrix and catalog text for `copilot-retrieve` (D41). Scopes
 * are unchanged and pinned; the catalog text describes the typed contract. A 403 (missing
 * Sites.Read.All, missing ExternalItem.Read.All, licensing) is a `capability_gap` after exactly ONE
 * request per dataSource with no retry or escalation; 401 is `reauth_required`; 400/404 are
 * `graph_error`; no token ever leaks. Documented in PERMISSIONS.md ("SM-14 part C").
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";

let catalog: EndpointCatalog;
let mcp: Client;
let calls: Array<{ url: string; method: string }>;
let responses: Array<() => Response>;

const json = (body: unknown, status = 200): Response =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });

type Env = {
  ok: boolean;
  data?: unknown;
  error?: { code: string; message: string; details?: Record<string, unknown> };
};
const retrieve = async (dataSource: string) => {
  const r = await mcp.callTool({
    name: "copilot-retrieve",
    arguments: { queryString: "vpn setup", dataSource },
  });
  return { env: r.structuredContent as Env, text: JSON.stringify(r) };
};

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  const fetch: FetchLike = (url, init) => {
    calls.push({ url, method: String(init.method) });
    const next = responses.shift();
    return Promise.resolve(next ? next() : json({ retrievalHits: [] }));
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

const SOURCES = ["sharePoint", "oneDriveBusiness", "externalItem"];

describe("copilot-retrieve catalog entry (D41)", () => {
  it("is a v1.0 POST entry with unchanged scopes and a body schema consistent with the typed inputs", () => {
    const e = catalog.find((x) => x.toolName === "copilot-retrieve");
    expect(e?.method).toBe("post");
    expect(e?.["apiVersion"]).not.toBe("beta");
    expect(e?.pathPattern).toBe("/copilot/retrieval");
    expect(e?.["workScopes"]).toEqual([
      ["Files.Read.All", "Sites.Read.All"],
      ["ExternalItem.Read.All"],
    ]);
    const schema = e?.["requestBodySchema"] as {
      required: string[];
      properties: Record<
        string,
        { maxLength?: number; maximum?: number; minimum?: number; enum?: string[] }
      >;
    };
    expect(schema.required).toEqual(["queryString", "dataSource"]);
    expect(Object.keys(schema.properties).sort()).toEqual([
      "dataSource",
      "filterExpression",
      "maximumNumberOfResults",
      "queryString",
      "resourceMetadata",
    ]);
    expect(schema.properties["queryString"]?.maxLength).toBe(1500);
    expect(schema.properties["dataSource"]?.enum).toEqual(SOURCES);
    expect(schema.properties["maximumNumberOfResults"]).toMatchObject({ minimum: 1, maximum: 25 });
  });

  it("pins the typed-contract catalog text", () => {
    const e = catalog.find((x) => x.toolName === "copilot-retrieve");
    const desc = String(e?.["descriptionOverride"]);
    const tip = String(e?.["llmTip"]);
    expect(desc).toMatch(/POST \/copilot\/retrieval, v1\.0/);
    expect(desc).toMatch(/semantic\/hybrid grounding/);
    expect(desc).toMatch(/not lexical KQL search \(use search-query for Microsoft Search\)/);
    expect(desc).toMatch(/permission-trimmed text extracts/);
    expect(desc).toMatch(/flat typed top-level arguments/);
    expect(desc).toMatch(/queryString \(required, not blank, at most 1500 characters/);
    expect(desc).toMatch(/exactly one of sharePoint, oneDriveBusiness or externalItem per call/);
    expect(desc).toMatch(
      /filterExpression \(optional KQL scoping string, at most 1000 characters\)/,
    );
    expect(desc).toMatch(/at most 20 plain retrievable metadata field names/);
    expect(desc).toMatch(
      /maximumNumberOfResults \(optional, 1-25; the service default is up to 25\)/,
    );
    expect(desc).toMatch(
      /dataSourceConfiguration \(connector connection ids\) and thumbnails are not supported/,
    );
    expect(desc).toMatch(/not a security boundary: Microsoft may silently ignore invalid KQL/);
    expect(desc).toMatch(
      /sharePoint and oneDriveBusiness each need both Files\.Read\.All and Sites\.Read\.All; externalItem needs ExternalItem\.Read\.All/,
    );
    expect(desc).toMatch(/never requests, selects or escalates consent/);
    expect(desc).toMatch(/capability_gap after one request/);
    expect(desc).toMatch(/Work\/school accounts only/);
    expect(desc).toMatch(/Microsoft 365 Copilot add-on license/);
    expect(desc).toMatch(/pay-as-you-go consumption \(preview\)/);
    expect(desc).toMatch(/tenant-level SharePoint and Copilot connector content but not OneDrive/);
    expect(desc).toMatch(/up to 200 requests per user per hour/);
    expect(desc).toMatch(/not a reason to change scopes/);
    expect(desc).toMatch(/An empty retrievalHits list means no relevant results/);
    expect(desc).toMatch(/untrusted data/);
    expect(tip).toMatch(/lexical Microsoft Search\/KQL/);
    expect(tip).toMatch(/Semantic\/hybrid retrieval/);
    expect(tip).toMatch(/one source per call/);
    expect(tip).toMatch(/results are unordered/);
    expect(tip).toMatch(/never rely on it to restrict access/);
    expect(tip).toMatch(/externalItem needs ExternalItem\.Read\.All/);
    expect(tip).toMatch(/not OneDrive/);
    for (const text of [desc, tip]) {
      expect(text).not.toMatch(/--extra-scopes/);
      expect(text).not.toMatch(/requested by default/);
      expect(text).not.toMatch(/nested under/i);
      expect(text).not.toMatch(/'body'|\bbody\b object|under a body/i);
      expect(text).not.toMatch(/\bCLI\b|--[a-z]+-[a-z]+/);
      expect(text).not.toMatch(/includeThumbnails|beta\b/i);
    }
  });
});

describe.each(SOURCES)("copilot-retrieve %s: unavailable permission", (dataSource) => {
  it("403 is a capability_gap with the no-escalation guidance after exactly one request", async () => {
    responses.push(() =>
      json({ error: { code: "accessDenied", message: `Access denied. Bearer ${TOKEN}` } }, 403),
    );
    const { env, text } = await retrieve(dataSource);
    expect(env.ok).toBe(false);
    expect(env.data).toBeUndefined();
    expect(env.error?.code).toBe("capability_gap");
    expect(env.error?.message).toContain("do not broaden scopes automatically");
    expect(env.error?.details).toMatchObject({ status: 403, graphCode: "accessDenied" });
    expect(calls).toHaveLength(1); // no retry, no fallback to another data source or scope
    expect(calls[0]?.method).toBe("POST");
    expect(text).not.toContain(TOKEN);
  });

  it("403 without a JSON body is still a capability_gap and leaks nothing", async () => {
    responses.push(() => new Response(`body-marker-xyz ${TOKEN}`, { status: 403 }));
    const { env, text } = await retrieve(dataSource);
    expect(env.error?.code).toBe("capability_gap");
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
    expect(text).not.toContain("body-marker-xyz");
  });

  it("401 is reauth_required after one request and never mentions scopes", async () => {
    responses.push(() =>
      json({ error: { code: "InvalidAuthenticationToken", message: `expired ${TOKEN}` } }, 401),
    );
    const { env, text } = await retrieve(dataSource);
    expect(env.error?.code).toBe("reauth_required");
    expect(env.error?.message).not.toMatch(/scope/i);
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
  });

  it("a Graph 400 or 404 is a graph_error, never an empty successful result", async () => {
    for (const status of [400, 404]) {
      calls = [];
      responses.push(() =>
        json({ error: { code: "badRequest", message: `bad ${TOKEN}` } }, status),
      );
      const { env, text } = await retrieve(dataSource);
      expect(env.ok, String(status)).toBe(false);
      expect(env.data).toBeUndefined();
      expect(env.error?.code).toBe("graph_error");
      expect(env.error?.details).toMatchObject({ status });
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("a 200 that is not a retrievalResponse is a failure, not an empty success", async () => {
    responses.push(() => json({}));
    const { env } = await retrieve(dataSource);
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("graph_error");
  });
});

describe("copilot-retrieve invalid input makes no Graph request", () => {
  it("an unknown dataSource is invalid_input with zero requests (not a capability_gap)", async () => {
    const { env } = await retrieve("mailbox");
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });
});
