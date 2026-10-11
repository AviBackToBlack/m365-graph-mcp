import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";

/**
 * SM-14 part B permission/capability matrix for `search-query` (D38). The catalog's per-entity-type
 * scope alternatives are pinned exactly (People.Read dropped); a 403 is a `capability_gap` with the
 * "do not broaden scopes automatically" guidance after exactly ONE request for every entity type
 * and a mixed SharePoint query (no retry, no per-type split, no escalation); a 401 is
 * `reauth_required`; a 400/404 is a `graph_error`; no token ever leaks. The server never selects or
 * pre-checks scopes. Documented in PERMISSIONS.md ("SM-14 part B").
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
const search = async (entityTypes: string[]) => {
  const r = await mcp.callTool({
    name: "search-query",
    arguments: { entityTypes, queryString: "x" },
  });
  return { env: r.structuredContent as Env, text: JSON.stringify(r) };
};

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  const fetch: FetchLike = (url, init) => {
    calls.push({ url, method: String(init.method) });
    const next = responses.shift();
    return Promise.resolve(next ? next() : json({ value: [{ hitsContainers: [] }] }));
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

const CASES: Array<[string, string[]]> = [
  ["message", ["message"]],
  ["event", ["event"]],
  ["chatMessage", ["chatMessage"]],
  ["driveItem", ["driveItem"]],
  ["mixed SharePoint/OneDrive", ["drive", "driveItem", "list", "listItem", "site"]],
];

describe("search-query catalog scopes (D38)", () => {
  it("is a v1.0 POST entry whose workScopes are exactly the per-entity-type alternatives", () => {
    const e = catalog.find((x) => x.toolName === "search-query");
    expect(e?.method).toBe("post");
    expect(e?.["apiVersion"]).not.toBe("beta");
    expect(e?.pathPattern).toBe("/search/query");
    expect(e?.["workScopes"]).toEqual([
      ["Mail.Read"],
      ["Calendars.Read"],
      ["Files.Read.All"],
      ["Sites.Read.All"],
      ["Chat.Read"],
      ["ChannelMessage.Read.All"],
    ]);
    expect(JSON.stringify(e)).not.toContain("People.Read");
  });

  it("pins the typed-contract catalog text", () => {
    const e = catalog.find((x) => x.toolName === "search-query");
    const desc = String(e?.["descriptionOverride"]);
    const tip = String(e?.["llmTip"]);
    expect(desc).toMatch(/lexical keyword\/KQL retrieval/);
    expect(desc).toMatch(/not Copilot semantic retrieval/);
    expect(desc).toMatch(/one typed search request per call/);
    expect(desc).toMatch(
      /message \(Outlook mail\), event \(calendar events\) or chatMessage \(Teams messages\), each of which must be searched alone/,
    );
    expect(desc).toMatch(/any non-empty combination of drive, driveItem, list, listItem and site/);
    expect(desc).toMatch(/at most 1000 characters/);
    expect(desc).toMatch(
      /SharePoint\/OneDrive search may return any of drive, driveItem, list, listItem or site, not only the selected ones/,
    );
    expect(desc).not.toMatch(/dropped when its type was not requested/);
    expect(desc).toMatch(/at most 25 for message and event and at most 50 for every other type/);
    expect(desc).toMatch(
      /message needs Mail\.Read; event Calendars\.Read; chatMessage Chat\.Read or ChannelMessage\.Read\.All/,
    );
    expect(desc).toMatch(
      /drive and driveItem Files\.Read\.All or Sites\.Read\.All; list, listItem and site Sites\.Read\.All/,
    );
    expect(desc).toMatch(/alternatives per entity type/);
    expect(desc).toMatch(/needs a token that already covers every selected type/);
    expect(desc).toMatch(/never selects or broadens consent/);
    expect(desc).toMatch(/capability_gap/);
    expect(desc).toMatch(/Work\/school accounts only/);
    expect(desc).toMatch(/Not supported: fields, sorting, aggregations, collapse, queryTemplate/);
    expect(desc).toMatch(/untrusted data/);
    expect(tip).toMatch(/Lexical Microsoft Search \(KQL\), not semantic retrieval/);
    expect(tip).toMatch(/for message and event the first page must use from 0/);
    expect(tip).toMatch(/hitId is a search hit identifier, not a request target/);
    for (const text of [desc, tip]) {
      expect(text).not.toMatch(/People\.Read|\$filter|\$search|Prefer/);
    }
  });
});

describe.each(CASES)("search-query %s: unavailable permission", (_label, types) => {
  it("403 is a capability_gap with the no-escalation guidance after exactly one request", async () => {
    responses.push(() =>
      json({ error: { code: "accessDenied", message: `Access denied. Bearer ${TOKEN}` } }, 403),
    );
    const { env, text } = await search(types);
    expect(env.ok).toBe(false);
    expect(env.data).toBeUndefined();
    expect(env.error?.code).toBe("capability_gap");
    expect(env.error?.message).toContain("do not broaden scopes automatically");
    expect(env.error?.details).toMatchObject({ status: 403, graphCode: "accessDenied" });
    expect(calls).toHaveLength(1); // no retry, no split per entity type, no other scope
    expect(calls[0]?.method).toBe("POST");
    expect(text).not.toContain(TOKEN);
  });

  it("403 without a JSON body is still a capability_gap and leaks nothing", async () => {
    responses.push(() => new Response(`body-marker-xyz ${TOKEN}`, { status: 403 }));
    const { env, text } = await search(types);
    expect(env.error?.code).toBe("capability_gap");
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
    expect(text).not.toContain("body-marker-xyz");
  });

  it("401 is reauth_required after one request and never mentions scopes", async () => {
    responses.push(() =>
      json({ error: { code: "InvalidAuthenticationToken", message: `expired ${TOKEN}` } }, 401),
    );
    const { env, text } = await search(types);
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
      const { env, text } = await search(types);
      expect(env.ok, String(status)).toBe(false);
      expect(env.data).toBeUndefined();
      expect(env.error?.code).toBe("graph_error");
      expect(env.error?.details).toMatchObject({ status });
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("a 200 that is not a searchResponse is a failure, not an empty success", async () => {
    responses.push(() => json({}));
    const { env } = await search(types);
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("graph_error");
  });
});

describe("search-query invalid input makes no Graph request", () => {
  it("a cross-family request is invalid_input with zero requests (not a capability_gap)", async () => {
    const { env } = await search(["message", "driveItem"]);
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });
});
