import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";
import {
  MAX_EXTRACT_TEXT,
  MAX_EXTRACTS_PER_HIT,
  MAX_FILTER_EXPRESSION,
  MAX_LABEL_TEXT,
  MAX_METADATA_VALUE,
  MAX_QUERY_STRING,
  MAX_RESOURCE_METADATA_KEYS,
  MAX_RESULTS,
  boundText,
} from "./copilot-retrieve.js";

/**
 * SM-14 part C behavior tests for `copilot-retrieve` (D41): exact constructed request, local input
 * validation, strict unknown-property rejection, response allowlist projection, fail-closed bodies
 * and batch unreachability. The permission/capability matrix and catalog text live in
 * copilot-retrieve-capability.test.ts.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";

interface Call {
  url: string;
  method: string;
  headers: Record<string, string>;
  body: unknown;
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
  data?: { retrievalHits: Array<Record<string, unknown>> };
  meta?: Record<string, unknown>;
  error?: { code: string; message: string; details?: Record<string, unknown> };
};
const call = async (name: string, args: Record<string, unknown> = {}) => {
  const r = await mcp.callTool({ name, arguments: args });
  return { r, env: r.structuredContent as Env, text: JSON.stringify(r) };
};
const retrieve = (args: Record<string, unknown>) =>
  call("copilot-retrieve", { queryString: "How to setup corporate VPN?", ...args });
const sp = (args: Record<string, unknown> = {}) => retrieve({ dataSource: "sharePoint", ...args });

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  const fetch: FetchLike = (url, init) => {
    calls.push({
      url,
      method: String(init.method),
      headers: { ...(init.headers as Record<string, string>) },
      body: typeof init.body === "string" ? JSON.parse(init.body) : init.body,
    });
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

const reply = (hits: unknown[]) => json({ retrievalHits: hits });
const hit = (extra: Record<string, unknown> = {}) => ({
  webUrl: "https://contoso.sharepoint.com/sites/HR/VPNAccess.docx",
  extracts: [{ text: "To configure the VPN, click the Wi-Fi icon.", relevanceScore: 0.83 }],
  resourceType: "listItem",
  resourceMetadata: {},
  ...extra,
});

describe("copilot-retrieve request construction", () => {
  it("sends exactly POST /v1.0/copilot/retrieval with only the two required body fields", async () => {
    responses.push(() => reply([]));
    const { env } = await retrieve({ dataSource: "sharePoint" });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    const c = calls[0] as Call;
    expect(c.method).toBe("POST");
    const url = new URL(c.url);
    expect(url.origin + url.pathname).toBe("https://graph.microsoft.com/v1.0/copilot/retrieval");
    expect(url.search).toBe("");
    expect(c.body).toStrictEqual({
      queryString: "How to setup corporate VPN?",
      dataSource: "sharePoint",
    });
    expect(c.headers["Prefer"]).toBeUndefined();
    expect(c.headers["ConsistencyLevel"]).toBeUndefined();
  });

  it("sends every optional field top-level, exactly as provided, and nothing else", async () => {
    responses.push(() => reply([]));
    await retrieve({
      dataSource: "externalItem",
      filterExpression: 'Author:"Megan Bowen" AND FileType:"docx"',
      resourceMetadata: ["title", "Label_Title"],
      maximumNumberOfResults: 4,
    });
    expect(calls[0]?.body).toStrictEqual({
      queryString: "How to setup corporate VPN?",
      dataSource: "externalItem",
      filterExpression: 'Author:"Megan Bowen" AND FileType:"docx"',
      resourceMetadata: ["title", "Label_Title"],
      maximumNumberOfResults: 4,
    });
  });

  it.each(["sharePoint", "oneDriveBusiness", "externalItem"])(
    "accepts dataSource %s",
    async (dataSource) => {
      responses.push(() => reply([]));
      expect((await retrieve({ dataSource })).env.ok).toBe(true);
      expect((calls[0]?.body as Record<string, unknown>)["dataSource"]).toBe(dataSource);
    },
  );
});

describe("copilot-retrieve local validation (zero Graph requests)", () => {
  it("pins the limits: Microsoft 1500/25, project 1000/20 and response bounds", () => {
    expect(MAX_QUERY_STRING).toBe(1500);
    expect(MAX_RESULTS).toBe(25);
    expect(MAX_FILTER_EXPRESSION).toBe(1000);
    expect(MAX_RESOURCE_METADATA_KEYS).toBe(20);
    expect(MAX_EXTRACT_TEXT).toBe(4000);
    expect(MAX_EXTRACTS_PER_HIT).toBe(20);
    expect(MAX_METADATA_VALUE).toBe(1000);
    expect(MAX_LABEL_TEXT).toBe(256);
  });

  it("queryString: blank, missing, over 1500, non-string are rejected; 1500 is accepted", async () => {
    for (const queryString of ["", "   ", "\n\t", "x".repeat(1501), 5, null]) {
      const { env } = await call("copilot-retrieve", { queryString, dataSource: "sharePoint" });
      expect(env.error?.code, String(queryString).slice(0, 8)).toBe("invalid_input");
    }
    expect((await call("copilot-retrieve", { dataSource: "sharePoint" })).env.error?.code).toBe(
      "invalid_input",
    );
    expect(calls).toHaveLength(0);
    responses.push(() => reply([]));
    expect((await sp({ queryString: "x".repeat(1500) })).env.ok).toBe(true);
  });

  it("dataSource: missing, unknown, array, wrong case are rejected (one source per call)", async () => {
    for (const dataSource of [
      "mail",
      "SharePoint",
      "",
      ["sharePoint"],
      ["sharePoint", "externalItem"],
      1,
    ]) {
      const { env } = await retrieve({ dataSource });
      expect(env.error?.code, JSON.stringify(dataSource)).toBe("invalid_input");
    }
    expect((await retrieve({})).env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("filterExpression: blank, over the project limit (1000), non-string are rejected", async () => {
    for (const filterExpression of ["", "  ", "x".repeat(1001), 5, {}]) {
      const { env } = await sp({ filterExpression });
      expect(env.error?.code, String(filterExpression).slice(0, 8)).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
    responses.push(() => reply([]));
    expect((await sp({ filterExpression: "x".repeat(1000) })).env.ok).toBe(true);
  });

  it("resourceMetadata: empty, too many, duplicates, non-identifier and non-string items are rejected", async () => {
    const tooMany = Array.from({ length: 21 }, (_, i) => `f${i}`);
    for (const resourceMetadata of [
      [],
      tooMany,
      ["title", "title"],
      ["title", "Title"], // case-insensitive: both would match the same returned key
      ["a b"],
      ["title;DROP"],
      ["../x"],
      ["1abc"],
      ["_x"],
      ["__proto__"],
      [""],
      ["x".repeat(65)],
      [5],
      "title",
    ]) {
      const { env } = await sp({ resourceMetadata });
      expect(env.error?.code, JSON.stringify(resourceMetadata)?.slice(0, 40)).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
    responses.push(() => reply([]));
    const twenty = Array.from({ length: 20 }, (_, i) => `f${i}`);
    expect((await sp({ resourceMetadata: twenty })).env.ok).toBe(true);
  });

  it("maximumNumberOfResults: 0, 26, negative, fractional and non-number are rejected", async () => {
    for (const maximumNumberOfResults of [0, -1, 26, 100, 1.5, "5", null]) {
      const { env } = await sp({ maximumNumberOfResults });
      expect(env.error?.code, String(maximumNumberOfResults)).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
    for (const n of [1, 25]) {
      responses.push(() => reply([]));
      expect((await sp({ maximumNumberOfResults: n })).env.ok).toBe(true);
    }
  });

  it.each([
    "body",
    "dataSourceConfiguration",
    "includeThumbnails",
    "headers",
    "url",
    "version",
    "apiVersion",
    "method",
    "path",
    "top",
    "select",
    "filter",
    "$filter",
    "query",
    "requests",
    "pageNumbers",
    "connections",
    "connectionId",
    "Prefer",
  ])("unknown input property %s is rejected before any request", async (key) => {
    const { env, text } = await sp({ [key]: "x" });
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
    expect(text).not.toContain(TOKEN);
  });

  it("a nested body object (the old Softeria shape) is rejected, not unwrapped", async () => {
    const { env } = await call("copilot-retrieve", {
      body: { queryString: "x", dataSource: "sharePoint" },
    });
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });
});

describe("boundText", () => {
  it("truncates with an ellipsis that stays within the limit", () => {
    expect(boundText("x".repeat(5000), 4000)).toBe(`${"x".repeat(3999)}…`);
    expect(boundText("x".repeat(5000), 4000).length).toBe(4000);
    expect(boundText("y".repeat(4000), 4000)).toBe("y".repeat(4000));
  });
});

describe("copilot-retrieve response projection", () => {
  it("pinned leniencies: null relevanceScore is absent; a hit without extracts, webUrl or resourceType is kept with extracts []", async () => {
    responses.push(() =>
      json({
        retrievalHits: [
          {
            webUrl: "https://contoso.sharepoint.com/a.docx",
            extracts: [{ text: "t", relevanceScore: null }],
          },
          { extracts: null },
          {},
        ],
      }),
    );
    const { env } = await sp({});
    expect(env.ok).toBe(true);
    expect(env.data).toStrictEqual({
      retrievalHits: [
        {
          webUrl: "https://contoso.sharepoint.com/a.docx",
          extracts: [{ text: "t" }],
          resourceMetadata: {},
        },
        { extracts: [], resourceMetadata: {} },
        { extracts: [], resourceMetadata: {} },
      ],
    });
  });

  it("keeps only allowlisted fields; nothing else leaks", async () => {
    responses.push(() =>
      json({
        "@odata.context": "CONTEXT_SECRET",
        retrievalHits: [
          {
            "@odata.type": "#microsoft.graph.retrievalHit",
            webUrl: "https://contoso.sharepoint.com/sites/HR/VPNAccess.docx",
            extracts: [
              {
                "@odata.type": "microsoft.graph.retrievalExtract",
                text: "To configure the VPN",
                relevanceScore: 0.8374363553387588,
                pageNumbers: [1],
                extractExtra: "EXTRACT_EXTRA_SECRET",
              },
              { text: "No score here" },
            ],
            resourceType: "listItem",
            resourceMetadata: { title: "VPN Access", author: "John Doe", hidden: "META_SECRET" },
            sensitivityLabel: {
              sensitivityLabelId: "f71f1f74",
              displayName: "Confidential",
              tooltip: "TOOLTIP_SECRET",
              toolTip: "TOOLTIP2_SECRET",
              priority: 4,
              color: "#FF8C00",
            },
            thumbnails: [{ pageNumber: 1, content: "THUMB_SECRET", mimeType: "image/jpg" }],
            hitExtra: "HIT_EXTRA_SECRET",
          },
        ],
      }),
    );
    const { env, text } = await sp({ resourceMetadata: ["title", "author"] });
    expect(env.ok).toBe(true);
    expect(env.data).toStrictEqual({
      retrievalHits: [
        {
          webUrl: "https://contoso.sharepoint.com/sites/HR/VPNAccess.docx",
          resourceType: "listItem",
          extracts: [
            { text: "To configure the VPN", relevanceScore: 0.8374363553387588 },
            { text: "No score here" },
          ],
          resourceMetadata: { title: "VPN Access", author: "John Doe" },
          sensitivityLabel: {
            sensitivityLabelId: "f71f1f74",
            displayName: "Confidential",
            priority: 4,
          },
        },
      ],
    });
    expect(text).not.toMatch(/SECRET/);
    expect(text).not.toContain("@odata");
    expect(text).not.toContain(TOKEN);
  });

  it("resourceMetadata: nothing is returned when none was requested, and only requested keys otherwise", async () => {
    responses.push(() =>
      reply([hit({ resourceMetadata: { title: "T_SECRET", author: "A_SECRET" } })]),
    );
    const none = await sp();
    expect(none.env.data?.retrievalHits[0]?.["resourceMetadata"]).toStrictEqual({});
    expect(none.text).not.toMatch(/SECRET/);
    responses.push(() => reply([hit({ resourceMetadata: { Title: "kept", author: "A_SECRET" } })]));
    const some = await sp({ resourceMetadata: ["title"] });
    expect(some.env.data?.retrievalHits[0]?.["resourceMetadata"]).toStrictEqual({ title: "kept" });
    expect(some.text).not.toMatch(/SECRET/);
  });

  it("resourceMetadata: only scalar values survive and long strings are bounded", async () => {
    responses.push(() =>
      reply([
        hit({
          resourceMetadata: {
            a: "s",
            b: 3,
            c: false,
            d: { nested: "NESTED_SECRET" },
            e: ["ARRAY_SECRET"],
            f: null,
            g: "z".repeat(5000),
          },
        }),
      ]),
    );
    const { env, text } = await sp({
      resourceMetadata: ["a", "b", "c", "d", "e", "f", "g", "missing"],
    });
    expect(env.data?.retrievalHits[0]?.["resourceMetadata"]).toStrictEqual({
      a: "s",
      b: 3,
      c: false,
      g: `${"z".repeat(MAX_METADATA_VALUE - 1)}…`,
    });
    expect(text).not.toMatch(/SECRET/);
  });

  it("a prototype-named key in the response cannot pollute or leak", async () => {
    responses.push(() =>
      reply([
        hit({
          resourceMetadata: JSON.parse('{"__proto__": {"x": "PROTO_SECRET"}, "constructor": "c"}'),
        }),
      ]),
    );
    const { env, text } = await sp({ resourceMetadata: ["constructor"] });
    expect(env.data?.retrievalHits[0]?.["resourceMetadata"]).toStrictEqual({ constructor: "c" });
    expect(text).not.toMatch(/SECRET/);
  });

  it("bounds extract text with a marker inside the limit", async () => {
    responses.push(() =>
      reply([hit({ extracts: [{ text: "t".repeat(9000), relevanceScore: 0.5 }] })]),
    );
    const { env } = await sp();
    const text = (env.data?.retrievalHits[0]?.["extracts"] as Array<{ text: string }>)[0]?.text;
    expect(text).toBe(`${"t".repeat(MAX_EXTRACT_TEXT - 1)}…`);
    expect(text?.length).toBe(MAX_EXTRACT_TEXT);
  });

  it("label strings are bounded; an empty label is dropped", async () => {
    responses.push(() =>
      reply([
        hit({ sensitivityLabel: { displayName: "d".repeat(1000), priority: 1 } }),
        hit({ sensitivityLabel: { color: "#fff", tooltip: "only dropped fields" } }),
        hit({ sensitivityLabel: null }),
      ]),
    );
    const { env } = await sp();
    const hits = env.data?.retrievalHits ?? [];
    expect(hits[0]?.["sensitivityLabel"]).toStrictEqual({
      displayName: `${"d".repeat(MAX_LABEL_TEXT - 1)}…`,
      priority: 1,
    });
    expect(hits[1]).not.toHaveProperty("sensitivityLabel");
    expect(hits[2]).not.toHaveProperty("sensitivityLabel");
  });

  it("drops unknown resourceType members and non-http(s) webUrls but keeps the hit", async () => {
    responses.push(() =>
      reply([
        hit({ resourceType: "unknownFutureValue", webUrl: "javascript:alert(1)" }),
        hit({ resourceType: "externalItem", webUrl: "https://contoso.service-now.com/kb?id=1" }),
      ]),
    );
    const { env, text } = await sp();
    const hits = env.data?.retrievalHits ?? [];
    expect(hits[0]).not.toHaveProperty("resourceType");
    expect(hits[0]).not.toHaveProperty("webUrl");
    expect(hits[0]?.["extracts"]).toHaveLength(1);
    expect(hits[1]?.["resourceType"]).toBe("externalItem");
    expect(hits[1]?.["webUrl"]).toBe("https://contoso.service-now.com/kb?id=1");
    expect(text).not.toContain("javascript");
  });

  it("an empty retrievalHits array is a successful empty result", async () => {
    responses.push(() => reply([]));
    const { env } = await sp();
    expect(env.ok).toBe(true);
    expect(env.data).toStrictEqual({ retrievalHits: [] });
  });

  it("returns exactly the requested maximum when the service honors it", async () => {
    responses.push(() => reply([hit(), hit()]));
    const { env } = await sp({ maximumNumberOfResults: 2 });
    expect(env.data?.retrievalHits).toHaveLength(2);
  });
});

describe("copilot-retrieve fails closed on malformed bodies", () => {
  const bad: Array<[string, unknown, Record<string, unknown>?]> = [
    ["array body", []],
    ["string body", "x"],
    ["empty object (retrievalHits absent)", {}],
    ["retrievalHits null", { retrievalHits: null }],
    ["retrievalHits an object", { retrievalHits: {} }],
    ["retrievalHits a string", { retrievalHits: "x" }],
    [
      "more hits than maximumNumberOfResults",
      { retrievalHits: [hit(), hit()] },
      { maximumNumberOfResults: 1 },
    ],
    ["more than 25 hits", { retrievalHits: Array.from({ length: 26 }, () => hit()) }],
    ["hit not an object", { retrievalHits: ["x"] }],
    ["hit null", { retrievalHits: [null] }],
    ["webUrl not a string", { retrievalHits: [hit({ webUrl: 5 })] }],
    ["webUrl too long", { retrievalHits: [hit({ webUrl: `https://x/${"a".repeat(3000)}` })] }],
    ["extracts not an array", { retrievalHits: [hit({ extracts: "x" })] }],
    ["extracts an object", { retrievalHits: [hit({ extracts: {} })] }],
    ["too many extracts", { retrievalHits: [hit({ extracts: Array(21).fill({ text: "t" }) })] }],
    ["extract not an object", { retrievalHits: [hit({ extracts: ["x"] })] }],
    ["extract without text", { retrievalHits: [hit({ extracts: [{ relevanceScore: 1 }] })] }],
    ["extract text not a string", { retrievalHits: [hit({ extracts: [{ text: {} }] })] }],
    [
      "relevanceScore a string",
      { retrievalHits: [hit({ extracts: [{ text: "t", relevanceScore: "1" }] })] },
    ],
    ["resourceMetadata an array", { retrievalHits: [hit({ resourceMetadata: [] })] }],
    ["resourceMetadata a string", { retrievalHits: [hit({ resourceMetadata: "x" })] }],
    ["sensitivityLabel a string", { retrievalHits: [hit({ sensitivityLabel: "x" })] }],
  ];
  it.each(bad)("%s", async (_name, body, args = {}) => {
    responses.push(() => json(body));
    const { env, text } = await sp(args);
    expect(env.ok).toBe(false);
    expect(env.data).toBeUndefined();
    expect(env.error?.code).toBe("graph_error");
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
  });
});

describe("copilot-retrieve is not reachable through graph-batch (batch is GET-only)", () => {
  it("rejects a POST and a GET subrequest to /copilot/retrieval with zero Graph requests", async () => {
    for (const method of ["POST", "GET"]) {
      const { env, text } = await call("graph-batch", {
        requests: [{ id: "1", method, url: "/copilot/retrieval" }],
      });
      expect(env.ok, method).toBe(false);
      expect(calls, method).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("the handler declares no batchPolicy and the tool is a v1.0 POST catalog entry", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    expect(handlers.get("copilot-retrieve")?.batchPolicy).toBeUndefined();
    const entry = catalog.find((e) => e.toolName === "copilot-retrieve");
    expect(entry?.method).toBe("post");
    expect(entry?.["apiVersion"]).not.toBe("beta");
  });
});
