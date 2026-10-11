import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";

/**
 * SM-13 permission/capability matrix for the three SharePoint site tools (part A), the two site
 * drive tools (part B) and the six list/list-item/column tools (part C). Mirrors
 * teams-capability.test.ts: the catalog's scope statement is pinned exactly, a 403 is a
 * `capability_gap` carrying the "do not broaden scopes automatically" guidance after exactly ONE
 * request (no retry, no escalation), a 401 is `reauth_required`, a 401/403 on a followed page fails
 * the whole call, no token ever leaks (single path and batch, even when Graph echoes it), and a Graph
 * failure never becomes an empty successful result. The server never selects or pre-checks scopes
 * (`workScopes` is catalog documentation; Graph decides), so the Sites.Read.All vs Sites.Selected
 * alternatives are asserted as catalog data plus the one-request/no-escalation behavior below.
 * Documented in PERMISSIONS.md ("SM-13 part A permission/capability matrix", "SM-13 part B").
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const HOST = "contoso.sharepoint.com";
const LIST_ID = "f17e0000-0000-4000-8000-000000000003";
const COL_ID = "f17e0000-0000-4000-8000-000000000013";
const LISTS = `/sites/${HOST}/lists`;

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

type Case = {
  name: string;
  args: Record<string, unknown>;
  paged: boolean;
  /** Absolute collection path a followed page must stay on (paged only). */
  path: string;
  /** Query of the continuation link a followed page carries (paged only). */
  next?: string;
  batchUrl: string;
  /** The catalog's `workScopes`, pinned exactly (never narrowed or broadened by this unit). */
  scopes: unknown;
};
const SM13_TOOLS: Case[] = [
  {
    name: "search-sharepoint-sites",
    args: { search: "hr" },
    paged: true,
    path: "/sites",
    next: "search=hr&$skiptoken=abc",
    batchUrl: "/sites?search=hr",
    // The site-search page: "This method doesn't support the Sites.Selected application permission."
    scopes: ["Sites.Read.All"],
  },
  {
    name: "get-sharepoint-site",
    args: { "site-id": HOST },
    paged: false,
    path: "",
    batchUrl: `/sites/${HOST}`,
    scopes: [["Sites.Read.All"], ["Sites.Selected"]],
  },
  {
    name: "get-sharepoint-site-by-path",
    args: { "site-id": HOST, "relative-path": "sites/hr" },
    paged: false,
    path: "",
    batchUrl: `/sites/${HOST}:/sites/hr`,
    scopes: [["Sites.Read.All"], ["Sites.Selected"]],
  },
  {
    name: "list-sharepoint-site-drives",
    args: { "site-id": HOST },
    paged: true,
    path: `/sites/${HOST}/drives`,
    next: "$top=25&$skiptoken=abc",
    batchUrl: `/sites/${HOST}/drives`,
    // D33: Files.Read.All or Sites.Read.All (both on the drive pages) or Sites.Selected (D30).
    scopes: [["Files.Read.All"], ["Sites.Read.All"], ["Sites.Selected"]],
  },
  {
    name: "get-sharepoint-site-drive-by-id",
    args: {
      "site-id": HOST,
      "drive-id": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
    },
    paged: false,
    path: "",
    batchUrl: `/sites/${HOST}/drives/b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd`,
    scopes: [["Files.Read.All"], ["Sites.Read.All"], ["Sites.Selected"]],
  },
  {
    name: "list-sharepoint-site-lists",
    args: { "site-id": HOST },
    paged: true,
    path: LISTS,
    next: "$skiptoken=abc",
    batchUrl: LISTS,
    // The list pages document Sites.Read.All only (no Sites.Selected); D30 keeps Selected.
    scopes: [["Sites.Read.All"], ["Sites.Selected"]],
  },
  {
    name: "get-sharepoint-site-list",
    args: { "site-id": HOST, "list-id": LIST_ID },
    paged: false,
    path: "",
    batchUrl: `${LISTS}/${LIST_ID}`,
    scopes: [["Sites.Read.All"], ["Sites.Selected"]],
  },
  {
    name: "list-sharepoint-site-list-items",
    args: { "site-id": HOST, "list-id": LIST_ID },
    paged: true,
    path: `${LISTS}/${LIST_ID}/items`,
    next: "$skiptoken=Paged%3dTRUE%26p_ID%3d5",
    batchUrl: `${LISTS}/${LIST_ID}/items`,
    scopes: [["Sites.Read.All"], ["Sites.Selected"]],
  },
  {
    name: "get-sharepoint-site-list-item",
    args: { "site-id": HOST, "list-id": LIST_ID, "listItem-id": "5" },
    paged: false,
    path: "",
    batchUrl: `${LISTS}/${LIST_ID}/items/5`,
    scopes: [["Sites.Read.All"], ["Sites.Selected"]],
  },
  {
    name: "list-sharepoint-list-columns",
    args: { "site-id": HOST, "list-id": LIST_ID },
    paged: true,
    path: `${LISTS}/${LIST_ID}/columns`,
    next: "$skiptoken=abc",
    batchUrl: `${LISTS}/${LIST_ID}/columns`,
    scopes: [["Sites.Read.All"], ["Sites.Selected"]],
  },
  {
    name: "get-sharepoint-list-column",
    args: { "site-id": HOST, "list-id": LIST_ID, "columnDefinition-id": COL_ID },
    paged: false,
    path: "",
    batchUrl: `${LISTS}/${LIST_ID}/columns/${COL_ID}`,
    scopes: [["Sites.Read.All"], ["Sites.Selected"]],
  },
];

const PART_A = ["search-sharepoint-sites", "get-sharepoint-site", "get-sharepoint-site-by-path"];
const PART_B = ["list-sharepoint-site-drives", "get-sharepoint-site-drive-by-id"];
const PART_C = [
  "list-sharepoint-site-lists",
  "get-sharepoint-site-list",
  "list-sharepoint-site-list-items",
  "get-sharepoint-site-list-item",
  "list-sharepoint-list-columns",
  "get-sharepoint-list-column",
];

describe("SM-13 tool set and catalog scopes", () => {
  it("part A covers exactly the three site tools (search, get, by-path)", () => {
    const names = catalog
      .filter(
        (e) =>
          e.method === "get" &&
          e.toolName !== undefined &&
          /^\/sites(\/\{site-id\}(:\/\{path\})?)?$/.test(e.pathPattern),
      )
      .map((e) => e.toolName)
      .sort();
    expect(names).toEqual([...PART_A].sort());
  });

  it("part C covers exactly the six list, list-item and column tools", () => {
    const names = catalog
      .filter(
        (e) =>
          e.method === "get" &&
          e.toolName !== undefined &&
          /^\/sites\/\{site-id\}\/lists(\/|$)/.test(e.pathPattern),
      )
      .map((e) => e.toolName)
      .sort();
    expect(names).toEqual([...PART_C].sort());
    expect(SM13_TOOLS.map((c) => c.name).sort()).toEqual([...PART_A, ...PART_B, ...PART_C].sort());
  });

  it("part B covers exactly the two drive tools; the site items routes are not catalogued (D32)", () => {
    const drives = catalog
      .filter(
        (e) =>
          e.method === "get" &&
          /^\/sites\/\{site-id\}\/drives(\/\{drive-id\})?$/.test(e.pathPattern),
      )
      .map((e) => e.toolName)
      .sort();
    expect(drives).toEqual(["get-sharepoint-site-drive-by-id", "list-sharepoint-site-drives"]);
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    // D32: the site `items` collection "can't be enumerated" and has no API page; removed.
    for (const n of ["list-sharepoint-site-items", "get-sharepoint-site-item"]) {
      expect(
        catalog.some((e) => e.toolName === n),
        n,
      ).toBe(false);
      expect(handlers.has(n), n).toBe(false);
    }
    expect(catalog.some((e) => /^\/sites\/\{site-id\}\/items/.test(e.pathPattern))).toBe(false);
  });

  it("implements each as a v1.0 GET handler whose catalog scopes are exactly as recorded", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const c of SM13_TOOLS) {
      expect(handlers.has(c.name), c.name).toBe(true);
      const e = catalog.find((x) => x.toolName === c.name);
      expect(e?.method, c.name).toBe("get");
      expect(e?.["apiVersion"], c.name).not.toBe("beta");
      expect(e?.["workScopes"], c.name).toEqual(c.scopes);
    }
  });

  it("search lists only Sites.Read.All; get and by-path list Sites.Read.All OR Sites.Selected", () => {
    const scopes = (n: string) => catalog.find((x) => x.toolName === n)?.["workScopes"];
    expect(scopes("search-sharepoint-sites")).toEqual(["Sites.Read.All"]);
    for (const n of ["get-sharepoint-site", "get-sharepoint-site-by-path"]) {
      expect(scopes(n)).toEqual([["Sites.Read.All"], ["Sites.Selected"]]);
    }
  });

  it("implements exactly the catalogued SharePoint tools (SM-13 complete)", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    const sharepoint = catalog.filter((e) => /sharepoint/.test(e.toolName)).map((e) => e.toolName);
    expect([...sharepoint].sort()).toEqual(SM13_TOOLS.map((c) => c.name).sort());
    for (const n of sharepoint) expect(handlers.has(n), n).toBe(true);
    // D36: the application-only sites delta is not catalogued (delegated access not supported).
    expect(catalog.some((e) => e.toolName === "get-sharepoint-sites-delta")).toBe(false);
    expect(catalog.some((e) => /^\/sites\/delta/.test(e.pathPattern))).toBe(false);
  });
});

describe.each(SM13_TOOLS)("$name: unavailable permission", (c) => {
  it("403 is a capability_gap with the no-escalation guidance after exactly one request", async () => {
    responses.push(() =>
      json(
        {
          error: {
            code: "accessDenied",
            message: `Access denied. Bearer ${TOKEN}`,
          },
        },
        403,
      ),
    );
    const { env, text } = await call(c.name, c.args);
    expect(env.ok).toBe(false);
    expect(env.data).toBeUndefined();
    expect(env.error?.code).toBe("capability_gap");
    expect(env.error?.message).toContain("do not broaden scopes automatically");
    expect(env.error?.details).toMatchObject({ status: 403, graphCode: "accessDenied" });
    expect(calls).toHaveLength(1); // no retry, no re-request with other scopes/headers
    expect(calls[0]?.method).toBe("GET");
    expect(text).not.toContain(TOKEN);
  });

  it("403 without a JSON body is still a capability_gap and leaks nothing", async () => {
    responses.push(() => new Response(`body-marker-xyz ${TOKEN}`, { status: 403 }));
    const { env, text } = await call(c.name, c.args);
    expect(env.error?.code).toBe("capability_gap");
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
    expect(text).not.toContain("body-marker-xyz");
  });

  it("401 is reauth_required after one request and never mentions scopes", async () => {
    responses.push(() =>
      json({ error: { code: "InvalidAuthenticationToken", message: `expired ${TOKEN}` } }, 401),
    );
    const { env, text } = await call(c.name, c.args);
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("reauth_required");
    expect(env.error?.message).not.toMatch(/scope/i);
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
  });

  it("a Graph 400 or 404 is a graph_error, never an empty successful result", async () => {
    for (const status of [400, 404]) {
      calls = [];
      responses.push(() =>
        json({ error: { code: "itemNotFound", message: `no such thing ${TOKEN}` } }, status),
      );
      const { env, text } = await call(c.name, c.args);
      expect(env.ok, String(status)).toBe(false);
      expect(env.data).toBeUndefined();
      expect(env.error?.code).toBe("graph_error");
      expect(env.error?.details).toMatchObject({ status });
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });

  it.runIf(c.paged)("a genuine empty 200 stays a successful empty list", async () => {
    responses.push(() => json({ value: [] }));
    const { env } = await call(c.name, c.args);
    expect(env.ok).toBe(true);
    expect(env.data).toEqual([]);
    expect(env.meta).toMatchObject({ count: 0, hasMore: false });
  });

  it.runIf(c.paged)("a 200 without a value array is a failure, not an empty list", async () => {
    responses.push(() => json({}));
    const { env } = await call(c.name, c.args);
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("graph_error");
  });

  it.runIf(!c.paged)("a 200 that is not a JSON object is a failure, not a success", async () => {
    responses.push(() => json([]));
    const { env } = await call(c.name, c.args);
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("graph_error");
  });
});

describe.each(SM13_TOOLS.filter((c) => c.paged))("$name: permission lost mid-listing", (c) => {
  it("a 403 or 401 on a followed page fails the call instead of returning the first page", async () => {
    for (const [status, code] of [
      [403, "capability_gap"],
      [401, "reauth_required"],
    ] as const) {
      calls = [];
      responses.push(() =>
        json({
          value: [{ id: "1" }],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0${c.path}?${c.next ?? ""}`,
        }),
      );
      responses.push(() => json({ error: { code: "x", message: TOKEN } }, status));
      const { env, text } = await call(c.name, { ...c.args, top: 50 });
      expect(env.ok, `${c.name} ${status}`).toBe(false);
      expect(env.data).toBeUndefined();
      expect(env.error?.code).toBe(code);
      expect(calls).toHaveLength(2); // first page + the denied follow-up, no further retries
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe("graph-batch: a denied SM-13 subrequest", () => {
  const batch = (url: string) =>
    call("graph-batch", { requests: [{ id: "1", method: "GET", url }] });

  it.each(SM13_TOOLS)(
    "$name: a 403 subresponse is a capability_gap without a body or token",
    async (c) => {
      responses.push(() =>
        json({
          responses: [
            {
              id: "1",
              status: 403,
              body: {
                error: { code: "accessDenied", message: "Access denied", extra: "raw-body-marker" },
              },
            },
          ],
        }),
      );
      const { env, text } = await batch(c.batchUrl);
      const first = (env.data as { responses: Array<{ error?: { code: string }; body?: unknown }> })
        .responses[0];
      expect(first?.error?.code).toBe("capability_gap");
      expect(first?.body).toBeUndefined();
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
      expect(text).not.toContain("raw-body-marker");
    },
  );

  it.each(SM13_TOOLS)("$name: a 401 subresponse is reauth_required", async (c) => {
    responses.push(() =>
      json({
        responses: [
          { id: "1", status: 401, body: { error: { code: "InvalidAuthenticationToken" } } },
        ],
      }),
    );
    const { env, text } = await batch(c.batchUrl);
    const first = (env.data as { responses: Array<{ error?: { code: string }; body?: unknown }> })
      .responses[0];
    expect(first?.error?.code).toBe("reauth_required");
    expect(first?.body).toBeUndefined();
    expect(text).not.toContain(TOKEN);
  });

  it.each(SM13_TOOLS)(
    "$name: a batch response that echoes the token fails closed without leaking it",
    async (c) => {
      responses.push(() =>
        json({
          responses: [
            { id: "1", status: 403, body: { error: { code: "x", message: `echo ${TOKEN}` } } },
          ],
        }),
      );
      const { env, text } = await batch(c.batchUrl);
      expect(env.ok).toBe(false);
      expect(env.error?.code).toBe("graph_error");
      expect(text).not.toContain(TOKEN);
    },
  );
});
