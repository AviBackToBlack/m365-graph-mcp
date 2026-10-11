import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";

/**
 * SM-14 part A permission/capability matrix for get-my-presence, get-user-presence and
 * list-trending-insights. Mirrors sharepoint-capability.test.ts: the catalog's scope statement is
 * pinned exactly, a 403 is a `capability_gap` with the "do not broaden scopes automatically"
 * guidance after exactly ONE request (no retry, no escalation), a 401 is `reauth_required`, a
 * 401/403 on a followed page fails the whole call, no token ever leaks (single path and batch, even
 * when Graph echoes it), and a Graph failure never becomes an empty successful result. The server
 * never selects or pre-checks scopes (`workScopes` is catalog documentation; Graph decides).
 * Documented in PERMISSIONS.md ("SM-14 part A").
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const GUID = "f17e0000-0000-4000-8000-000000000006";

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
  next?: string;
  batchUrl: string;
  /** The catalog's `workScopes`, pinned exactly (never narrowed or broadened by this unit). */
  scopes: unknown;
};
const TOOLS: Case[] = [
  {
    name: "get-my-presence",
    args: {},
    paged: false,
    path: "",
    batchUrl: "/me/presence",
    // presence-get: Presence.Read is least privileged for /me/presence.
    scopes: ["Presence.Read"],
  },
  {
    name: "get-user-presence",
    args: { "user-id": GUID },
    paged: false,
    path: "",
    batchUrl: `/users/${GUID}/presence`,
    // presence-get example 2: another user's presence requires Presence.Read.All.
    scopes: ["Presence.Read.All"],
  },
  {
    name: "list-trending-insights",
    args: {},
    paged: true,
    path: "/me/insights/trending",
    next: "$skiptoken=abc",
    batchUrl: "/me/insights/trending",
    // insights-list-trending: Sites.Read.All (delegated work/school only).
    scopes: ["Sites.Read.All"],
  },
];

describe("SM-14 part A tool set and catalog scopes", () => {
  it("implements each as a v1.0 GET handler whose catalog scopes are exactly as recorded", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const c of TOOLS) {
      expect(handlers.has(c.name), c.name).toBe(true);
      const e = catalog.find((x) => x.toolName === c.name);
      expect(e?.method, c.name).toBe("get");
      expect(e?.["apiVersion"], c.name).not.toBe("beta");
      expect(e?.["workScopes"], c.name).toEqual(c.scopes);
    }
  });
});

describe.each(TOOLS)("$name: unavailable permission", (c) => {
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

describe.each(TOOLS.filter((c) => c.paged))("$name: permission lost mid-listing", (c) => {
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

describe("graph-batch: a denied SM-14 part A subrequest", () => {
  const batch = (url: string) =>
    call("graph-batch", { requests: [{ id: "1", method: "GET", url }] });

  it.each(TOOLS)(
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

  it.each(TOOLS)("$name: a 401 subresponse is reauth_required", async (c) => {
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

  it.each(TOOLS)(
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
