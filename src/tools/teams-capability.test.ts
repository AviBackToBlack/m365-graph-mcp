import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";

/**
 * SM-12 permission/capability matrix, covering all ten Teams/channel tools. Mirrors
 * group-capability.test.ts: a 403 is a `capability_gap` carrying the "do not broaden scopes
 * automatically" guidance after exactly ONE request (no retry, no escalation), a 401 is
 * `reauth_required`, a 401/403 on a followed page fails the whole call, no token ever leaks (even when
 * Graph echoes it, single path and batch), and a Graph failure never becomes an empty successful
 * list while a genuine empty 200 stays a success. The matrix of scopes/boundaries is documented in
 * PERMISSIONS.md ("SM-12 permission/capability matrix"). Per-tool shape, projection, option and
 * pagination cases live in teams.test.ts, channel-messages.test.ts and channel-tabs.test.ts; the
 * channel-message tools also keep their own per-tool 401/403 cases in channel-messages.test.ts
 * (this file adds the 403-without-JSON, empty-vs-failure and batch/echo cases for all ten).
 * Graph's documented gaps (members 401 in new tenants, shared channels, unknownFutureValue) are not
 * asserted as specific Graph behaviour: only OUR guarantee is, namely that failures surface through
 * the normal error mapping.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const TID = "f17e0000-0000-4000-8000-000000000020";
const CID = "19:4a95f7d8db4c4e7fae857bcebe0623e6@thread.tacv2";
const ECID = encodeURIComponent(CID);
const MID = "1616963377068";
const CH = `/teams/${TID}/channels/${CID}`;
const ECH = `/teams/${TID}/channels/${ECID}`;

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
  /** Collection tool: pages are followed. */
  paged: boolean;
  /** Absolute collection path a followed page must stay on (paged only). */
  path: string;
  /** The batch subrequest URL for this tool. */
  batchUrl: string;
};
const T = { "team-id": TID };
const C = { ...T, "channel-id": CID };
const M = { ...C, "chatMessage-id": MID };
const SM12_TOOLS: Case[] = [
  { name: "get-team", args: T, paged: false, path: "", batchUrl: `/teams/${TID}` },
  {
    name: "list-team-channels",
    args: T,
    paged: true,
    path: `/teams/${TID}/channels`,
    // requiredInBatch: graph-batch needs an explicit allowlisted $select for the channel tools.
    batchUrl: `/teams/${TID}/channels?$select=displayName`,
  },
  {
    name: "get-team-channel",
    args: C,
    paged: false,
    path: "",
    batchUrl: `${ECH}?$select=displayName`,
  },
  {
    name: "list-team-members",
    args: T,
    paged: true,
    path: `/teams/${TID}/members`,
    batchUrl: `/teams/${TID}/members`,
  },
  {
    name: "list-channel-messages",
    args: C,
    paged: true,
    path: `${CH}/messages`,
    batchUrl: `${ECH}/messages`,
  },
  {
    name: "get-channel-message",
    args: M,
    paged: false,
    path: "",
    batchUrl: `${ECH}/messages/${MID}`,
  },
  {
    name: "list-channel-message-replies",
    args: M,
    paged: true,
    path: `${CH}/messages/${MID}/replies`,
    batchUrl: `${ECH}/messages/${MID}/replies`,
  },
  {
    name: "list-channel-message-hosted-contents",
    args: M,
    paged: true,
    path: `${CH}/messages/${MID}/hostedContents`,
    batchUrl: `${ECH}/messages/${MID}/hostedContents`,
  },
  {
    name: "list-channel-tabs",
    args: C,
    paged: true,
    path: `${CH}/tabs`,
    batchUrl: `${ECH}/tabs`,
  },
  {
    name: "get-channel-files-folder",
    args: C,
    paged: false,
    path: "",
    batchUrl: `${ECH}/filesFolder`,
  },
];

describe("SM-12 tool set", () => {
  it("covers exactly the ten implemented Teams/channel tools", () => {
    const names = catalog
      .filter(
        (e) =>
          e.method === "get" &&
          e.toolName !== undefined &&
          /^\/teams\/\{team-id\}(\/(channels|members)(\/|$)|$)/.test(e.pathPattern),
      )
      .map((e) => e.toolName)
      .sort();
    expect(names).toEqual(SM12_TOOLS.map((c) => c.name).sort());
  });

  it("implements every one of them as a v1.0 GET handler", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const c of SM12_TOOLS) {
      expect(handlers.has(c.name), c.name).toBe(true);
      const e = catalog.find((x) => x.toolName === c.name);
      expect(e?.method, c.name).toBe("get");
      expect(e?.["apiVersion"], c.name).not.toBe("beta");
    }
  });
});

describe.each(SM12_TOOLS)("$name: unavailable permission", (c) => {
  it("403 is a capability_gap with the no-escalation guidance after exactly one request", async () => {
    responses.push(() =>
      json(
        {
          error: {
            code: "Forbidden",
            message: `Insufficient privileges to complete the operation. Bearer ${TOKEN}`,
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
    expect(env.error?.details).toMatchObject({ status: 403, graphCode: "Forbidden" });
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
        json({ error: { code: "NotFound", message: `no such thing ${TOKEN}` } }, status),
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

describe.each(SM12_TOOLS.filter((c) => c.paged))("$name: permission lost mid-listing", (c) => {
  it("a 403 or 401 on a followed page fails the call instead of returning the first page", async () => {
    for (const [status, code] of [
      [403, "capability_gap"],
      [401, "reauth_required"],
    ] as const) {
      calls = [];
      responses.push(() =>
        json({
          value: [{ id: "1" }],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0${c.path}?$skiptoken=abc`,
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

describe("graph-batch: a denied SM-12 subrequest", () => {
  const batch = (url: string) =>
    call("graph-batch", { requests: [{ id: "1", method: "GET", url }] });

  it.each(SM12_TOOLS)(
    "$name: a 403 subresponse is a capability_gap without a body or token",
    async (c) => {
      responses.push(() =>
        json({
          responses: [
            {
              id: "1",
              status: 403,
              body: {
                error: {
                  code: "Forbidden",
                  message: "Insufficient privileges",
                  extra: "raw-body-marker",
                },
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

  it.each(SM12_TOOLS)("$name: a 401 subresponse is reauth_required", async (c) => {
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

  it.each(SM12_TOOLS)(
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
