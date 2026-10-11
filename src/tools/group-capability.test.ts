import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";

/**
 * SM-10 capability tests for unavailable permissions, covering all eight group tools. They
 * mirror the SM-6 shared-mailbox precedent (mail-shared-delta.test.ts): a 403 is a
 * `capability_gap` carrying the "do not broaden scopes automatically" guidance after exactly
 * ONE request (no retry, no escalation), a 401 is `reauth_required`, and no token ever leaks,
 * even when Graph's own error text echoes it. Group-specific boundaries (hidden membership,
 * limited-information members, group kinds without a calendar or conversations) are covered
 * below. The base per-tool code mappings also exist in groups.test.ts, group-calendar.test.ts
 * and group-conversations.test.ts; this file adds the guidance/no-retry/no-leak assertions.
 * Graph's behaviour for group kinds without a calendar or conversations is undocumented, so
 * only OUR guarantee is asserted: failures surface through the normal error mapping and are
 * never turned into an empty successful list (D15).
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const GID = "123e4567-e89b-12d3-a456-426614174000";
const WIN = { startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2024-01-08T00:00:00Z" };
const USER_T = "#microsoft.graph.user";
const GROUP_T = "#microsoft.graph.group";

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

type Case = { name: string; args: Record<string, unknown>; paged: boolean };
const SM10_TOOLS: Case[] = [
  { name: "list-groups", args: {}, paged: true },
  { name: "get-group", args: { "group-id": GID }, paged: false },
  { name: "list-group-members", args: { "group-id": GID }, paged: true },
  { name: "list-group-owners", args: { "group-id": GID }, paged: true },
  { name: "get-group-calendar-view", args: { "group-id": GID, ...WIN }, paged: true },
  { name: "list-group-events", args: { "group-id": GID }, paged: true },
  { name: "list-group-conversations", args: { "group-id": GID }, paged: true },
  { name: "list-group-threads", args: { "group-id": GID }, paged: true },
];

describe("SM-10 tool set", () => {
  it("covers exactly the eight implemented group tools", () => {
    const names = catalog
      .filter(
        (e) =>
          e.method === "get" && /^\/groups(\/|$)/.test(e.pathPattern) && e.toolName !== undefined,
      )
      .map((e) => e.toolName)
      .sort();
    expect(names).toEqual(SM10_TOOLS.map((c) => c.name).sort());
  });
});

describe.each(SM10_TOOLS)("$name: unavailable permission", (c) => {
  it("403 is a capability_gap with the no-escalation guidance after exactly one request", async () => {
    responses.push(() =>
      json(
        {
          error: {
            code: "Authorization_RequestDenied",
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
    expect(env.error?.details).toMatchObject({
      status: 403,
      graphCode: "Authorization_RequestDenied",
    });
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

  it("a 403 never surfaces a partial or empty list", async () => {
    responses.push(() => json({ error: { code: "x" } }, 403));
    const { env } = await call(c.name, c.args);
    expect(env.data).toBeUndefined();
    expect(env.ok).toBe(false);
  });
});

describe.each(SM10_TOOLS.filter((c) => c.paged))("$name: permission lost mid-listing", (c) => {
  it("a 403 or 401 on a followed page fails the call instead of returning the first page", async () => {
    for (const [status, code] of [
      [403, "capability_gap"],
      [401, "reauth_required"],
    ] as const) {
      calls = [];
      const path = new URL(
        `https://graph.microsoft.com/v1.0${
          c.name === "list-groups" ? "/groups" : `/groups/${GID}/${nav(c.name)}`
        }`,
      );
      path.searchParams.set("$skiptoken", "abc");
      if (c.name === "get-group-calendar-view") {
        path.searchParams.set("startDateTime", WIN.startDateTime);
        path.searchParams.set("endDateTime", WIN.endDateTime);
      }
      responses.push(() =>
        json({
          value: [{ "@odata.type": USER_T, id: "1" }],
          "@odata.nextLink": path.toString(),
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

function nav(tool: string): string {
  const map: Record<string, string> = {
    "list-group-members": "members",
    "list-group-owners": "owners",
    "get-group-calendar-view": "calendarView",
    "list-group-events": "events",
    "list-group-conversations": "conversations",
    "list-group-threads": "threads",
  };
  return map[tool] as string;
}

describe("hidden-membership groups (Member.Read.Hidden is not requested)", () => {
  it("members 403 is a capability_gap with the guidance, one request, no token", async () => {
    responses.push(() =>
      json(
        {
          error: {
            code: "Authorization_RequestDenied",
            message: `Insufficient privileges to complete the operation. ${TOKEN}`,
          },
        },
        403,
      ),
    );
    const { env, text } = await call("list-group-members", { "group-id": GID });
    expect(env.error?.code).toBe("capability_gap");
    expect(env.error?.message).toContain("do not broaden scopes automatically");
    expect(calls).toHaveLength(1);
    expect(new URL(calls[0]?.url as string).pathname).toBe(`/v1.0/groups/${GID}/members`);
    expect(text).not.toContain(TOKEN);
  });

  it("never asks for hidden members itself (no extra headers, query or scope material)", async () => {
    await call("list-group-members", { "group-id": GID });
    expect(new URL(calls[0]?.url as string).search).toBe("");
    expect(Object.keys(calls[0]?.headers ?? {}).map((k) => k.toLowerCase())).not.toContain(
      "consistencylevel",
    );
  });
});

describe("limited-information directoryObject items (D13, GroupMember.ReadBasic.All)", () => {
  const sparse = [
    { "@odata.type": USER_T, id: "u1" },
    { "@odata.type": GROUP_T, id: "g1" },
    { "@odata.type": "#microsoft.graph.device", id: "d1" },
  ];

  it.each(["list-group-members", "list-group-owners"])(
    "%s keeps id + @odata.type items with and without a select, count unchanged",
    async (name) => {
      for (const select of [undefined, ["displayName", "mail"]]) {
        responses.push(() => json({ value: sparse }));
        const { env } = await call(name, { "group-id": GID, ...(select ? { select } : {}) });
        expect(env.ok).toBe(true);
        expect(env.data).toEqual(sparse);
        expect(env.meta).toMatchObject({ count: 3, hasMore: false });
      }
    },
  );

  it("an item without a type collapses to id only and is not dropped", async () => {
    responses.push(() => json({ value: [{ id: "x1" }] }));
    const { env } = await call("list-group-members", { "group-id": GID });
    expect(env.data).toEqual([{ id: "x1" }]);
  });
});

describe("group kinds without a calendar or conversations (D15: no undocumented error code asserted)", () => {
  const KINDS = SM10_TOOLS.filter((c) =>
    [
      "get-group-calendar-view",
      "list-group-events",
      "list-group-conversations",
      "list-group-threads",
    ].includes(c.name),
  );

  it.each(KINDS)(
    "$name: a Graph 404 or 400 is a graph_error, never an empty successful list",
    async (c) => {
      for (const status of [404, 400]) {
        calls = [];
        responses.push(() =>
          json(
            { error: { code: "SomeUndocumentedCode", message: "group has no such feature" } },
            status,
          ),
        );
        const { env, text } = await call(c.name, c.args);
        expect(env.ok, String(status)).toBe(false);
        expect(env.data).toBeUndefined();
        expect(env.error?.code).toBe("graph_error");
        expect(env.error?.details).toMatchObject({ status });
        expect(calls).toHaveLength(1);
        expect(text).not.toContain(TOKEN);
      }
    },
  );

  it.each(KINDS)("$name: a genuine empty 200 list stays a successful empty list", async (c) => {
    responses.push(() => json({ value: [] }));
    const { env } = await call(c.name, c.args);
    expect(env.ok).toBe(true);
    expect(env.data).toEqual([]);
  });

  it.each(KINDS)(
    "$name: a 200 without a value array is a failure, not an empty list",
    async (c) => {
      responses.push(() => json({}));
      const { env } = await call(c.name, c.args);
      expect(env.ok).toBe(false);
      expect(env.error?.code).toBe("graph_error");
    },
  );
});

describe("graph-batch: a denied group subrequest", () => {
  const batch = (url: string) =>
    call("graph-batch", { requests: [{ id: "1", method: "GET", url }] });

  it.each([`/groups/${GID}/members`, `/groups/${GID}/conversations`, `/groups/${GID}/threads`])(
    "%s: a 403 subresponse is a capability_gap without a body or token",
    async (url) => {
      responses.push(() =>
        json({
          responses: [
            {
              id: "1",
              status: 403,
              body: {
                error: {
                  code: "Authorization_RequestDenied",
                  message: "Insufficient privileges",
                  extra: "raw-body-marker",
                },
              },
            },
          ],
        }),
      );
      const { env, text } = await batch(url);
      const first = (env.data as { responses: Array<{ error?: { code: string }; body?: unknown }> })
        .responses[0];
      expect(first?.error?.code).toBe("capability_gap");
      expect(first?.body).toBeUndefined();
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
      expect(text).not.toContain("raw-body-marker");
    },
  );

  it("a batch response that echoes the token fails closed without leaking it", async () => {
    responses.push(() =>
      json({
        responses: [
          { id: "1", status: 403, body: { error: { code: "x", message: `echo ${TOKEN}` } } },
        ],
      }),
    );
    const { env, text } = await batch(`/groups/${GID}/members`);
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("graph_error");
    expect(text).not.toContain(TOKEN);
  });
});
