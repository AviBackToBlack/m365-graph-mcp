import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const CONV = "list-group-conversations";
const THREADS = "list-group-threads";
const GID = "123e4567-e89b-12d3-a456-426614174000";
const OTHER_GID = "223e4567-e89b-12d3-a456-426614174000";

interface Call {
  url: string;
  method: string;
  headers: Record<string, string>;
  body: string | undefined;
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
  error?: { code: string; message: string };
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
      body: typeof init.body === "string" ? init.body : undefined,
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

const BASE = ["id", "topic", "hasAttachments", "lastDeliveredDateTime", "uniqueSenders", "preview"];
const TOOLS = [
  { name: CONV, nav: "conversations", defaultSelect: BASE, extra: [] as string[] },
  {
    name: THREADS,
    nav: "threads",
    defaultSelect: [...BASE, "isLocked"],
    extra: ["toRecipients", "ccRecipients"],
  },
] as const;

describe("catalog and schema", () => {
  it("pins the two v1.0 GET entries and the Group-Conversation.Read.All scope (D16)", () => {
    for (const t of TOOLS) {
      const e = catalog.find((x) => x.toolName === t.name);
      expect(e?.pathPattern).toBe(`/groups/{group-id}/${t.nav}`);
      expect(e?.method).toBe("get");
      expect(e?.["apiVersion"]).toBeUndefined();
      expect(e?.["workScopes"]).toEqual(["Group-Conversation.Read.All"]);
      expect(String(e?.["llmTip"])).toMatch(/^Legacy/);
    }
  });

  it("describes the typed contracts, not raw Graph syntax (D17)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    for (const n of [CONV, THREADS]) {
      expect(desc(n)).toMatch(/^Legacy: .*Microsoft recommends Teams channel APIs/);
      expect(desc(n)).toMatch(/group-id is the group's GUID/);
      expect(desc(n)).toMatch(/Metadata-only retrieval/);
      expect(desc(n)).toMatch(/posts, post bodies and attachments are not returned/);
      expect(desc(n)).toMatch(/never sends Graph \$top/);
      expect(desc(n)).toMatch(/never converted into an empty successful result/);
      expect(desc(n)).toMatch(/Requires Group-Conversation\.Read\.All/);
      expect(desc(n)).not.toMatch(/Group\.Read\.All/);
    }
    expect(desc(CONV)).toMatch(
      /topic, hasAttachments, lastDeliveredDateTime, uniqueSenders and preview/,
    );
    expect(desc(THREADS)).toMatch(/isLocked/);
    expect(desc(THREADS)).toMatch(/opt-in via select and never part of the default projection/);
  });

  it("exposes only group-id, top and select", async () => {
    const tools = (await mcp.listTools()).tools;
    for (const t of TOOLS) {
      const props = Object.keys(
        (tools.find((x) => x.name === t.name)?.inputSchema.properties ?? {}) as Record<
          string,
          unknown
        >,
      ).sort();
      expect(props, t.name).toEqual(["group-id", "select", "top"]);
    }
  });

  it("does not register any post/thread-detail tool beyond the two catalog entries", async () => {
    const names = (await mcp.listTools()).tools.map((x) => x.name);
    expect(names.filter((n) => /group-(conversation|thread|post)/.test(n)).sort()).toEqual([
      CONV,
      THREADS,
    ]);
  });
});

describe.each(TOOLS)("$name", (t) => {
  it("sends a GET with the documented default select, a local top and no $top/headers/body", async () => {
    responses.push(() => json({ value: [{ id: "c1", topic: "t" }] }));
    const { env } = await call(t.name, { "group-id": GID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/groups/${GID}/${t.nav}`);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().searchParams.get("$select")).toBe(t.defaultSelect.join(","));
    expect([...parsed().searchParams.keys()]).toEqual(["$select"]);
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(calls[0]?.body).toBeUndefined();
    expect(env.data).toEqual([{ id: "c1", topic: "t" }]);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
  });

  it("projects to the effective select: bodies, posts and unselected fields are dropped", async () => {
    const over = {
      "@odata.context": "ctx",
      value: [
        {
          "@odata.type": "#microsoft.graph.conversation",
          "@odata.id": "evil",
          id: "c1",
          topic: "T",
          preview: "short excerpt",
          uniqueSenders: ["Alice"],
          toRecipients: [{ emailAddress: { name: "RECIP", address: "r@x.com" } }],
          ccRecipients: [{ emailAddress: { name: "CCRECIP", address: "c@x.com" } }],
          body: { content: "SECRET-BODY" },
          posts: [{ id: "p" }],
          threads: [{ id: "t" }],
        },
        "not-an-object",
      ],
    };
    responses.push(() => json(over));
    const def = await call(t.name, { "group-id": GID });
    expect(def.env.data).toEqual([
      {
        "@odata.type": "#microsoft.graph.conversation",
        id: "c1",
        topic: "T",
        preview: "short excerpt",
        uniqueSenders: ["Alice"],
      },
    ]);
    for (const s of ["SECRET-BODY", "posts", "threads", "RECIP", "evil", "ctx"]) {
      expect(def.text).not.toContain(s);
    }
    responses.push(() => json(over));
    const only = await call(t.name, { "group-id": GID, select: ["id", "topic"] });
    expect(parsed(1).searchParams.get("$select")).toBe("id,topic");
    expect(only.env.data).toEqual([
      { "@odata.type": "#microsoft.graph.conversation", id: "c1", topic: "T" },
    ]);
  });

  it("applies the local top limit without sending $top and follows same-collection pages", async () => {
    const link = `https://graph.microsoft.com/v1.0/groups/${GID}/${t.nav}?$skiptoken=abc`;
    responses.push(() => json({ value: [{ id: "1" }, { id: "2" }], "@odata.nextLink": link }));
    responses.push(() => json({ value: [{ id: "3" }, { id: "4" }] }));
    const { env } = await call(t.name, { "group-id": GID, top: 3 });
    expect((env.data as unknown[]).length).toBe(3);
    expect(env.meta?.["hasMore"]).toBe(true);
    expect(calls).toHaveLength(2);
    expect(parsed(0).searchParams.has("$top")).toBe(false);
    expect(parsed(1).pathname).toBe(`/v1.0/groups/${GID}/${t.nav}`);
  });

  it("rejects options outside the contract with zero network calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      { filter: "hasAttachments eq true" },
      { search: "x" },
      { orderby: [{ property: "lastDeliveredDateTime" }] },
      { skip: 1 },
      { count: true },
      { expand: ["posts"] },
      { expand: ["threads"] },
      { headers: { ConsistencyLevel: "eventual" } },
      { Prefer: "outlook.allow-unsafe-html" },
      { top: 0 },
      { top: 101 },
      { select: ["body"] },
      { select: ["posts"] },
      { select: ["threads"] },
      { select: ["uniqueBody"] },
      { select: [] },
      { select: "id" },
      { unknownKey: 1 },
      // recipients exist on threads only
      ...(t.extra.length === 0 ? [{ select: ["toRecipients"] }, { select: ["isLocked"] }] : []),
    ];
    for (const extra of bad) {
      calls = [];
      const { env, text } = await call(t.name, { "group-id": GID, ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("rejects non-GUID group ids and path injection with zero network calls", async () => {
    for (const id of [
      "delta",
      "$ref",
      "$count",
      "..",
      "alice@contoso.com", // a UPN is a user-id shape, not a group id
      `${GID}/members`,
      `${GID}/${t.nav}`,
      `${GID}/threads/x/posts`,
      `${GID}%2Fmembers`,
      `${GID}%2F..%2Fusers`,
      `${GID}?$top=1`,
      `${GID}#x`,
      "microsoft.graph.getByIds()",
      "",
    ]) {
      calls = [];
      const { env } = await call(t.name, { "group-id": id });
      expect(env.ok, id).toBe(false);
      expect(env.error?.code, id).toBe("invalid_input");
      expect(calls, id).toHaveLength(0);
    }
    const missing = await call(t.name, {});
    expect(missing.env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("refuses a nextLink for another group, collection, host, sub-resource or API version (fail closed)", async () => {
    const other = t.nav === "threads" ? "conversations" : "threads";
    for (const bad of [
      `https://evil.example/v1.0/groups/${GID}/${t.nav}?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${OTHER_GID}/${t.nav}?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/${other}?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/${t.nav}/someId/posts`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/${t.nav}/someId`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/members`,
      `https://graph.microsoft.com/v1.0/groups/${GID}`,
      `https://graph.microsoft.com/v1.0/me/${t.nav}?$skiptoken=1`,
      `https://graph.microsoft.com/beta/groups/${GID}/${t.nav}?$skiptoken=1`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": bad }));
      const { env } = await call(t.name, { "group-id": GID, top: 5 });
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "x" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(t.name, { "group-id": GID })).env.error?.code).toBe("graph_error");
    }
  });

  it("maps 401, 403, 404 and 400 to the shared conventions without echoing the token", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
      [400, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: `m ${TOKEN}` } }, status));
      const { env, text } = await call(t.name, { "group-id": GID });
      expect(env.ok).toBe(false);
      expect(env.error?.code, String(status)).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe(THREADS, () => {
  it("offers recipients on explicit select only and keeps them then", async () => {
    responses.push(() =>
      json({
        value: [{ id: "t1", toRecipients: [{ emailAddress: { name: "G", address: "g@x.com" } }] }],
      }),
    );
    const { env } = await call(THREADS, {
      "group-id": GID,
      select: ["id", "toRecipients", "ccRecipients"],
    });
    expect(parsed().searchParams.get("$select")).toBe("id,toRecipients,ccRecipients");
    expect(env.data).toEqual([
      { id: "t1", toRecipients: [{ emailAddress: { name: "G", address: "g@x.com" } }] },
    ]);
  });

  it("keeps isLocked in the default select", async () => {
    await call(THREADS, { "group-id": GID });
    expect((parsed().searchParams.get("$select") ?? "").split(",")).toContain("isLocked");
    expect((parsed().searchParams.get("$select") ?? "").split(",")).not.toContain("toRecipients");
  });
});

describe("graph-batch reachability", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const ok = (body: unknown) => () => json({ responses: [{ id: "1", status: 200, body }] });

  it("registers both tools with a batch policy before graph-batch", () => {
    const handlers = createHandlers(
      new GraphClient({ tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) }, catalog }),
      catalog,
    );
    const names = [...handlers.keys()];
    for (const n of [CONV, THREADS]) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
      expect(names.indexOf(n)).toBeLessThan(names.indexOf("graph-batch"));
    }
  });

  it("accepts in-policy subrequests and projects like the single path", async () => {
    for (const nav of ["conversations", "threads"]) {
      for (const q of ["", "?$select=id,topic"]) {
        calls = [];
        responses.push(
          ok({
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/x",
            value: [{ id: "a", topic: "t", body: { content: "SECRET-BODY" }, posts: [1] }],
          }),
        );
        const { env, text } = await batch(`/groups/${GID}/${nav}${q}`);
        expect(env.ok, `${nav}${q}`).toBe(true);
        expect(calls).toHaveLength(1);
        expect(calls[0]?.url).toContain("/$batch");
        expect(text).not.toContain("SECRET-BODY");
        expect(text).not.toContain("posts");
        expect(text).not.toContain("nextLink");
      }
    }
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const c = `/groups/${GID}/conversations`;
    const t = `/groups/${GID}/threads`;
    const bad: Array<[string, Record<string, string>?]> = [
      [`${c}?$top=5`], // $top is never sent to Graph for these (not documented)
      [`${t}?$top=5`],
      [`${c}?$filter=hasAttachments%20eq%20true`],
      [`${c}?$search=%22x%22`],
      [`${c}?$orderby=lastDeliveredDateTime`],
      [`${t}?$skip=1`],
      [`${t}?$count=true`],
      [`${c}?$expand=threads`],
      [`${t}?$expand=posts`],
      [`${c}?$select=body`],
      [`${t}?$select=posts`],
      [`${c}?$select=toRecipients`], // threads only
      [`${c}?$select=isLocked`],
      [`${t}?$select=id&$select=topic`],
      [c, { ConsistencyLevel: "eventual" }],
      [t, { Prefer: "outlook.allow-unsafe-html" }],
      // sub-paths and look-alikes that are not catalog entries
      [`${c}/someId`],
      [`${c}/someId/threads`],
      [`${t}/someId`],
      [`${t}/someId/posts`],
      [`${t}/someId/reply`],
      [`${c}/$count`],
      [`${t}/$ref`],
      [`${t}/delta`],
      [`/groups/delta/conversations`],
      [`/groups/delta/threads`],
      [`/groups/$ref/threads`],
      [`/groups/../threads`],
      [`/groups/${GID}%2Fmembers/threads`],
      [`/groups/alice@contoso.com/conversations`],
      [`${t}#frag`],
    ];
    for (const [url, headers] of bad) {
      const { env, text } = await batch(url, headers);
      expect(env.ok, `${url} ${JSON.stringify(headers)}`).toBe(false);
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });
});
