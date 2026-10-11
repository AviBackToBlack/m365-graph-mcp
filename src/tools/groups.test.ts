import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { GROUP_DEFAULT_PROPERTIES, GROUP_PROPERTIES, projectGroupMember } from "./groups.js";
import { createHandlers } from "./index.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const LIST = "list-groups";
const GET = "get-group";
const MEMBERS = "list-group-members";
const OWNERS = "list-group-owners";
const GID = "123e4567-e89b-12d3-a456-426614174000";
const OTHER_GID = "223e4567-e89b-12d3-a456-426614174000";
const USER_T = "#microsoft.graph.user";
const GROUP_T = "#microsoft.graph.group";
const CONTACT_T = "#microsoft.graph.orgContact";
const DEVICE_T = "#microsoft.graph.device";
const SP_T = "#microsoft.graph.servicePrincipal";

interface Call {
  url: string;
  method: string;
  headers: Record<string, string>;
  body?: string | undefined;
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
      body: typeof init.body === "string" ? init.body : undefined,
      headers: { ...(init.headers as Record<string, string>) },
    });
    const next = responses.shift();
    return Promise.resolve(next ? next() : json({}));
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

describe("catalog", () => {
  it("pins the four v1.0 GET entries and their scopes (D12, D13)", () => {
    const by = (n: string) => catalog.find((e) => e.toolName === n);
    const expected: Array<[string, string, string]> = [
      // D12: Group.Read.All preserves the group-property contract.
      [LIST, "/groups", "Group.Read.All"],
      [GET, "/groups/{group-id}", "Group.Read.All"],
      // D13: documented least privilege; per-type member details are best-effort.
      [MEMBERS, "/groups/{group-id}/members", "GroupMember.ReadBasic.All"],
      [OWNERS, "/groups/{group-id}/owners", "GroupMember.Read.All"],
    ];
    for (const [name, path, scope] of expected) {
      expect(by(name)?.pathPattern).toBe(path);
      expect(by(name)?.method).toBe("get");
      expect(by(name)?.["apiVersion"]).toBeUndefined();
      expect(by(name)?.["workScopes"]).toEqual([scope]);
    }
  });

  it("exposes only the typed options", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = (n: string) =>
      Object.keys(
        (tools.find((t) => t.name === n)?.inputSchema.properties ?? {}) as Record<string, unknown>,
      ).sort();
    expect(props(LIST)).toEqual(["search", "select", "top"]);
    expect(props(GET)).toEqual(["group-id", "select"]);
    expect(props(MEMBERS)).toEqual(["group-id", "select", "top"]);
    expect(props(OWNERS)).toEqual(["group-id", "select", "top"]);
  });

  it("describes the typed contracts, not raw Graph syntax (D14)", async () => {
    const by = (n: string) => catalog.find((e) => e.toolName === n);
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    for (const n of [LIST, GET, MEMBERS, OWNERS]) {
      const tip = String(by(n)?.["llmTip"]);
      expect(tip).not.toMatch(
        /Use \$(filter|search|select|top)|groupTypes\/any|ConsistencyLevel: eventual header/,
      );
      expect(desc(n)).toContain("Typed");
      expect(desc(n)).not.toMatch(/\$filter=|groupTypes\/any/);
    }
    expect(desc(LIST)).toMatch(/exactly one displayName:text or description:text clause/);
    expect(desc(LIST)).toMatch(/\$count=true and ConsistencyLevel: eventual/);
    expect(desc(LIST)).toMatch(/not available through graph-batch/);
    expect(desc(GET)).toMatch(/no \$expand/);
    expect(desc(MEMBERS)).toMatch(/no \$top is sent/);
    expect(desc(MEMBERS)).toMatch(/best-effort/);
    expect(desc(MEMBERS)).toMatch(/Member\.Read\.Hidden/);
    expect(desc(OWNERS)).toMatch(/empty list can be correct/);
  });
});

describe(LIST, () => {
  it("sends the default top and select and no advanced-query material", async () => {
    responses.push(() => json({ value: [{ id: "g1", displayName: "A" }] }));
    const { env } = await call(LIST);
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe("/v1.0/groups");
    expect(parsed().searchParams.get("$top")).toBe("25");
    expect(parsed().searchParams.get("$select")).toBe(GROUP_DEFAULT_PROPERTIES.join(","));
    expect([...parsed().searchParams.keys()].sort()).toEqual(["$select", "$top"]);
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(calls[0]?.method).toBe("GET");
  });

  it("sets $search, $count and ConsistencyLevel itself, only when a search is given", async () => {
    responses.push(() => json({ value: [] }));
    await call(LIST, { search: "description:marketing team", top: 10, select: ["id", "mail"] });
    const u = parsed();
    expect(u.searchParams.get("$search")).toBe('"description:marketing team"');
    expect(u.searchParams.get("$count")).toBe("true");
    expect(u.searchParams.get("$top")).toBe("10");
    expect(u.searchParams.get("$select")).toBe("id,mail");
    expect(calls[0]?.headers["ConsistencyLevel"]).toBe("eventual");
  });

  it("pages in at most 100 and keeps ConsistencyLevel on follow-up pages", async () => {
    const link = "https://graph.microsoft.com/v1.0/groups?$skiptoken=abc";
    responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": link }));
    responses.push(() => json({ value: [{ id: "2" }] }));
    const { env } = await call(LIST, { search: "displayName:x", top: 150 });
    expect(env.ok).toBe(true);
    expect(parsed().searchParams.get("$top")).toBe("100");
    expect(calls).toHaveLength(2);
    expect(calls[1]?.headers["ConsistencyLevel"]).toBe("eventual");
  });

  it("rejects unknown options, filter/orderby/skip/count/expand, raw headers and unsafe values", async () => {
    const bad: Array<Record<string, unknown>> = [
      { filter: "groupTypes/any(c:c eq 'Unified')" },
      { orderby: [{ property: "displayName" }] },
      { skip: 1 },
      { count: true },
      { expand: ["members"] },
      { headers: { ConsistencyLevel: "eventual" } },
      { ConsistencyLevel: "eventual" },
      { $search: "x" },
      { top: 0 },
      { top: 201 },
      { select: ["assignedLicenses"] },
      { select: ["accessType"] },
      { select: ["unseenCount"] },
      { select: ["onPremisesSecurityIdentifier"] },
      { select: ["members"] },
      { select: [] },
      { search: "marketing" }, // no property
      { search: "mail:marketing" }, // property not allowlisted
      { search: 'displayName:mark"eting' },
      { search: "displayName:mark\\eting" },
      { search: "displayName:a:b" },
      { search: "displayName:a OR description:b" },
      { search: "displayName:a AND b" },
      { search: "displayName:NOT a" },
      { search: "displayName:(a)" },
      { search: "displayName:R&D" },
      { search: "displayName:a\nb" },
      { search: "displayName:" },
      { search: "displayName:***" },
      { search: `displayName:${"x".repeat(200)}` },
      { unknownKey: 1 },
    ];
    for (const args of bad) {
      calls = [];
      const { env, text } = await call(LIST, args);
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code, JSON.stringify(args)).toBe("invalid_input");
      expect(calls, JSON.stringify(args)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("keeps only the request's effective $select and drops annotations except @odata.type", async () => {
    const over = {
      "@odata.context": "x",
      value: [
        {
          id: "g1",
          displayName: "A",
          mail: "a@contoso.com",
          proxyAddresses: ["SMTP:a@contoso.com"],
          assignedLicenses: [{ skuId: "s" }],
          "@odata.type": GROUP_T,
          "@odata.id": "evil",
        },
        "not-an-object",
      ],
    };
    responses.push(() => json(over));
    const def = await call(LIST);
    expect(def.env.data).toEqual([
      { id: "g1", displayName: "A", mail: "a@contoso.com", "@odata.type": GROUP_T },
    ]);
    for (const s of ["proxyAddresses", "assignedLicenses", "evil"]) {
      expect(def.text).not.toContain(s);
    }
    responses.push(() => json(over));
    expect((await call(LIST, { select: ["id"] })).env.data).toEqual([
      { id: "g1", "@odata.type": GROUP_T },
    ]);
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "backend failure" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(LIST)).env.error?.code, JSON.stringify(body)).toBe("graph_error");
    }
  });

  it("refuses a nextLink that leaves the endpoint or origin (fail closed)", async () => {
    for (const bad of [
      "https://evil.example/v1.0/groups?$skiptoken=1",
      `https://graph.microsoft.com/v1.0/groups/${GID}/members`,
      "https://graph.microsoft.com/v1.0/users",
      "https://graph.microsoft.com/beta/groups",
    ]) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": bad }));
      const { env } = await call(LIST, { top: 5 });
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("maps 401 and 403 to the shared error conventions", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      const { env, text } = await call(LIST);
      expect(env.error?.code).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe(GET, () => {
  it("sends the default select to /groups/{id} with no other options or headers", async () => {
    responses.push(() => json({ id: GID, displayName: "HR", visibility: "Private" }));
    const { env } = await call(GET, { "group-id": GID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/groups/${GID}`);
    expect(parsed().searchParams.get("$select")).toBe(GROUP_DEFAULT_PROPERTIES.join(","));
    expect([...parsed().searchParams.keys()]).toEqual(["$select"]);
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(env.data).toEqual({ id: GID, displayName: "HR", visibility: "Private" });
  });

  it("projects to the request's select; over-returned and non-allowlisted fields are dropped", async () => {
    const body = {
      "@odata.context": "ctx",
      id: GID,
      displayName: "HR",
      mail: "hr@contoso.com",
      membershipRule: 'user.department -eq "HR"',
      proxyAddresses: ["SMTP:hr@contoso.com"],
      assignedLicenses: [{ skuId: "s" }],
      unseenCount: 3,
    };
    responses.push(() => json(body));
    const { env, text } = await call(GET, { "group-id": GID, select: ["mail"] });
    expect(parsed().searchParams.get("$select")).toBe("mail");
    expect(env.data).toEqual({ id: GID, mail: "hr@contoso.com" });
    for (const s of ["proxyAddresses", "assignedLicenses", "unseenCount", "ctx"]) {
      expect(text).not.toContain(s);
    }
    responses.push(() => json(body));
    const def = (await call(GET, { "group-id": GID })).env.data as Record<string, unknown>;
    expect(def["membershipRule"]).toBe('user.department -eq "HR"');
    expect(def["proxyAddresses"]).toBeUndefined();
  });

  it("rejects non-GUID group ids and unknown options with zero network calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      {},
      { "group-id": "delta" },
      { "group-id": "$ref" },
      { "group-id": "$count" },
      { "group-id": ".." },
      { "group-id": `${GID}/members` },
      { "group-id": `${GID}%2Fmembers` },
      { "group-id": `${GID}?$select=x` },
      { "group-id": `${GID}#x` },
      { "group-id": "microsoft.graph.getByIds()" },
      { "group-id": GID, expand: ["members"] },
      { "group-id": GID, select: ["accessType"] },
      { "group-id": GID, select: ["hasMembersWithLicenseErrors"] },
      { "group-id": GID, filter: "x" },
    ];
    for (const args of bad) {
      calls = [];
      const { env } = await call(GET, args);
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code, JSON.stringify(args)).toBe("invalid_input");
      expect(calls, JSON.stringify(args)).toHaveLength(0);
    }
  });

  it("fails closed on a non-object body", async () => {
    for (const body of [[{ id: "x" }], "text", null]) {
      responses.push(() => json(body));
      expect((await call(GET, { "group-id": GID })).env.error?.code).toBe("graph_error");
    }
  });

  it("maps 401, 403 and 404", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      expect((await call(GET, { "group-id": GID })).env.error?.code).toBe(code);
    }
  });
});

describe.each([
  [MEMBERS, "members"],
  [OWNERS, "owners"],
] as const)("%s", (name, nav) => {
  it("sends a bare GET (no $top, no $select unless chosen, no headers)", async () => {
    responses.push(() => json({ value: [{ "@odata.type": USER_T, id: "u1", displayName: "A" }] }));
    const { env } = await call(name, { "group-id": GID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/groups/${GID}/${nav}`);
    expect(parsed().search).toBe("");
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
  });

  it("projects each item per @odata.type; unknown types keep id + @odata.type only", async () => {
    responses.push(() =>
      json({
        "@odata.context": "ctx",
        value: [
          {
            "@odata.type": USER_T,
            id: "u1",
            displayName: "User",
            jobTitle: "T",
            passwordProfile: { p: 1 },
            assignedLicenses: [{ s: 1 }],
          },
          {
            "@odata.type": GROUP_T,
            id: "g1",
            displayName: "Nested",
            groupTypes: ["Unified"],
            proxyAddresses: ["SMTP:x"],
            jobTitle: "nope",
          },
          {
            "@odata.type": CONTACT_T,
            id: "c1",
            displayName: "Contact",
            companyName: "Ext",
            userPrincipalName: "nope@x.com",
          },
          { "@odata.type": DEVICE_T, id: "d1", displayName: "Laptop", deviceId: "secret-dev" },
          { "@odata.type": SP_T, id: "s1", displayName: "App", appId: "secret-app" },
          { "@odata.type": "#microsoft.graph.somethingNew", id: "n1", displayName: "New" },
          "not-an-object",
        ],
      }),
    );
    const { env, text } = await call(name, { "group-id": GID });
    expect(env.data).toEqual([
      { "@odata.type": USER_T, id: "u1", displayName: "User", jobTitle: "T" },
      { "@odata.type": GROUP_T, id: "g1", displayName: "Nested", groupTypes: ["Unified"] },
      { "@odata.type": CONTACT_T, id: "c1", displayName: "Contact", companyName: "Ext" },
      { "@odata.type": DEVICE_T, id: "d1" },
      { "@odata.type": SP_T, id: "s1" },
      { "@odata.type": "#microsoft.graph.somethingNew", id: "n1" },
    ]);
    for (const s of ["passwordProfile", "assignedLicenses", "proxyAddresses", "secret-", "nope"]) {
      expect(text).not.toContain(s);
    }
  });

  it("restricts to the request's $select per type (user-only fields never leak onto a group)", async () => {
    responses.push(() =>
      json({
        value: [
          { "@odata.type": USER_T, id: "u1", displayName: "U", jobTitle: "T", mail: "u@x.com" },
          { "@odata.type": GROUP_T, id: "g1", displayName: "G", jobTitle: "T", mail: "g@x.com" },
        ],
      }),
    );
    const { env } = await call(name, { "group-id": GID, select: ["displayName", "jobTitle"] });
    expect(parsed().searchParams.get("$select")).toBe("displayName,jobTitle");
    expect(env.data).toEqual([
      { "@odata.type": USER_T, id: "u1", displayName: "U", jobTitle: "T" },
      { "@odata.type": GROUP_T, id: "g1", displayName: "G" },
    ]);
  });

  it("applies the local top limit without sending $top and follows pages", async () => {
    const link = `https://graph.microsoft.com/v1.0/groups/${GID}/${nav}?$skiptoken=abc`;
    const page = (ids: string[], next?: string) =>
      json({
        value: ids.map((id) => ({ "@odata.type": USER_T, id })),
        ...(next === undefined ? {} : { "@odata.nextLink": next }),
      });
    responses.push(() => page(["1", "2"], link));
    responses.push(() => page(["3", "4"]));
    const { env } = await call(name, { "group-id": GID, top: 3 });
    expect((env.data as unknown[]).length).toBe(3);
    expect(env.meta?.["hasMore"]).toBe(true);
    expect(calls).toHaveLength(2);
    expect(parsed(0).searchParams.has("$top")).toBe(false);
  });

  it("rejects unknown options, group-id escapes and bad bounds with zero network calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      { "group-id": GID, filter: "x" },
      { "group-id": GID, search: "displayName:x" },
      { "group-id": GID, orderby: [{ property: "displayName" }] },
      { "group-id": GID, skip: 1 },
      { "group-id": GID, count: true },
      { "group-id": GID, expand: ["members"] },
      { "group-id": GID, headers: { ConsistencyLevel: "eventual" } },
      { "group-id": GID, top: 0 },
      { "group-id": GID, top: 101 },
      { "group-id": GID, select: ["passwordProfile"] },
      { "group-id": GID, select: ["assignedLicenses"] },
      { "group-id": GID, select: ["accessType"] },
      { "group-id": GID, select: [] },
      { "group-id": "delta" },
      { "group-id": "$ref" },
      { "group-id": ".." },
      { "group-id": `${GID}/transitiveMembers` },
      { "group-id": `${GID}%2F..%2Fusers` },
      { "group-id": `${GID}?$top=1` },
      { "group-id": "microsoft.graph.getByIds()" },
      { "group-id": "alice@contoso.com" }, // a UPN is a user-id shape, not a group id
      {},
    ];
    for (const args of bad) {
      calls = [];
      const { env, text } = await call(name, args);
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code, JSON.stringify(args)).toBe("invalid_input");
      expect(calls, JSON.stringify(args)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("refuses a nextLink for another group, collection, host or API version (fail closed)", async () => {
    const other = nav === "members" ? "owners" : "members";
    for (const bad of [
      "https://evil.example/v1.0/groups/" + GID + "/" + nav + "?$skiptoken=1",
      `https://graph.microsoft.com/v1.0/groups/${OTHER_GID}/${nav}?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/${other}?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/transitiveMembers?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/${nav}/microsoft.graph.user`,
      `https://graph.microsoft.com/beta/groups/${GID}/${nav}`,
      "https://graph.microsoft.com/v1.0/users",
    ]) {
      calls = [];
      responses.push(() =>
        json({ value: [{ "@odata.type": USER_T, id: "1" }], "@odata.nextLink": bad }),
      );
      const { env } = await call(name, { "group-id": GID, top: 5 });
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "x" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(name, { "group-id": GID })).env.error?.code).toBe("graph_error");
    }
  });

  it("maps 401, 403 (e.g. hidden membership) and 404", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      const { env, text } = await call(name, { "group-id": GID });
      expect(env.error?.code).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe("projectGroupMember", () => {
  it("keeps only id when the type is missing (fails closed, even for a selected field)", () => {
    const typeless = {
      id: "1",
      displayName: "A",
      mail: "a@x.com",
      proxyAddresses: ["smtp:a@x.com"],
      userPrincipalName: "u",
    };
    expect(projectGroupMember(typeless, "")).toEqual({ id: "1" });
    expect(projectGroupMember(typeless, "displayName,mail,proxyAddresses")).toEqual({ id: "1" });
  });

  it("drops non-objects and always keeps id and @odata.type", () => {
    for (const v of [null, undefined, 1, "x", [1]])
      expect(projectGroupMember(v, undefined)).toBe(undefined);
    expect(projectGroupMember({ "@odata.type": GROUP_T, id: "g" }, "displayName")).toEqual({
      "@odata.type": GROUP_T,
      id: "g",
    });
    expect(
      projectGroupMember({ "@odata.type": DEVICE_T, id: "d", displayName: "x" }, "displayName"),
    ).toEqual({
      "@odata.type": DEVICE_T,
      id: "d",
    });
  });

  it("exposes group properties that are returned by default only", () => {
    for (const forbidden of ["accessType", "unseenCount", "assignedLicenses", "proxyAddresses"]) {
      expect((GROUP_PROPERTIES as readonly string[]).includes(forbidden)).toBe(false);
    }
  });
});

describe("graph-batch reachability", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const ok = (body: unknown) => () => json({ responses: [{ id: "1", status: 200, body }] });

  it("registers all four tools with a batch policy", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const n of [LIST, GET, MEMBERS, OWNERS]) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
    }
  });

  it("accepts in-policy forms and projects like the single path", async () => {
    responses.push(ok({ value: [{ id: "g1", displayName: "A", assignedLicenses: [{ s: 1 }] }] }));
    const a = await batch("/groups?$top=5&$select=id,displayName");
    expect(a.env.ok).toBe(true);
    expect(a.text).not.toContain("assignedLicenses");

    responses.push(ok({ id: GID, displayName: "HR", proxyAddresses: ["SMTP:x"] }));
    const g = await batch(`/groups/${GID}`);
    expect(g.env.ok).toBe(true);
    expect(g.text).not.toContain("proxyAddresses");

    for (const nav of ["members", "owners"]) {
      responses.push(
        ok({
          value: [
            { "@odata.type": USER_T, id: "u", displayName: "U", passwordProfile: {} },
            { "@odata.type": DEVICE_T, id: "d", displayName: "secret-device" },
          ],
        }),
      );
      const r = await batch(`/groups/${GID}/${nav}`);
      expect(r.env.ok, nav).toBe(true);
      expect(r.text).not.toContain("passwordProfile");
      expect(r.text).not.toContain("secret-device");
    }
  });

  it("clips a members batch page to the same default the single call uses", async () => {
    const page = {
      value: Array.from({ length: 30 }, (_, i) => ({ "@odata.type": USER_T, id: `u${i}` })),
    };
    responses.push(() => json(page));
    const single = (await call(MEMBERS, { "group-id": GID })).env.data as unknown[];
    responses.push(ok(page));
    const b = await batch(`/groups/${GID}/members`);
    const body = (
      b.env.data as { responses: Array<{ body?: { value?: unknown[]; hasMore?: boolean } }> }
    ).responses[0]?.body;
    expect(body?.value).toHaveLength(single.length);
    expect(body?.hasMore).toBe(true);
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const bad: Array<[string, Record<string, string>?]> = [
      [
        "/groups?$search=%22displayName%3Amarketing%22&$count=true",
        { ConsistencyLevel: "eventual" },
      ],
      ["/groups?$search=%22displayName%3Amarketing%22"],
      ["/groups?$filter=groupTypes%2Fany(c%3Ac%20eq%20'Unified')"],
      ["/groups?$orderby=displayName"],
      ["/groups?$count=true"],
      ["/groups?$expand=members"],
      ["/groups?$skip=1"],
      ["/groups?$top=101"],
      ["/groups?$select=assignedLicenses"],
      ["/groups", { ConsistencyLevel: "eventual" }],
      [`/groups/${GID}?$expand=members`],
      [`/groups/${GID}?$select=accessType`],
      [`/groups/${GID}?$top=1`],
      [`/groups/${GID}/members?$top=5`], // $top never reaches the wire for members/owners
      [`/groups/${GID}/members?$filter=startswith(displayName%2C'a')`],
      [`/groups/${GID}/members?$search=%22displayName%3Aa%22`],
      [`/groups/${GID}/members?$count=true`],
      [`/groups/${GID}/members?$expand=x`],
      [`/groups/${GID}/members?$select=passwordProfile`],
      [`/groups/${GID}/members`, { ConsistencyLevel: "eventual" }],
      [`/groups/${GID}/owners?$top=5`],
      [`/groups/${GID}/owners`, { ConsistencyLevel: "eventual" }],
      [`/groups/${GID}/members/microsoft.graph.user`],
      [`/groups/${GID}/members/$count`],
      [`/groups/${GID}/members/$ref`],
      [`/groups/${GID}/owners/microsoft.graph.user`],
      [`/groups/${GID}/transitiveMembers`],
      [`/groups/${GID}/memberOf`],
      ["/groups/delta"],
      ["/groups/delta/members"],
      [`/groups/${GID}/members?$select=id&$select=mail`],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });
});
