import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { projectDirectoryObject } from "./directory.js";
import { USER_DEFAULT_PROPERTIES, USER_PROPERTIES } from "./identity.js";
import { createHandlers } from "./index.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const USERS = "list-users";
const MY_MANAGER = "get-my-manager";
const MY_REPORTS = "list-my-direct-reports";
const USER_MANAGER = "get-user-manager";
const USER_REPORTS = "list-user-direct-reports";
const GUID = "123e4567-e89b-12d3-a456-426614174000";
const UPN = "alice@contoso.com";
const USER_T = "#microsoft.graph.user";
const CONTACT_T = "#microsoft.graph.orgContact";

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
const parsed = (i = 0) => new URL((calls[i] as Call).url);

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  const fetch: FetchLike = (url, init) => {
    calls.push({
      url,
      method: String(init.method),
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
  it("pins the five v1.0 GET entries", () => {
    const by = (n: string) => catalog.find((e) => e.toolName === n);
    const expected: Array<[string, string]> = [
      [USERS, "/users"],
      [MY_MANAGER, "/me/manager"],
      [MY_REPORTS, "/me/directReports"],
      [USER_MANAGER, "/users/{user-id}/manager"],
      [USER_REPORTS, "/users/{user-id}/directReports"],
    ];
    for (const [name, path] of expected) {
      expect(by(name)?.pathPattern).toBe(path);
      expect(by(name)?.method).toBe("get");
      expect(by(name)?.["apiVersion"]).toBeUndefined();
      // D7/D8: every directory tool keeps User.Read.All (the select contract exceeds ReadBasic).
      expect(by(name)?.["workScopes"]).toEqual(["User.Read.All"]);
    }
  });

  it("list-users metadata matches the typed search contract (D9)", async () => {
    const tip = String(catalog.find((e) => e.toolName === USERS)?.["llmTip"]);
    expect(tip).toContain("ONE property:text clause");
    expect(tip).toContain("the server itself emits");
    expect(tip).toContain("not supported through graph-batch");
    expect(tip).not.toContain("CRITICAL");
    expect(tip).not.toContain('OR "displayName');
    const tool = (await mcp.listTools()).tools.find((t) => t.name === USERS);
    expect(tool?.description).toMatch(/exactly ONE property:text clause/);
    expect(tool?.description).toMatch(/ConsistencyLevel: eventual/);
    expect(tool?.description).toMatch(/not available through graph-batch/);
  });
});

describe(USERS, () => {
  it("sends the default top and select and no advanced-query material", async () => {
    responses.push(() => json({ value: [{ id: "u1", displayName: "A" }] }));
    const { env } = await call(USERS);
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe("/v1.0/users");
    expect(parsed().searchParams.get("$top")).toBe("25");
    expect(parsed().searchParams.get("$select")).toBe(USER_DEFAULT_PROPERTIES.join(","));
    expect(parsed().searchParams.has("$count")).toBe(false);
    expect(parsed().searchParams.has("$search")).toBe(false);
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(calls[0]?.method).toBe("GET");
  });

  it("sets $search, $count and ConsistencyLevel itself, only when a search is given", async () => {
    responses.push(() => json({ value: [] }));
    await call(USERS, { search: "displayName:john smith", top: 10, select: ["id", "mail"] });
    const u = parsed();
    expect(u.searchParams.get("$search")).toBe('"displayName:john smith"');
    expect(u.searchParams.get("$count")).toBe("true");
    expect(u.searchParams.get("$top")).toBe("10");
    expect(u.searchParams.get("$select")).toBe("id,mail");
    expect(calls[0]?.headers["ConsistencyLevel"]).toBe("eventual");
  });

  it("keeps ConsistencyLevel on follow-up pages and binds them to /users", async () => {
    const link = "https://graph.microsoft.com/v1.0/users?$skiptoken=abc";
    responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": link }));
    responses.push(() => json({ value: [{ id: "2" }] }));
    const { env } = await call(USERS, { search: "mail:contoso", top: 5 });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(2);
    expect(calls[1]?.headers["ConsistencyLevel"]).toBe("eventual");
    expect(calls[1]?.headers["Authorization"]).toBe(`Bearer ${TOKEN}`);
  });

  it("rejects unknown options, raw headers, filter/orderby/skip/count and unsafe values", async () => {
    const bad: Array<Record<string, unknown>> = [
      { filter: "startswith(displayName,'a')" },
      { orderby: [{ property: "displayName" }] },
      { skip: 1 },
      { count: true },
      { expand: ["manager"] },
      { headers: { ConsistencyLevel: "eventual" } },
      { ConsistencyLevel: "eventual" },
      { top: 0 },
      { top: 201 },
      { select: ["passwordProfile"] },
      { select: ["accountEnabled"] },
      { select: ["signInActivity"] },
      { select: ["mailboxSettings"] },
      { select: ["assignedLicenses"] },
      { search: "john" }, // no property
      { search: "department:john" }, // property not allowlisted
      { search: 'displayName:jo"hn' },
      { search: "displayName:jo\\hn" },
      { search: "displayName:a OR mail:b" },
      { search: 'displayName:"a"' },
      { search: "displayName:(a)" },
      { search: "displayName:R&D" },
      { search: "displayName:john OR jane" },
      { search: "displayName:john AND x" },
      { search: "displayName:NOT john" },
      { search: "displayName:a\nb" },
      { search: "displayName:" },
      { search: "displayName:***" },
      { search: `displayName:${"x".repeat(200)}` },
      { unknownKey: 1 },
    ];
    for (const args of bad) {
      calls = [];
      const { env, text } = await call(USERS, args);
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code, JSON.stringify(args)).toBe("invalid_input");
      expect(calls, JSON.stringify(args)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("drops over-returned and non-allowlisted properties, and annotations except @odata.type", async () => {
    responses.push(() =>
      json({
        "@odata.context": "x",
        "@odata.count": 9,
        value: [
          {
            id: "u1",
            displayName: "A",
            assignedLicenses: [{ skuId: "s" }],
            passwordProfile: { password: "p" },
            signInActivity: {},
            "@odata.type": USER_T,
            "@odata.id": "evil",
          },
          "not-an-object",
        ],
      }),
    );
    const { env, text } = await call(USERS);
    expect(env.data).toEqual([{ id: "u1", displayName: "A", "@odata.type": USER_T }]);
    for (const s of ["assignedLicenses", "passwordProfile", "signInActivity", "evil"]) {
      expect(text).not.toContain(s);
    }
  });

  it("keeps only the request's effective $select when Graph over-returns allowlisted fields", async () => {
    const over = {
      value: [{ id: "u1", displayName: "A", mail: "a@contoso.com", department: "IT" }],
    };
    responses.push(() => json(over));
    expect((await call(USERS, { select: ["id"] })).env.data).toEqual([{ id: "u1" }]);
    responses.push(() => json(over));
    const def = (await call(USERS)).env.data as Array<Record<string, unknown>>;
    expect(def[0]?.["department"]).toBeUndefined();
    expect(def[0]?.["mail"]).toBe("a@contoso.com");
  });

  it("fails closed on a 200 page without a value array (never an empty success)", async () => {
    for (const body of [{ error: "backend failure" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(USERS)).env.error?.code, JSON.stringify(body)).toBe("graph_error");
      responses.push(() => json(body));
      expect((await call(MY_REPORTS)).env.error?.code, JSON.stringify(body)).toBe("graph_error");
    }
    responses.push(() => json({ value: [] }));
    expect((await call(USERS)).env.data).toEqual([]);
  });

  it("refuses a nextLink that leaves the endpoint or origin (fail closed)", async () => {
    for (const bad of [
      "https://evil.example/v1.0/users?$skiptoken=1",
      "https://graph.microsoft.com/v1.0/users/x/directReports",
      "https://graph.microsoft.com/v1.0/me/manager",
      "https://graph.microsoft.com/beta/users",
    ]) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": bad }));
      const { env } = await call(USERS, { top: 5 });
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("maps 401, 403 and 404 to the shared error conventions", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      const { env, text } = await call(USERS);
      expect(env.error?.code).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe.each([
  [MY_MANAGER, "/v1.0/me/manager", {}],
  [USER_MANAGER, `/v1.0/users/${GUID}/manager`, { "user-id": GUID }],
  [USER_MANAGER, `/v1.0/users/${encodeURIComponent(UPN)}/manager`, { "user-id": UPN }],
] as const)("manager tool %s %s", (name, path, base) => {
  it("sends a bare GET (no $select unless chosen, no headers) and projects a user", async () => {
    responses.push(() =>
      json({
        "@odata.context": "c",
        "@odata.type": USER_T,
        id: "m1",
        displayName: "Boss",
        jobTitle: "VP",
        assignedLicenses: [{ skuId: "s" }],
        passwordProfile: { x: 1 },
      }),
    );
    const { env, text } = await call(name, { ...base });
    expect(calls[0] && new URL(calls[0].url).pathname).toBe(path);
    expect(parsed().search).toBe("");
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(env.data).toEqual({
      "@odata.type": USER_T,
      id: "m1",
      displayName: "Boss",
      jobTitle: "VP",
    });
    expect(text).not.toContain("assignedLicenses");
    expect(text).not.toContain("passwordProfile");
  });

  it("projects to the request's $select only", async () => {
    responses.push(() =>
      json({ "@odata.type": USER_T, id: "m1", displayName: "Boss", mail: "b@c.com", city: "X" }),
    );
    const { env } = await call(name, { ...base, select: ["mail"] });
    expect(parsed().searchParams.get("$select")).toBe("mail");
    expect(env.data).toEqual({ "@odata.type": USER_T, id: "m1", mail: "b@c.com" });
  });

  it("keeps an orgContact to orgContact properties and a foreign type to id and type", async () => {
    responses.push(() =>
      json({
        "@odata.type": CONTACT_T,
        id: "c1",
        displayName: "Contact",
        companyName: "Ext",
        userPrincipalName: "nope@x.com",
        employeeId: "9",
      }),
    );
    expect((await call(name, { ...base })).env.data).toEqual({
      "@odata.type": CONTACT_T,
      id: "c1",
      displayName: "Contact",
      companyName: "Ext",
    });
    responses.push(() =>
      json({
        "@odata.type": "#microsoft.graph.group",
        id: "g1",
        displayName: "G",
        mail: "g@c.com",
      }),
    );
    expect((await call(name, { ...base })).env.data).toEqual({
      "@odata.type": "#microsoft.graph.group",
      id: "g1",
    });
  });

  it("maps 404 (no manager / unknown user), 403 and 401 to the shared error conventions", async () => {
    for (const [status, code] of [
      [404, "graph_error"],
      [403, "capability_gap"],
      [401, "reauth_required"],
    ] as const) {
      responses.push(() =>
        json({ error: { code: "Request_ResourceNotFound", message: "no manager" } }, status),
      );
      const { env, text } = await call(name, { ...base });
      expect(env.ok).toBe(false);
      expect(env.error?.code).toBe(code);
      if (status === 404) expect(env.error?.details?.["status"]).toBe(404);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("fails closed on a non-object response", async () => {
    responses.push(() => json([{ id: "x" }]));
    const { env } = await call(name, { ...base });
    expect(env.error?.code).toBe("graph_error");
  });

  it("rejects unknown options, raw headers and disallowed select values with no request", async () => {
    for (const extra of [
      { expand: ["manager"] },
      { top: 1 },
      { filter: "x" },
      { headers: { ConsistencyLevel: "eventual" } },
      { select: ["passwordProfile"] },
      { select: ["accountEnabled"] },
      { select: [] },
      { nope: 1 },
    ]) {
      calls = [];
      const { env } = await call(name, { ...base, ...extra });
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });
});

describe.each([
  [MY_REPORTS, "/v1.0/me/directReports", {}],
  [USER_REPORTS, `/v1.0/users/${GUID}/directReports`, { "user-id": GUID }],
] as const)("direct-reports tool %s", (name, path, base) => {
  it("sends no $top (undocumented), no headers, and projects per item type", async () => {
    responses.push(() =>
      json({
        "@odata.context": "c",
        value: [
          {
            "@odata.type": USER_T,
            id: "1",
            displayName: "U",
            jobTitle: "T",
            assignedLicenses: [1],
          },
          { "@odata.type": CONTACT_T, id: "2", displayName: "C", employeeId: "9" },
          { id: "3", displayName: "N", userPrincipalName: "x@y.com", givenName: "G" },
          { "@odata.type": "#microsoft.graph.device", id: "4", displayName: "D" },
          "junk",
        ],
      }),
    );
    const { env, text } = await call(name, { ...base });
    expect(parsed().pathname).toBe(path);
    expect(parsed().search).toBe("");
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(env.data).toEqual([
      { "@odata.type": USER_T, id: "1", displayName: "U", jobTitle: "T" },
      { "@odata.type": CONTACT_T, id: "2", displayName: "C" },
      { id: "3", displayName: "N", givenName: "G" },
      { "@odata.type": "#microsoft.graph.device", id: "4" },
    ]);
    expect(text).not.toContain("assignedLicenses");
    expect(text).not.toContain("employeeId");
  });

  it("applies top locally, follows pages, and reports hasMore without exposing the link", async () => {
    const link = (n: number) => `https://graph.microsoft.com/v1.0/${path.slice(6)}?$skiptoken=${n}`;
    responses.push(() => json({ value: [{ id: "1" }, { id: "2" }], "@odata.nextLink": link(1) }));
    responses.push(() => json({ value: [{ id: "3" }, { id: "4" }] }));
    const { env, text } = await call(name, { ...base, top: 3 });
    expect((env.data as unknown[]).map((d) => (d as { id: string }).id)).toEqual(["1", "2", "3"]);
    expect(env.meta?.["hasMore"]).toBe(true);
    expect(text).not.toContain("nextLink");
    expect(parsed(1).searchParams.get("$skiptoken")).toBe("1");
  });

  it("rejects top beyond the limit, $top-style and other options", async () => {
    for (const extra of [{ top: 0 }, { top: 101 }, { filter: "x" }, { search: "x" }, { skip: 1 }]) {
      calls = [];
      const { env } = await call(name, { ...base, ...extra });
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });

  it("refuses a nextLink on another endpoint, origin or (for user-id) principal", async () => {
    const other = "123e4567-e89b-12d3-a456-426614174999";
    const bads = [
      "https://evil.example/v1.0" + path.slice(5) + "?$skiptoken=1",
      "https://graph.microsoft.com/v1.0/users",
      "https://graph.microsoft.com/v1.0/me/manager",
      ...(name === USER_REPORTS
        ? [
            `https://graph.microsoft.com/v1.0/users/${other}/directReports?$skiptoken=1`,
            "https://graph.microsoft.com/v1.0/me/directReports?$skiptoken=1",
            `https://graph.microsoft.com/v1.0/users/${UPN}/directReports?$skiptoken=1`,
          ]
        : []),
    ];
    for (const bad of bads) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": bad }));
      const { env } = await call(name, { ...base, top: 5 });
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("maps 403 and 404 to the shared error conventions", async () => {
    responses.push(() => json({ error: { code: "x", message: "m" } }, 403));
    expect((await call(name, { ...base })).env.error?.code).toBe("capability_gap");
    responses.push(() => json({ error: { code: "x", message: "m" } }, 404));
    expect((await call(name, { ...base })).env.error?.code).toBe("graph_error");
  });
});

describe("{user-id} validation", () => {
  it("rejects anything that is not a GUID or UPN with zero requests", async () => {
    const bad = [
      "..",
      "../me",
      "a/b",
      "a%2Fb",
      "%2e%2e",
      "%252e%252e",
      "me/messages",
      "delta",
      "$count",
      "$ref",
      "microsoft.graph.foo()",
      "a b@contoso.com",
      "a@b",
      "alice@contoso.com/manager",
      "alice@contoso.com?$select=id",
      "alice@contoso.com#x",
      "",
    ];
    for (const tool of [USER_MANAGER, USER_REPORTS]) {
      for (const id of bad) {
        calls = [];
        const { env } = await call(tool, { "user-id": id });
        expect(env.ok, `${tool} ${id}`).toBe(false);
        expect(calls, `${tool} ${id}`).toHaveLength(0);
      }
    }
  });

  it("requires user-id for the user-scoped tools and refuses it on the /me tools", async () => {
    expect((await call(USER_MANAGER, {})).env.ok).toBe(false);
    expect((await call(MY_MANAGER, { "user-id": GUID })).env.ok).toBe(false);
    expect((await call(MY_REPORTS, { "user-id": GUID })).env.ok).toBe(false);
    expect(calls).toHaveLength(0);
  });
});

describe("projectDirectoryObject", () => {
  it("ignores select names outside the type's allowlist and non-objects", () => {
    expect(projectDirectoryObject(null, undefined)).toBeUndefined();
    expect(projectDirectoryObject([], undefined)).toBeUndefined();
    expect(projectDirectoryObject("x", undefined)).toBeUndefined();
    expect(
      projectDirectoryObject(
        { "@odata.type": CONTACT_T, id: "1", mail: "a", city: "c", passwordProfile: 1 },
        "mail,city,passwordProfile",
      ),
    ).toEqual({ "@odata.type": CONTACT_T, id: "1", mail: "a" });
  });

  it("allowlists exactly the SM-9A user properties for user objects", () => {
    const obj: Record<string, unknown> = { "@odata.type": USER_T, id: "1", accountEnabled: true };
    for (const p of USER_PROPERTIES) obj[p] = "v";
    const out = projectDirectoryObject(obj, USER_PROPERTIES.join(",")) as Record<string, unknown>;
    expect(Object.keys(out).sort()).toEqual(
      [...new Set(["@odata.type", "id", ...USER_PROPERTIES])].sort(),
    );
  });
});

describe("graph-batch reachability", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const ok = (body: unknown) => () => json({ responses: [{ id: "1", status: 200, body }] });

  it("registers all five tools with a batch policy", async () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const n of [USERS, MY_MANAGER, MY_REPORTS, USER_MANAGER, USER_REPORTS]) {
      expect(handlers.get(n)?.batchPolicy, n).toBeDefined();
    }
  });

  it("accepts in-policy forms and projects manager and report bodies like the single path", async () => {
    responses.push(ok({ "@odata.type": USER_T, id: "m", displayName: "B", assignedLicenses: [1] }));
    const m = await batch("/me/manager");
    expect(m.env.ok).toBe(true);
    expect(JSON.stringify(m.env.data)).toContain('"displayName":"B"');
    expect(m.text).not.toContain("assignedLicenses");

    responses.push(
      ok({ value: [{ "@odata.type": CONTACT_T, id: "c", displayName: "C", employeeId: "1" }] }),
    );
    const r = await batch(`/users/${GUID}/directReports?$select=displayName`);
    expect(r.env.ok).toBe(true);
    expect(r.text).not.toContain("employeeId");

    responses.push(ok({ value: [{ id: "1", displayName: "A", passwordProfile: 1 }] }));
    const u = await batch("/users?$top=5&$select=id,displayName");
    expect(u.env.ok).toBe(true);
    expect(u.text).not.toContain("passwordProfile");

    // Batch collections are projected from the subrequest's own $select, like the single path.
    responses.push(ok({ value: [{ id: "1", displayName: "A", mail: "a@contoso.com" }] }));
    const narrow = await batch("/users?$select=id");
    expect(narrow.env.ok).toBe(true);
    expect(narrow.text).not.toContain("a@contoso.com");
    expect(narrow.text).not.toContain('"displayName"');
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const bad: Array<[string, Record<string, string>?]> = [
      ["/me/manager?$expand=manager($levels=max)"],
      ["/me/manager?$select=passwordProfile"],
      ["/me/manager", { ConsistencyLevel: "eventual" }],
      ["/me/directReports?$top=5"],
      ["/me/directReports?$filter=x"],
      ["/me/directReports?$count=true"],
      ["/users?$search=%22displayName%3Ajohn%22&$count=true", { ConsistencyLevel: "eventual" }],
      ["/users?$search=%22displayName%3Ajohn%22"],
      ["/users?$count=true", { ConsistencyLevel: "eventual" }],
      ["/users?$filter=startswith(displayName,'a')"],
      ["/users?$top=101"],
      ["/users?$skip=1"],
      ["/users/..%2Fme/manager"],
      ["/users/a%2Fb/directReports"],
      [`/users/${GUID}/manager/$ref`],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });
});
