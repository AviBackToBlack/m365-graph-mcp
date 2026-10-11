import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";
import { PERSON_DEFAULT_PROPERTIES, PERSON_PROPERTIES, projectMembership } from "./people.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const PEOPLE = "list-relevant-people";
const MEMBERS = "list-my-memberships";
const GROUP_T = "#microsoft.graph.group";
const ROLE_T = "#microsoft.graph.directoryRole";
const AU_T = "#microsoft.graph.administrativeUnit";

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
  it("pins the two v1.0 GET entries and their current (non-normative) scopes", () => {
    const by = (n: string) => catalog.find((e) => e.toolName === n);
    expect(by(PEOPLE)?.pathPattern).toBe("/me/people");
    expect(by(PEOPLE)?.["workScopes"]).toEqual(["People.Read"]);
    expect(by(MEMBERS)?.pathPattern).toBe("/me/memberOf");
    // D10: documented least privilege; per-type details are best-effort.
    expect(by(MEMBERS)?.["workScopes"]).toEqual(["User.Read"]);
    for (const n of [PEOPLE, MEMBERS]) {
      expect(by(n)?.method).toBe("get");
      expect(by(n)?.["apiVersion"]).toBeUndefined();
    }
  });

  it("describes the typed contracts, not raw Graph syntax (D11)", async () => {
    const by = (n: string) => catalog.find((e) => e.toolName === n);
    const peopleTip = String(by(PEOPLE)?.["llmTip"]);
    expect(peopleTip).toContain("Use the tool parameters, not raw OData");
    expect(peopleTip).not.toMatch(/Use \$search|Use \$top/);
    const membersTip = String(by(MEMBERS)?.["llmTip"]);
    expect(membersTip).not.toContain("isof");
    expect(membersTip).toContain("direct memberships (not transitive)");
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    expect(desc(PEOPLE)).toMatch(/cannot be combined with search/);
    expect(desc(PEOPLE)).toMatch(/one quoted \$search value/);
    expect(desc(PEOPLE)).toMatch(/nextSkip/);
    expect(desc(MEMBERS)).toMatch(/DIRECT memberships \(not transitive\)/);
    expect(desc(MEMBERS)).toMatch(/best-effort/);
    expect(desc(MEMBERS)).toMatch(/local result limit/);
    expect(desc(MEMBERS)).not.toMatch(/isof|\$filter=/);
  });
});

describe(PEOPLE, () => {
  it("sends the documented default top and select, no headers, no search", async () => {
    responses.push(() => json({ value: [{ id: "p1", displayName: "A" }] }));
    const { env } = await call(PEOPLE);
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe("/v1.0/me/people");
    expect(parsed().searchParams.get("$top")).toBe("10");
    expect(parsed().searchParams.get("$select")).toBe(PERSON_DEFAULT_PROPERTIES.join(","));
    expect([...parsed().searchParams.keys()].sort()).toEqual(["$select", "$top"]);
    expect(calls[0]?.method).toBe("GET");
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(calls[0]?.headers["X-PeopleQuery-QuerySources"]).toBeUndefined();
  });

  it("sends $search as one quoted plain string, top, select and skip", async () => {
    responses.push(() => json({ value: [] }));
    await call(PEOPLE, { search: "tyler lee", top: 5, select: ["id", "phones"] });
    const u = parsed();
    expect(u.searchParams.get("$search")).toBe('"tyler lee"');
    expect(u.searchParams.get("$top")).toBe("5");
    expect(u.searchParams.get("$select")).toBe("id,phones");
    expect(u.searchParams.has("$count")).toBe(false);
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();

    calls = [];
    responses.push(() => json({ value: [] }));
    await call(PEOPLE, { search: "tylerle@example.com" });
    expect(parsed().searchParams.get("$search")).toBe('"tylerle@example.com"');

    calls = [];
    responses.push(() => json({ value: [] }));
    await call(PEOPLE, { top: 3, skip: 10 });
    expect(parsed().searchParams.get("$skip")).toBe("10");
  });

  it("rejects unknown options, raw filter/orderby/headers, unsafe search text and bad bounds", async () => {
    const bad: Array<Record<string, unknown>> = [
      { filter: "personType/class eq 'Person'" },
      { orderby: [{ property: "displayName" }] },
      { count: true },
      { expand: ["x"] },
      { headers: { "X-PeopleQuery-QuerySources": "Mailbox,Directory" } },
      { ConsistencyLevel: "eventual" },
      { $search: "x" },
      { unknownKey: 1 },
      { top: 0 },
      { top: 101 },
      { skip: -1 },
      { skip: 1001 },
      { select: ["personNotes"] },
      { select: ["birthday"] },
      { select: ["websites"] },
      { select: ["postalAddresses"] },
      { select: ["isFavorite"] },
      { select: [] },
      { search: 'ty"ler' },
      { search: "ty\\ler" },
      { search: "topic:budget" },
      { search: "name:x" },
      { search: "(tyler)" },
      { search: "a&b" },
      { search: "a\nb" },
      { search: "tyler OR jane*" },
      { search: "***" },
      { search: "" },
      { search: `x${"y".repeat(100)}` },
      { search: "tyler", skip: 1 }, // documented: $skip is not supported with $search
    ];
    for (const args of bad) {
      calls = [];
      const { env, text } = await call(PEOPLE, args);
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code, JSON.stringify(args)).toBe("invalid_input");
      expect(calls, JSON.stringify(args)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("reports offset paging honestly: a full page means hasMore with nextSkip", async () => {
    const people = (n: number) => ({
      value: Array.from({ length: n }, (_, i) => ({ id: `p${i}`, displayName: `P${i}` })),
    });
    responses.push(() => json(people(10)));
    const full = await call(PEOPLE, { top: 10 });
    expect(full.env.meta?.["hasMore"]).toBe(true);
    expect(full.env.meta?.["nextSkip"]).toBe(10);
    responses.push(() => json(people(5)));
    const next = await call(PEOPLE, { top: 5, skip: 20 });
    expect(next.env.meta?.["nextSkip"]).toBe(25);
    responses.push(() => json(people(3)));
    const short = await call(PEOPLE, { top: 10 });
    expect(short.env.meta?.["hasMore"]).toBe(false);
    expect(short.env.meta?.["nextSkip"]).toBeUndefined();
    // Past the skip maximum (1000) no uncallable nextSkip is offered.
    responses.push(() => json(people(100)));
    const edge = await call(PEOPLE, { top: 100, skip: 950 });
    expect(edge.env.meta?.["hasMore"]).toBe(true);
    expect(edge.env.meta?.["nextSkip"]).toBeUndefined();
    // search cannot be combined with skip, so no nextSkip is offered for a full search page.
    responses.push(() => json(people(10)));
    const searched = await call(PEOPLE, { top: 10, search: "tyler" });
    expect(searched.env.meta?.["hasMore"]).toBe(true);
    expect(searched.env.meta?.["nextSkip"]).toBeUndefined();
  });

  it("accepts names with apostrophes, dots, hyphens and non-ASCII letters", async () => {
    for (const s of ["O'Brien", "j.smith", "Anne-Marie", "Müller", "user_1@contoso.com"]) {
      calls = [];
      responses.push(() => json({ value: [] }));
      expect((await call(PEOPLE, { search: s })).env.ok, s).toBe(true);
      expect(parsed().searchParams.get("$search")).toBe(`"${s}"`);
    }
  });

  it("drops over-returned fields (notes, birthday, ...) and keeps only the request's select", async () => {
    const over = {
      "@odata.context": "ctx",
      value: [
        {
          "@odata.type": "#microsoft.graph.person",
          id: "p1",
          displayName: "A",
          jobTitle: "T",
          department: "D",
          personNotes: "private note",
          birthday: "1990-01-01",
          websites: [{ address: "x" }],
          scoredEmailAddresses: [{ address: "a@contoso.com", relevanceScore: 8 }],
        },
        "junk",
        null,
      ],
    };
    responses.push(() => json(over));
    const def = await call(PEOPLE);
    expect(def.env.data).toEqual([
      {
        "@odata.type": "#microsoft.graph.person",
        id: "p1",
        displayName: "A",
        jobTitle: "T",
        department: "D",
        scoredEmailAddresses: [{ address: "a@contoso.com", relevanceScore: 8 }],
      },
    ]);
    expect(def.text).not.toContain("private note");
    expect(def.text).not.toContain("1990");
    responses.push(() => json(over));
    expect((await call(PEOPLE, { select: ["displayName"] })).env.data).toEqual([
      { "@odata.type": "#microsoft.graph.person", id: "p1", displayName: "A" },
    ]);
  });

  it("allowlists exactly the documented person properties", () => {
    for (const bad of ["personNotes", "birthday", "yomiCompany", "profession", "websites"]) {
      expect((PERSON_PROPERTIES as readonly string[]).includes(bad), bad).toBe(false);
    }
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "x" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(PEOPLE)).env.error?.code, JSON.stringify(body)).toBe("graph_error");
    }
    responses.push(() => json({ value: [] }));
    expect((await call(PEOPLE)).env.data).toEqual([]);
  });

  it("refuses a nextLink that leaves the endpoint or origin, and hides it from callers", async () => {
    for (const bad of [
      "https://evil.example/v1.0/me/people?$skip=1",
      "https://graph.microsoft.com/v1.0/users/x/people",
      "https://graph.microsoft.com/v1.0/me/memberOf",
      "https://graph.microsoft.com/beta/me/people",
    ]) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": bad }));
      const { env } = await call(PEOPLE, { top: 5 });
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
    responses.push(() =>
      json({
        value: [{ id: "1" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/people?$skip=10",
      }),
    );
    responses.push(() => json({ value: [{ id: "2" }] }));
    const ok = await call(PEOPLE, { top: 5 });
    expect(ok.env.ok).toBe(true);
    expect(ok.text).not.toContain("nextLink");
  });

  it("maps 401 and 403 to the shared error conventions", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      const { env, text } = await call(PEOPLE);
      expect(env.error?.code).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe(MEMBERS, () => {
  it("sends a bare GET (no $top, no $select, no headers) and projects per object type", async () => {
    responses.push(() =>
      json({
        "@odata.context": "c",
        value: [
          {
            "@odata.type": GROUP_T,
            id: "g1",
            displayName: "All Users",
            mail: "all@contoso.com",
            mailEnabled: true,
            securityEnabled: true,
            groupTypes: ["Unified"],
            proxyAddresses: ["smtp:x"],
            onPremisesSecurityIdentifier: "S-1",
            membershipRule: "(user.x -eq 1)",
          },
          {
            "@odata.type": ROLE_T,
            id: "r1",
            displayName: "Reader",
            description: "d",
            roleTemplateId: "t1",
            mail: "nope@x.com",
          },
          {
            "@odata.type": AU_T,
            id: "a1",
            displayName: "AU",
            visibility: "HiddenMembership",
            membershipRule: "rule",
            isMemberManagementRestricted: true,
          },
          { "@odata.type": "#microsoft.graph.device", id: "d1", displayName: "Dev" },
          { id: "n1", displayName: "NoType", mail: "x@y.com" },
          "junk",
          7,
        ],
      }),
    );
    const { env, text } = await call(MEMBERS);
    expect(parsed().pathname).toBe("/v1.0/me/memberOf");
    expect(parsed().search).toBe("");
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(env.data).toEqual([
      {
        "@odata.type": GROUP_T,
        id: "g1",
        displayName: "All Users",
        mail: "all@contoso.com",
        mailEnabled: true,
        securityEnabled: true,
        groupTypes: ["Unified"],
      },
      {
        "@odata.type": ROLE_T,
        id: "r1",
        displayName: "Reader",
        description: "d",
        roleTemplateId: "t1",
      },
      { "@odata.type": AU_T, id: "a1", displayName: "AU", visibility: "HiddenMembership" },
      { "@odata.type": "#microsoft.graph.device", id: "d1" },
      { id: "n1" },
    ]);
    for (const leaked of [
      "proxyAddresses",
      "onPremises",
      "membershipRule",
      "nope@x.com",
      "x@y.com",
    ]) {
      expect(text).not.toContain(leaked);
    }
  });

  it("sends the chosen $select and projects each type to it intersected with its allowlist", async () => {
    responses.push(() =>
      json({
        value: [
          { "@odata.type": GROUP_T, id: "g", displayName: "G", mail: "g@x.com", description: "d" },
          {
            "@odata.type": ROLE_T,
            id: "r",
            displayName: "R",
            roleTemplateId: "t",
            mail: "r@x.com",
          },
        ],
      }),
    );
    const { env } = await call(MEMBERS, { select: ["displayName", "mail", "roleTemplateId"] });
    expect(parsed().searchParams.get("$select")).toBe("displayName,mail,roleTemplateId");
    expect(env.data).toEqual([
      { "@odata.type": GROUP_T, id: "g", displayName: "G", mail: "g@x.com" },
      { "@odata.type": ROLE_T, id: "r", displayName: "R", roleTemplateId: "t" },
    ]);
  });

  it("accepts an id-only select and returns id + @odata.type per membership", async () => {
    responses.push(() =>
      json({ value: [{ "@odata.type": GROUP_T, id: "g", displayName: "G", mail: "g@x.com" }] }),
    );
    const { env } = await call(MEMBERS, { select: ["id"] });
    expect(env.ok).toBe(true);
    expect(parsed().searchParams.get("$select")).toBe("id");
    expect(env.data).toEqual([{ "@odata.type": GROUP_T, id: "g" }]);
  });

  it("applies top locally, follows pages, and reports hasMore without exposing the link", async () => {
    const link = (n: number) => `https://graph.microsoft.com/v1.0/me/memberOf?$skiptoken=${n}`;
    responses.push(() => json({ value: [{ id: "1" }, { id: "2" }], "@odata.nextLink": link(1) }));
    responses.push(() => json({ value: [{ id: "3" }, { id: "4" }] }));
    const { env, text } = await call(MEMBERS, { top: 3 });
    expect((env.data as unknown[]).map((d) => (d as { id: string }).id)).toEqual(["1", "2", "3"]);
    expect(env.meta?.["hasMore"]).toBe(true);
    expect(text).not.toContain("nextLink");
    expect(parsed(0).search).toBe("");
    expect(parsed(1).searchParams.get("$skiptoken")).toBe("1");
  });

  it("rejects top beyond the limit, advanced-query options, raw filters/casts and headers", async () => {
    const bad: Array<Record<string, unknown>> = [
      { top: 0 },
      { top: 101 },
      { filter: "isof('microsoft.graph.group')" },
      { search: "displayName:x" },
      { orderby: [{ property: "displayName" }] },
      { skip: 1 },
      { count: true },
      { expand: ["members"] },
      { type: "group" },
      { headers: { ConsistencyLevel: "eventual" } },
      { ConsistencyLevel: "eventual" },
      { select: ["passwordProfile"] },
      { select: ["proxyAddresses"] },
      { select: ["onPremisesSecurityIdentifier"] },
      { select: ["id", "displayName", "assignedLicenses"] },
      { select: [] },
      { unknownKey: 1 },
    ];
    for (const args of bad) {
      calls = [];
      const { env } = await call(MEMBERS, args);
      expect(env.error?.code, JSON.stringify(args)).toBe("invalid_input");
      expect(calls, JSON.stringify(args)).toHaveLength(0);
    }
  });

  it("fails closed on a page without a value array", async () => {
    for (const body of [{ error: "x" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(MEMBERS)).env.error?.code, JSON.stringify(body)).toBe("graph_error");
    }
  });

  it("refuses a nextLink on another endpoint, origin or version", async () => {
    for (const bad of [
      "https://evil.example/v1.0/me/memberOf?$skiptoken=1",
      "https://graph.microsoft.com/v1.0/me/memberOf/microsoft.graph.group?$skiptoken=1",
      "https://graph.microsoft.com/v1.0/me/transitiveMemberOf?$skiptoken=1",
      "https://graph.microsoft.com/v1.0/users/x/memberOf?$skiptoken=1",
      "https://graph.microsoft.com/beta/me/memberOf?$skiptoken=1",
    ]) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": bad }));
      const { env } = await call(MEMBERS, { top: 5 });
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("maps 401 and 403 to the shared error conventions", async () => {
    responses.push(() => json({ error: { code: "x", message: "m" } }, 401));
    expect((await call(MEMBERS)).env.error?.code).toBe("reauth_required");
    responses.push(() => json({ error: { code: "x", message: "m" } }, 403));
    const r = await call(MEMBERS);
    expect(r.env.error?.code).toBe("capability_gap");
    expect(r.text).not.toContain(TOKEN);
  });
});

describe("projectMembership", () => {
  it("ignores select names outside the type's allowlist and non-objects", () => {
    expect(projectMembership(null, undefined)).toBeUndefined();
    expect(projectMembership([], undefined)).toBeUndefined();
    expect(projectMembership("x", undefined)).toBeUndefined();
    expect(
      projectMembership(
        { "@odata.type": ROLE_T, id: "1", mail: "a", roleTemplateId: "t", passwordProfile: 1 },
        "mail,roleTemplateId,passwordProfile",
      ),
    ).toEqual({ "@odata.type": ROLE_T, id: "1", roleTemplateId: "t" });
  });

  it("keeps unknown and untyped objects to id (and the type annotation when present)", () => {
    expect(
      projectMembership(
        { "@odata.type": "#microsoft.graph.servicePrincipal", id: "1", displayName: "x" },
        "displayName",
      ),
    ).toEqual({ "@odata.type": "#microsoft.graph.servicePrincipal", id: "1" });
    expect(projectMembership({ id: "1", displayName: "x" }, "displayName")).toEqual({ id: "1" });
  });
});

describe("graph-batch reachability", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const ok = (body: unknown) => () => json({ responses: [{ id: "1", status: 200, body }] });

  it("registers both tools with a batch policy", async () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const n of [PEOPLE, MEMBERS]) expect(handlers.get(n)?.batchPolicy, n).toBeDefined();
  });

  it("reports people offset paging in batch like the single path", async () => {
    const page = (n: number) => ({
      value: Array.from({ length: n }, (_, i) => ({ id: `p${i}`, displayName: `P${i}` })),
    });
    type Body = { value?: unknown[]; hasMore?: boolean; nextSkip?: number };
    const bodyOf = (r: { env: { data?: unknown } }) =>
      (r.env.data as { responses: Array<{ body?: Body }> }).responses[0]?.body;
    responses.push(ok(page(10)));
    expect(bodyOf(await batch("/me/people?$top=10&$skip=20"))).toMatchObject({
      hasMore: true,
      nextSkip: 30,
    });
    responses.push(ok(page(3)));
    const short = bodyOf(await batch("/me/people?$top=10"));
    expect(short?.hasMore).toBeUndefined();
    expect(short?.nextSkip).toBeUndefined();
    responses.push(ok(page(10)));
    const searched = bodyOf(await batch('/me/people?$top=10&$search="tyler"'));
    expect(searched?.hasMore).toBe(true);
    expect(searched?.nextSkip).toBeUndefined();
    // Past the skip maximum, hasMore stays true but no uncallable nextSkip is offered.
    responses.push(ok(page(100)));
    const edge = bodyOf(await batch("/me/people?$top=100&$skip=950"));
    expect(edge?.hasMore).toBe(true);
    expect(edge?.nextSkip).toBeUndefined();
  });

  it("clips a memberOf batch page to the same default the single call uses", async () => {
    const page = {
      value: Array.from({ length: 30 }, (_, i) => ({ "@odata.type": GROUP_T, id: `g${i}` })),
    };
    responses.push(() => json(page));
    const single = (await call(MEMBERS)).env.data as unknown[];
    responses.push(ok(page));
    const b = await batch("/me/memberOf");
    const body = (
      b.env.data as { responses: Array<{ body?: { value?: unknown[]; hasMore?: boolean } }> }
    ).responses[0]?.body;
    expect(body?.value).toHaveLength(single.length);
    expect(body?.hasMore).toBe(true);
  });

  it("accepts in-policy people forms and projects like the single path", async () => {
    responses.push(
      ok({ value: [{ id: "1", displayName: "A", personNotes: "secret", jobTitle: "T" }] }),
    );
    const a = await batch("/me/people?$top=5&$select=id,displayName");
    expect(a.env.ok).toBe(true);
    expect(a.text).not.toContain("secret");
    expect(a.text).not.toContain("jobTitle");

    // $search round-trips: exactly the form the single path sends, no headers or $count needed.
    responses.push(ok({ value: [{ id: "1", displayName: "Tyler Lee" }] }));
    const s = await batch("/me/people?$search=%22tyler%20lee%22&$top=3");
    expect(s.env.ok).toBe(true);
    expect(calls.at(-1)?.body).toContain("/me/people?$search=%22tyler%20lee%22&$top=3");

    responses.push(ok({ value: [{ id: "1", displayName: "A" }] }));
    expect((await batch("/me/people?$skip=5")).env.ok).toBe(true);
  });

  it("rejects out-of-policy people subrequests with zero network calls", async () => {
    const bad: Array<[string, Record<string, string>?]> = [
      ["/me/people?$search=%22topic%3Abudget%22"],
      ["/me/people?$search=%22a%5C%22b%22"],
      ["/me/people?$search=tyler"], // not the quoted form the single path emits
      ["/me/people?$search=%22tyler%22&$skip=1"],
      ["/me/people?$filter=personType%2Fclass%20eq%20'Person'"],
      ["/me/people?$orderby=displayName"],
      ["/me/people?$count=true", { ConsistencyLevel: "eventual" }],
      ["/me/people", { "X-PeopleQuery-QuerySources": "Mailbox,Directory" }],
      ["/me/people?$select=personNotes"],
      ["/me/people?$top=101"],
      ["/me/people?$top=0"],
      ["/me/people?$top=1&$top=2"],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });

  it("projects memberOf bodies per type and rejects out-of-policy subrequests", async () => {
    responses.push(
      ok({
        value: [
          { "@odata.type": GROUP_T, id: "g", displayName: "G", proxyAddresses: ["x"] },
          { "@odata.type": "#microsoft.graph.device", id: "d", displayName: "D" },
        ],
      }),
    );
    const r = await batch("/me/memberOf");
    expect(r.env.ok).toBe(true);
    expect(r.text).not.toContain("proxyAddresses");
    expect(r.text).not.toContain('"D"');

    const bad: Array<[string, Record<string, string>?]> = [
      ["/me/memberOf?$top=5"], // never on the wire (undocumented)
      ["/me/memberOf?$filter=isof('microsoft.graph.group')", { ConsistencyLevel: "eventual" }],
      ["/me/memberOf?$count=true", { ConsistencyLevel: "eventual" }],
      ["/me/memberOf", { ConsistencyLevel: "eventual" }],
      ["/me/memberOf?$search=%22displayName%3Ax%22"],
      ["/me/memberOf?$select=proxyAddresses"],
      ["/me/memberOf/microsoft.graph.group"],
      ["/me/memberOf/$count"],
      ["/me/transitiveMemberOf"],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });
});
