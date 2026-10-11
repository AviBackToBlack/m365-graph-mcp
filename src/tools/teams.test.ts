import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { projectChatMember } from "./chats.js";
import { createHandlers } from "./index.js";
import { probeUnimplementedGet } from "./unimplemented-get-fixture.js";
import { sanitizeTeamsEntity } from "./teams.js";

/**
 * SM-12 part A: get-team, list-team-channels, get-team-channel, list-team-members. Includes the
 * capability cases (403 -> capability_gap with the no-escalation guidance after exactly one
 * request; 401 -> reauth_required) mirroring chats.test.ts / group-capability.test.ts.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const TEAM_T = "get-team";
const LIST_CH = "list-team-channels";
const GET_CH = "get-team-channel";
const MEMBERS = "list-team-members";
const LIST_SEL =
  "displayName,description,createdDateTime,membershipType,isArchived,isFavoriteByDefault,webUrl";
const GET_SEL = `${LIST_SEL},layoutType`;
const TID = "f17e0000-0000-4000-8000-000000000020";
const OTHER_TID = "f17e0000-0000-4000-8000-000000000009";
const TENANT_ID = "f17e0000-0000-4000-8000-000000000014";
const CID = "19:561fbdbbfca848a484f0a6f00ce9dbbd@thread.tacv2";
const OTHER_CID = "19:982abbfca323a582f0a6d00ae2deca@thread.tacv2";
const AAD_T = "#microsoft.graph.aadUserConversationMember";
const GUEST_T = "#microsoft.graph.anonymousGuestConversationMember";
const enc = encodeURIComponent;

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

const TEAM_DEFAULT =
  "displayName,description,classification,specialization,visibility,webUrl,isArchived,createdDateTime";
const LIST_DEFAULT =
  "displayName,description,createdDateTime,membershipType,isArchived,isFavoriteByDefault,webUrl";
const GET_DEFAULT = `${LIST_DEFAULT},layoutType`;

/** A team as Graph returns it by default, settings objects included (all must be dropped). */
const rawTeam = (over: Record<string, unknown> = {}) => ({
  "@odata.context": "ctx-secret",
  id: TID,
  displayName: "Contoso Team",
  description: "A team",
  classification: "General",
  specialization: "none",
  visibility: "private",
  webUrl: "https://teams.microsoft.com/l/team/x/conversations?groupId=g&tenantId=t",
  isArchived: false,
  createdDateTime: "2020-05-27T19:22:25.692Z",
  tenantId: TENANT_ID,
  internalId: "internal-secret",
  memberSettings: { allowCreateUpdateChannels: true, x: "member-secret" },
  guestSettings: { allowCreateUpdateChannels: false },
  messagingSettings: { allowUserEditMessages: true },
  funSettings: { allowGiphy: true },
  discoverySettings: { showInTeamsSearchAndSuggestions: true },
  classSettings: { x: "class-secret" },
  summary: { ownersCount: 2, membersCount: 3, guestsCount: 4, extra: "summary-secret" },
  ...over,
});
const team = (over: Record<string, unknown> = {}) => ({
  id: TID,
  displayName: "Contoso Team",
  description: "A team",
  classification: "General",
  specialization: "none",
  visibility: "private",
  webUrl: "https://teams.microsoft.com/l/team/x/conversations?groupId=g&tenantId=t",
  isArchived: false,
  createdDateTime: "2020-05-27T19:22:25.692Z",
  ...over,
});

const rawChannel = (over: Record<string, unknown> = {}) => ({
  "@odata.type": "#microsoft.graph.channel",
  id: CID,
  displayName: "General",
  description: "d",
  createdDateTime: "2020-05-27T19:22:25.692Z",
  membershipType: "standard",
  isArchived: false,
  isFavoriteByDefault: true,
  webUrl: "https://teams.microsoft.com/l/channel/x/General?groupId=g",
  layoutType: "post",
  email: "general@x.example",
  tenantId: "tenant-secret",
  migrationMode: "inProgress",
  originalCreatedDateTime: "2020-05-28T10:00:00Z",
  moderationSettings: { x: "moderation-secret" },
  summary: { ownersCount: 1, membersCount: 2, guestsCount: 0, hasMembersFromOtherTenants: false },
  ...over,
});
const channel = (over: Record<string, unknown> = {}) => ({
  id: CID,
  displayName: "General",
  description: "d",
  createdDateTime: "2020-05-27T19:22:25.692Z",
  membershipType: "standard",
  isArchived: false,
  isFavoriteByDefault: true,
  webUrl: "https://teams.microsoft.com/l/channel/x/General?groupId=g",
  ...over,
});

describe("catalog and surface", () => {
  it("pins the four v1.0 GET entries and their documented least-privilege scopes", () => {
    const by = (n: string) => catalog.find((e) => e.toolName === n);
    const expected: Array<[string, string, string]> = [
      [TEAM_T, "/teams/{team-id}", "Team.ReadBasic.All"],
      [LIST_CH, "/teams/{team-id}/channels", "Channel.ReadBasic.All"],
      [GET_CH, "/teams/{team-id}/channels/{channel-id}", "Channel.ReadBasic.All"],
      [MEMBERS, "/teams/{team-id}/members", "TeamMember.Read.All"],
    ];
    for (const [name, path, scope] of expected) {
      expect(by(name)?.pathPattern).toBe(path);
      expect(by(name)?.method).toBe("get");
      expect(by(name)?.["apiVersion"]).toBeUndefined();
      expect(by(name)?.["workScopes"]).toEqual([scope]);
    }
  });

  it("describes the typed contracts, not raw Graph syntax (D25)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    const tip = (n: string) => String(catalog.find((x) => x.toolName === n)?.["llmTip"]);
    for (const n of [TEAM_T, LIST_CH, GET_CH, MEMBERS]) {
      expect(desc(n)).toMatch(/team-id is a GUID|by team-id \(a GUID\)/);
      expect(desc(n)).toMatch(/Retrieval only; work\/school accounts only\. No raw OData options/);
      expect(desc(n) + tip(n)).not.toMatch(/Use \$|\$expand=|\$filter=/);
    }
    expect(desc(TEAM_T)).toMatch(/tenantId and summary are opt-in through select/);
    expect(desc(TEAM_T)).toMatch(/Team settings objects and internalId are not exposed/);
    expect(desc(TEAM_T)).toMatch(/Requires Team\.ReadBasic\.All/);
    expect(desc(LIST_CH)).toMatch(
      /top is a local result limit \(1-200, default 50\): no Graph \$top is sent/,
    );
    expect(desc(LIST_CH)).toMatch(/email is opt-in/);
    expect(desc(LIST_CH)).toMatch(/unknownFutureValue.*must treat as valid/);
    expect(desc(LIST_CH)).toMatch(/layoutType is not useful here/);
    expect(desc(GET_CH)).toMatch(
      /channel-id is an opaque Teams channel\/thread identifier \(not a GUID\)/,
    );
    expect(desc(GET_CH)).toMatch(/email and summary are opt-in/);
    expect(desc(GET_CH)).toMatch(/Moderation and migration metadata and tenantId are not exposed/);
    expect(desc(MEMBERS)).toMatch(/top is a local result limit \(1-250, default 100\)/);
    expect(desc(MEMBERS)).toMatch(/other member types expose only id and @odata\.type/);
    expect(desc(MEMBERS)).toMatch(/Membership ids are opaque/);
    expect(desc(MEMBERS)).toMatch(/Requires TeamMember\.Read\.All/);
  });

  it("exposes only the typed options", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = (n: string) =>
      Object.keys(
        (tools.find((t) => t.name === n)?.inputSchema.properties ?? {}) as Record<string, unknown>,
      ).sort();
    expect(props(TEAM_T)).toEqual(["select", "team-id"]);
    expect(props(LIST_CH)).toEqual(["select", "team-id", "top"]);
    expect(props(GET_CH)).toEqual(["channel-id", "select", "team-id"]);
    expect(props(MEMBERS)).toEqual(["team-id", "top"]);
  });

  it("offers only allowlisted select values (no settings objects, internalId, moderation)", async () => {
    const tools = (await mcp.listTools()).tools;
    const enumOf = (n: string) =>
      (
        (tools.find((t) => t.name === n)?.inputSchema.properties?.["select"] ?? {}) as {
          items?: { enum?: string[] };
        }
      ).items?.enum ?? [];
    expect(enumOf(TEAM_T)).toContain("summary");
    for (const n of [TEAM_T, LIST_CH, GET_CH]) {
      for (const bad of [
        "funSettings",
        "guestSettings",
        "memberSettings",
        "messagingSettings",
        "discoverySettings",
        "classSettings",
        "internalId",
        "moderationSettings",
        "migrationMode",
        "id",
      ]) {
        expect(enumOf(n), `${n} ${bad}`).not.toContain(bad);
      }
    }
    expect(enumOf(LIST_CH)).toContain("email");
    expect(enumOf(LIST_CH)).not.toContain("summary");
    expect(enumOf(LIST_CH)).not.toContain("layoutType");
    expect(enumOf(GET_CH)).toEqual(expect.arrayContaining(["summary", "layoutType", "email"]));
  });

  it("covers exactly the four team tools implemented in this part", () => {
    const names = catalog
      .filter(
        (e) =>
          e.method === "get" &&
          /^\/teams\/\{team-id\}(\/channels(\/\{channel-id\})?|\/members)?$/.test(e.pathPattern),
      )
      .map((e) => e.toolName)
      .sort();
    expect(names).toEqual([GET_CH, LIST_CH, MEMBERS, TEAM_T].sort());
  });

  it("refuses a catalogued GET whose tool is not implemented (synthetic entry), zero requests", async () => {
    // Real catalog GETs are all implemented after SM-14 part A; see unimplemented-get-fixture.ts.
    const probe = await probeUnimplementedGet(catalog);
    expect(probe.catalogued).toBe(true);
    expect(probe.hasHandler).toBe(false);
    expect(probe.batch.ok).toBe(false);
    expect(probe.batch.error?.code).toBe("invalid_input");
    expect(probe.direct.error?.code).toBe("not_implemented");
    expect(probe.requests).toBe(0);
  });
});

describe(TEAM_T, () => {
  it("requests /teams/{id} with the default $select only and projects to the allowlist", async () => {
    responses.push(() => json(rawTeam()));
    const { env, text } = await call(TEAM_T, { "team-id": TID });
    expect(env.ok).toBe(true);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe(`/v1.0/teams/${TID}`);
    expect([...parsed().searchParams.keys()]).toEqual(["$select"]);
    expect(parsed().searchParams.get("$select")).toBe(TEAM_DEFAULT);
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    // tenantId/summary were not selected: dropped even though Graph over-returned them.
    expect(env.data).toEqual(team());
    for (const s of ["secret", TENANT_ID, "Settings", "internalId"]) {
      expect(text).not.toContain(s);
    }
  });

  it("returns tenantId and the reduced summary only when selected", async () => {
    responses.push(() => json(rawTeam()));
    const { env, text } = await call(TEAM_T, {
      "team-id": TID,
      select: ["displayName", "tenantId", "summary"],
    });
    expect(parsed().searchParams.get("$select")).toBe("displayName,tenantId,summary");
    expect(env.data).toEqual({
      id: TID,
      displayName: "Contoso Team",
      tenantId: TENANT_ID,
      summary: { ownersCount: 2, membersCount: 3, guestsCount: 4 },
    });
    expect(text).not.toContain("summary-secret");
  });

  it("rejects options outside the typed contract with zero network calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      { expand: ["channels"] },
      { expand: ["template"] },
      { $expand: "channels" },
      { filter: "x" },
      { top: 1 },
      { search: "x" },
      { skip: 1 },
      { orderby: [{ property: "displayName" }] },
      { select: ["funSettings"] },
      { select: ["memberSettings"] },
      { select: ["internalId"] },
      { select: ["classSettings"] },
      { select: [] },
      { select: "displayName" },
      { headers: { ConsistencyLevel: "eventual" } },
      { unknownKey: 1 },
    ];
    for (const extra of bad) {
      calls = [];
      const { env, text } = await call(TEAM_T, { "team-id": TID, ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("fails closed on a non-object 200 body", async () => {
    for (const body of [[], "x", null]) {
      responses.push(() => json(body));
      expect((await call(TEAM_T, { "team-id": TID })).env.error?.code, JSON.stringify(body)).toBe(
        "graph_error",
      );
    }
  });

  it("maps 401, 403, 404 and 400", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
      [400, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      const { env, text } = await call(TEAM_T, { "team-id": TID });
      expect(env.error?.code, String(status)).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe(LIST_CH, () => {
  it("sends the default $select (no email, the documented perf advice) and no $top", async () => {
    responses.push(() => json({ value: [rawChannel()] }));
    const { env, text } = await call(LIST_CH, { "team-id": TID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/teams/${TID}/channels`);
    expect([...parsed().searchParams.keys()]).toEqual(["$select"]);
    expect(parsed().searchParams.get("$select")).toBe(LIST_DEFAULT);
    expect(parsed().searchParams.get("$select")).not.toContain("email");
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    // email/layoutType/tenantId/migration/moderation/summary over-returned: all dropped.
    expect(env.data).toEqual([channel()]);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
    for (const s of ["secret", "general@x.example", "layoutType", "migrationMode", "summary"]) {
      expect(text).not.toContain(s);
    }
  });

  it("returns email only when explicitly selected", async () => {
    responses.push(() => json({ value: [rawChannel()] }));
    const { env } = await call(LIST_CH, {
      "team-id": TID,
      select: ["displayName", "email", "membershipType"],
    });
    expect(parsed().searchParams.get("$select")).toBe("displayName,email,membershipType");
    expect(env.data).toEqual([
      {
        id: CID,
        displayName: "General",
        membershipType: "standard",
        email: "general@x.example",
      },
    ]);
  });

  it("passes membershipType through verbatim (shared channels may report unknownFutureValue)", async () => {
    responses.push(() =>
      json({
        value: [
          rawChannel({ membershipType: "private" }),
          rawChannel({ id: OTHER_CID, membershipType: "unknownFutureValue" }),
          rawChannel({ id: "19:z@thread.tacv2", membershipType: { evil: 1 } }),
        ],
      }),
    );
    const { env } = await call(LIST_CH, { "team-id": TID });
    const items = env.data as Array<Record<string, unknown>>;
    expect(items.map((i) => i["membershipType"])).toEqual([
      "private",
      "unknownFutureValue",
      undefined,
    ]);
  });

  it("treats top as a local limit, sends no $top and follows validated pages", async () => {
    const link = `https://graph.microsoft.com/v1.0/teams/${TID}/channels?$select=displayName&$skiptoken=abc`;
    responses.push(() => json({ value: [rawChannel()], "@odata.nextLink": link }));
    responses.push(() => json({ value: [rawChannel({ id: OTHER_CID }), rawChannel({ id: "x" })] }));
    const { env } = await call(LIST_CH, { "team-id": TID, top: 2 });
    expect(calls).toHaveLength(2);
    expect(parsed(0).searchParams.has("$top")).toBe(false);
    expect(parsed(1).pathname).toBe(`/v1.0/teams/${TID}/channels`);
    expect((env.data as unknown[]).length).toBe(2);
    expect(env.meta).toMatchObject({ count: 2, hasMore: true });
  });

  it("accepts the OData key-syntax form Graph uses in links for the same team", async () => {
    const link = `https://graph.microsoft.com/v1.0/teams('${TID}')/channels?$skiptoken=abc`;
    responses.push(() => json({ value: [rawChannel()], "@odata.nextLink": link }));
    responses.push(() => json({ value: [rawChannel({ id: OTHER_CID })] }));
    const { env } = await call(LIST_CH, { "team-id": TID, top: 5 });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(2);
  });

  it("rejects options outside the typed contract with zero network calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      { top: 0 },
      { top: 201 },
      { top: 1.5 },
      { filter: "membershipType eq 'private'" },
      { $filter: "membershipType eq 'private'" },
      { membershipType: "private" },
      { expand: ["tabs"] },
      { search: "x" },
      { skip: 1 },
      { orderby: [{ property: "displayName" }] },
      { count: true },
      { select: ["summary"] },
      { select: ["layoutType"] },
      { select: ["moderationSettings"] },
      { select: ["tenantId"] },
      { select: [] },
      { headers: { Prefer: "include-unknown-enum-members" } },
      { unknownKey: 1 },
    ];
    for (const extra of bad) {
      calls = [];
      const { env, text } = await call(LIST_CH, { "team-id": TID, ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("drops non-object items and non-scalar values", async () => {
    responses.push(() =>
      json({
        value: [rawChannel({ displayName: { evil: 1 }, description: ["x"] }), "x", ["a"], null],
      }),
    );
    const { env } = await call(LIST_CH, { "team-id": TID });
    expect((env.data as Array<Record<string, unknown>>).length).toBe(1);
    const [first] = env.data as Array<Record<string, unknown>>;
    expect(first).not.toHaveProperty("displayName");
    expect(first).not.toHaveProperty("description");
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "backend failure" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(LIST_CH, { "team-id": TID })).env.error?.code, JSON.stringify(body)).toBe(
        "graph_error",
      );
    }
  });

  it("refuses a nextLink for another team, collection, host or version (fail closed)", async () => {
    for (const bad of [
      `https://graph.microsoft.com/v1.0/teams/${OTHER_TID}/channels?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/teams/${TID}/allChannels?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/teams/${TID}/incomingChannels`,
      `https://graph.microsoft.com/v1.0/teams/${TID}/members`,
      `https://graph.microsoft.com/v1.0/teams/${TID}`,
      `https://graph.microsoft.com/v1.0/teams/${TID}/channels/${enc(CID)}`,
      `https://graph.microsoft.com/v1.0/teams/${TID}/channels/${enc(CID)}/messages`,
      `https://graph.microsoft.com/v1.0/groups/${TID}/team/channels`,
      `https://graph.microsoft.com/v1.0/me/joinedTeams`,
      `https://evil.example/v1.0/teams/${TID}/channels`,
      `https://graph.microsoft.com.evil.example/v1.0/teams/${TID}/channels`,
      `https://graph.microsoft.com:444/v1.0/teams/${TID}/channels`,
      `http://graph.microsoft.com/v1.0/teams/${TID}/channels`,
      `https://graph.microsoft.com/beta/teams/${TID}/channels`,
      `https://graph.microsoft.com/v1.0/teams/${TID}%2f..%2f${OTHER_TID}/channels`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [rawChannel()], "@odata.nextLink": bad }));
      const { env } = await call(LIST_CH, { "team-id": TID, top: 5 });
      expect(env.ok, bad).toBe(false);
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("maps 400/404 to graph_error", async () => {
    for (const status of [400, 404]) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      expect((await call(LIST_CH, { "team-id": TID })).env.error?.code, String(status)).toBe(
        "graph_error",
      );
    }
  });
});

describe(GET_CH, () => {
  it("requests the channel with the default $select (incl. layoutType, no email/summary)", async () => {
    responses.push(() => json(rawChannel()));
    const { env, text } = await call(GET_CH, { "team-id": TID, "channel-id": CID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/teams/${TID}/channels/${enc(CID)}`);
    expect([...parsed().searchParams.keys()]).toEqual(["$select"]);
    expect(parsed().searchParams.get("$select")).toBe(GET_DEFAULT);
    expect(env.data).toEqual(channel({ layoutType: "post" }));
    for (const s of ["secret", "general@x.example", "migrationMode", "summary"]) {
      expect(text).not.toContain(s);
    }
  });

  it("returns email and the reduced summary only when selected", async () => {
    responses.push(() =>
      json({
        "@odata.context": "ctx",
        summary: {
          ownersCount: 2,
          membersCount: 3,
          guestsCount: 1,
          hasMembersFromOtherTenants: false,
          extra: "summary-secret",
        },
        email: "e@x.example",
        displayName: "General",
      }),
    );
    const { env, text } = await call(GET_CH, {
      "team-id": TID,
      "channel-id": CID,
      select: ["summary", "email"],
    });
    expect(parsed().searchParams.get("$select")).toBe("summary,email");
    expect(env.data).toEqual({
      email: "e@x.example",
      summary: {
        ownersCount: 2,
        membersCount: 3,
        guestsCount: 1,
        hasMembersFromOtherTenants: false,
      },
    });
    expect(text).not.toContain("summary-secret");
  });

  it("accepts the documented channel id shapes", async () => {
    for (const id of [
      CID,
      "19:abc123def@thread.skype",
      "19:meeting_MjdhNjM4YzUtYzExZi00OTFk@thread.v2",
    ]) {
      calls = [];
      responses.push(() => json(rawChannel({ id })));
      const { env } = await call(GET_CH, { "team-id": TID, "channel-id": id });
      expect(env.ok, id).toBe(true);
      expect(calls, id).toHaveLength(1);
    }
  });

  it("rejects options outside the typed contract with zero network calls", async () => {
    for (const extra of [
      { expand: ["members"] },
      { expand: ["tabs"] },
      { filter: "x" },
      { top: 1 },
      { select: ["moderationSettings"] },
      { select: ["migrationMode"] },
      { select: ["tenantId"] },
      { select: [] },
      { unknownKey: 1 },
    ]) {
      calls = [];
      const { env } = await call(GET_CH, { "team-id": TID, "channel-id": CID, ...extra });
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });

  it("fails closed on a non-object 200 body and maps errors", async () => {
    for (const body of [[], "x", null]) {
      responses.push(() => json(body));
      expect(
        (await call(GET_CH, { "team-id": TID, "channel-id": CID })).env.error?.code,
        JSON.stringify(body),
      ).toBe("graph_error");
    }
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
      [400, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      const { env, text } = await call(GET_CH, { "team-id": TID, "channel-id": CID });
      expect(env.error?.code, String(status)).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe(MEMBERS, () => {
  const member = (over: Record<string, unknown> = {}) => ({
    "@odata.type": AAD_T,
    id: "bWVtYmVy",
    roles: ["owner"],
    displayName: "John Doe",
    userId: "f17e0000-0000-4000-8000-000000000011",
    email: "john@contoso.com",
    tenantId: TENANT_ID,
    ...over,
  });

  it("requests /teams/{id}/members with NO query (no $top, no $select) and projects the users", async () => {
    responses.push(() => json({ "@odata.context": "ctx", "@odata.count": 1, value: [member()] }));
    const { env, text } = await call(MEMBERS, { "team-id": TID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/teams/${TID}/members`);
    expect([...parsed().searchParams.keys()]).toEqual([]);
    expect(env.data).toEqual([member()]);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
    expect(text).not.toContain("ctx");
  });

  it("projects per @odata.type exactly like chat members", async () => {
    const items = [
      member(),
      member({ id: "guest", roles: ["guest"], extra: "x" }),
      { "@odata.type": GUEST_T, id: "g", displayName: "guest-secret", email: "g@x.example" },
      { id: "typeless", displayName: "typeless-secret", email: "t@x.example" },
      "not-an-object",
    ];
    responses.push(() => json({ value: items }));
    const { env, text } = await call(MEMBERS, { "team-id": TID });
    expect(env.data).toEqual(items.slice(0, 4).map((i) => projectChatMember(i)));
    expect(env.data).toEqual([
      member(),
      member({ id: "guest", roles: ["guest"] }),
      { "@odata.type": GUEST_T, id: "g" },
      { id: "typeless" },
    ]);
    expect(text).not.toContain("secret");
  });

  it("treats top as a local limit and follows pages without sending $top", async () => {
    const link = `https://graph.microsoft.com/v1.0/teams/${TID}/members?$skiptoken=abc`;
    responses.push(() => json({ value: [member({ id: "1" })], "@odata.nextLink": link }));
    responses.push(() => json({ value: [member({ id: "2" }), member({ id: "3" })] }));
    const { env } = await call(MEMBERS, { "team-id": TID, top: 2 });
    expect(calls).toHaveLength(2);
    expect(parsed(0).searchParams.has("$top")).toBe(false);
    expect((env.data as unknown[]).length).toBe(2);
    expect(env.meta).toMatchObject({ count: 2, hasMore: true });
  });

  it("rejects options and unsafe values with zero network calls", async () => {
    for (const extra of [
      { top: 0 },
      { top: 251 },
      { select: ["email"] },
      { filter: "(microsoft.graph.aadUserConversationMember/userId eq 'x')" },
      { expand: ["x"] },
      { search: "x" },
      { skip: 1 },
      { unknownKey: 1 },
    ]) {
      calls = [];
      const { env } = await call(MEMBERS, { "team-id": TID, ...extra });
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });

  it("refuses a nextLink for another team, collection, host or version (fail closed)", async () => {
    for (const bad of [
      `https://graph.microsoft.com/v1.0/teams/${OTHER_TID}/members?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/teams/${TID}/channels`,
      `https://graph.microsoft.com/v1.0/teams/${TID}/channels/${enc(CID)}/members`,
      `https://graph.microsoft.com/v1.0/teams/${TID}`,
      `https://graph.microsoft.com/v1.0/groups/${TID}/members`,
      `https://evil.example/v1.0/teams/${TID}/members`,
      `https://graph.microsoft.com/beta/teams/${TID}/members`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [member()], "@odata.nextLink": bad }));
      const { env } = await call(MEMBERS, { "team-id": TID, top: 5 });
      expect(env.ok, bad).toBe(false);
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("fails closed on a 200 page without a value array and maps errors", async () => {
    responses.push(() => json({ error: "backend failure" }));
    expect((await call(MEMBERS, { "team-id": TID })).env.error?.code).toBe("graph_error");
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
      [400, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      expect((await call(MEMBERS, { "team-id": TID })).env.error?.code, String(status)).toBe(code);
    }
  });
});

describe("sanitizeTeamsEntity", () => {
  it("keeps scalars, reduces summary and yields undefined for non-objects", () => {
    expect(
      sanitizeTeamsEntity(
        {
          id: "1",
          displayName: { a: 1 },
          isArchived: false,
          summary: { ownersCount: 1, nested: { a: 1 } },
          funSettings: { allowGiphy: true },
          "@odata.type": "x",
        },
        "team",
      ),
    ).toEqual({ id: "1", isArchived: false, summary: { ownersCount: 1 } });
    expect(sanitizeTeamsEntity({ id: "1", summary: null }, "team")).toEqual({
      id: "1",
      summary: null,
    });
    expect(sanitizeTeamsEntity({ id: "1", summary: "x" }, "team")).toEqual({ id: "1" });
    for (const bad of [null, "x", 1, [], undefined])
      expect(sanitizeTeamsEntity(bad, "team")).toBeUndefined();
  });

  it("keeps hasMembersFromOtherTenants only on a channel summary, never on a team", () => {
    const summary = {
      ownersCount: 1,
      membersCount: 2,
      guestsCount: 0,
      hasMembersFromOtherTenants: true,
    };
    expect(sanitizeTeamsEntity({ id: "t", summary }, "team")).toEqual({
      id: "t",
      summary: { ownersCount: 1, membersCount: 2, guestsCount: 0 },
    });
    expect(sanitizeTeamsEntity({ id: "c", summary }, "channel")).toEqual({ id: "c", summary });
  });
});

describe("team-id and channel-id injection (zero network calls)", () => {
  const badTeams = [
    "delta",
    "getAllMessages",
    "$ref",
    "$count",
    "..",
    ".",
    "%2e%2e",
    `${TID}/channels`,
    `${TID}%2Fchannels`,
    `${TID}%2f..%2f${OTHER_TID}`,
    `${TID}?$expand=channels`,
    `${TID}#frag`,
    `${TID}$ref`,
    ` ${TID}`,
    `${TID} `,
    `${TID}\n`,
    `${TID}\\x`,
    "123e4567-e89b-12d3-a456",
    "user@contoso.com",
    CID,
    "microsoft.graph.team",
    "",
  ];
  const badChannels = [
    "19:abc@thread.tacv2/messages",
    "19%3Aabc%40thread.tacv2%2Fmessages",
    "19:abc@thread.tacv2%2F..%2F..",
    "..",
    "%2e%2e",
    "getAllMessages",
    "delta",
    "$ref",
    "19:abc@thread.tacv2?$select=id",
    "19:abc@thread.tacv2#frag",
    "19:abc@thread.tacv2/tabs",
    " 19:abc@thread.tacv2",
    "19:abc @thread.tacv2",
    `19:${"a".repeat(600)}@thread.tacv2`,
    "19:abc@evil.example",
    "19:@thread.tacv2",
    TID,
    "48:notes",
    "",
  ];
  const cases: Array<[string, Record<string, unknown>]> = [
    [TEAM_T, {}],
    [LIST_CH, {}],
    [MEMBERS, {}],
    [GET_CH, { "channel-id": CID }],
  ];
  it.each(cases)("%s rejects malformed or route-changing team ids", async (name, extra) => {
    for (const id of badTeams) {
      calls = [];
      const { env, text } = await call(name, { "team-id": id, ...extra });
      expect(env.ok, `${name} ${JSON.stringify(id)}`).toBe(false);
      expect(env.error?.code, JSON.stringify(id)).toBe("invalid_input");
      expect(calls, JSON.stringify(id)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("get-team-channel rejects malformed or route-changing channel ids", async () => {
    for (const id of badChannels) {
      calls = [];
      const { env } = await call(GET_CH, { "team-id": TID, "channel-id": id });
      expect(env.ok, JSON.stringify(id)).toBe(false);
      expect(env.error?.code, JSON.stringify(id)).toBe("invalid_input");
      expect(calls, JSON.stringify(id)).toHaveLength(0);
    }
  });

  it("rejects missing or non-string ids", async () => {
    for (const [name, base] of cases) {
      for (const bad of [undefined, 5, null, [TID]]) {
        calls = [];
        const { env } = await call(name, { ...base, "team-id": bad });
        expect(env.ok, `${name} ${JSON.stringify(bad)}`).toBe(false);
        expect(calls).toHaveLength(0);
      }
    }
    for (const bad of [undefined, 5, null, [CID]]) {
      calls = [];
      const { env } = await call(GET_CH, { "team-id": TID, "channel-id": bad });
      expect(env.ok).toBe(false);
      expect(calls).toHaveLength(0);
    }
  });
});

describe("graph-batch reachability", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const ok = (body: unknown) => () => json({ responses: [{ id: "1", status: 200, body }] });
  const bodyOf = (r: { env: Env }) =>
    (
      r.env.data as {
        responses: Array<{
          body?: { value?: unknown[]; hasMore?: boolean } & Record<string, unknown>;
        }>;
      }
    ).responses[0]?.body;
  const member = (over: Record<string, unknown> = {}) => ({
    "@odata.type": AAD_T,
    id: "bWVtYmVy",
    roles: [],
    displayName: "John Doe",
    userId: "u",
    email: "j@x.example",
    ...over,
  });

  it("registers all four tools with a batch policy", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const n of [TEAM_T, LIST_CH, GET_CH, MEMBERS]) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
    }
  });

  it("accepts the in-policy forms and projects like the single path", async () => {
    // get-team: unselected (default) and selected forms
    responses.push(ok(rawTeam()));
    const t = await batch(`/teams/${TID}`);
    expect(t.env.ok).toBe(true);
    expect(calls[0]?.url).toContain("/v1.0/$batch");
    expect(bodyOf(t)).toEqual(team());
    expect(t.text).not.toContain("secret");

    responses.push(ok(rawTeam()));
    const ts = await batch(`/teams/${TID}?$select=displayName,summary`);
    expect(bodyOf(ts)).toEqual({
      id: TID,
      displayName: "Contoso Team",
      summary: { ownersCount: 2, membersCount: 3, guestsCount: 4 },
    });
    expect(ts.text).not.toContain("summary-secret");

    // list-team-channels: default projection drops email; explicit select keeps it
    responses.push(ok({ value: [rawChannel()] }));
    const l = await batch(`/teams/${TID}/channels?$select=${LIST_SEL}`);
    expect(bodyOf(l)?.value).toEqual([channel()]);
    expect(l.text).not.toContain("general@x.example");
    expect(l.text).not.toContain("secret");
    responses.push(ok({ value: [rawChannel()] }));
    const le = await batch(`/teams/${TID}/channels?$select=displayName,email`);
    expect(bodyOf(le)?.value).toEqual([
      { id: CID, displayName: "General", email: "general@x.example" },
    ]);

    // get-team-channel: raw and encoded id, defaults and summary
    for (const id of [enc(CID), CID]) {
      responses.push(ok(rawChannel()));
      const g = await batch(`/teams/${TID}/channels/${id}?$select=${GET_SEL}`);
      expect(g.env.ok, id).toBe(true);
      expect(bodyOf(g)).toEqual(channel({ layoutType: "post" }));
      expect(g.text).not.toContain("secret");
    }
    responses.push(ok(rawChannel()));
    const gs = await batch(`/teams/${TID}/channels/${enc(CID)}?$select=summary`);
    expect(bodyOf(gs)).toEqual({
      id: CID,
      summary: {
        ownersCount: 1,
        membersCount: 2,
        guestsCount: 0,
        hasMembersFromOtherTenants: false,
      },
    });

    // members: per-type projection
    responses.push(
      ok({
        value: [
          member(),
          { "@odata.type": GUEST_T, id: "g", displayName: "guest-secret" },
          { id: "typeless", email: "typeless-secret" },
        ],
      }),
    );
    const m = await batch(`/teams/${TID}/members`);
    expect(bodyOf(m)?.value).toEqual([
      member(),
      { "@odata.type": GUEST_T, id: "g" },
      { id: "typeless" },
    ]);
    expect(m.text).not.toContain("secret");
  });

  it("clips batch pages to the same limits as the single path", async () => {
    const page = (n: number) => ({
      value: Array.from({ length: n }, (_, i) => rawChannel({ id: `19:c${i}@thread.tacv2` })),
    });
    responses.push(() => json(page(60)));
    const single = (await call(LIST_CH, { "team-id": TID })).env.data as unknown[];
    responses.push(ok(page(60)));
    const b = bodyOf(await batch(`/teams/${TID}/channels?$select=${LIST_SEL}`));
    expect(single).toHaveLength(50);
    expect(b?.value).toHaveLength(50);
    expect(b?.hasMore).toBe(true);
    responses.push(
      ok({ value: Array.from({ length: 120 }, (_, i) => ({ "@odata.type": AAD_T, id: `m${i}` })) }),
    );
    const mm = bodyOf(await batch(`/teams/${TID}/members`));
    expect(mm?.value).toHaveLength(100);
    expect(mm?.hasMore).toBe(true);
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const cid = enc(CID);
    const bad: Array<[string, Record<string, string>?]> = [
      [`/teams/${TID}?$expand=channels`],
      [`/teams/${TID}?$expand=members`],
      [`/teams/${TID}?$select=funSettings`],
      [`/teams/${TID}?$select=internalId`],
      [`/teams/${TID}?$select=`],
      [`/teams/${TID}?$select=displayName&$select=webUrl`],
      [`/teams/${TID}?$top=1`],
      [`/teams/${TID}?$filter=x`],
      [`/teams/${TID}`, { ConsistencyLevel: "eventual" }],
      [`/teams/${TID}`, { Prefer: "include-unknown-enum-members" }],
      // Batch must carry $select for channels, so Graph never populates the expensive email/summary.
      [`/teams/${TID}/channels`],
      [`/teams/${TID}/channels/${cid}`],
      [`/teams/${TID}/channels?$top=5`],
      [`/teams/${TID}/channels?$filter=membershipType%20eq%20'private'`],
      [`/teams/${TID}/channels?$select=summary`],
      [`/teams/${TID}/channels?$select=layoutType`],
      [`/teams/${TID}/channels?$select=moderationSettings`],
      [`/teams/${TID}/channels?$expand=tabs`],
      [`/teams/${TID}/channels?$orderby=displayName`],
      [`/teams/${TID}/channels?$count=true`],
      [`/teams/${TID}/channels`, { Prefer: "include-unknown-enum-members" }],
      [`/teams/${TID}/channels/${cid}?$expand=members`],
      [`/teams/${TID}/channels/${cid}?$select=migrationMode`],
      [`/teams/${TID}/channels/${cid}?$filter=x`],
      [`/teams/${TID}/channels/${cid}?$top=1`],
      [`/teams/${TID}/channels/${cid}`, { ConsistencyLevel: "eventual" }],
      [`/teams/${TID}/members?$top=5`],
      [`/teams/${TID}/members?$select=email`],
      [`/teams/${TID}/members?$filter=x`],
      [`/teams/${TID}/members`, { ConsistencyLevel: "eventual" }],
      [`/teams/${TID}/members/microsoft.graph.aadUserConversationMember`],
      [`/teams/${TID}/members/$count`],
      [`/teams/${TID}/allChannels`],
      [`/teams/${TID}/incomingChannels`],
      [`/teams/${TID}/primaryChannel`],
      [`/teams/${TID}/installedApps`],
      [`/teams/${TID}/photo`],
      [`/teams/${TID}/schedule`],
      [`/teams/${TID}/operations`],
      [`/teams/${TID}/channels/${cid}/members`],
      [`/teams/${TID}/channels/${cid}/allMembers`],
      [`/teams/${TID}/channels/${cid}/sharedWithTeams`],
      [`/teams/${TID}/channels/${enc(`${CID}/messages`)}`],
      [`/teams/${TID}%2f..%2f${OTHER_TID}/channels`],
      [`/teams/${enc(`${TID}/channels`)}`],
      [`/groups/${TID}/team`],
      [`/groups/${TID}/team/channels`],
      ["/teams"],
      ["/teams/delta"],
      ["/teams/getAllMessages"],
      [`/teams/${TID}/channels/getAllMessages`],
      [`/teams/${TID}/channels/delta`],
      ["/me/joinedTeams"],
      [`/beta/teams/${TID}`],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });
});

type Case = { name: string; args: Record<string, unknown>; paged: boolean; next: string };
const SM12A: Case[] = [
  { name: TEAM_T, args: { "team-id": TID }, paged: false, next: "" },
  { name: LIST_CH, args: { "team-id": TID }, paged: true, next: `/teams/${TID}/channels` },
  { name: GET_CH, args: { "team-id": TID, "channel-id": CID }, paged: false, next: "" },
  { name: MEMBERS, args: { "team-id": TID }, paged: true, next: `/teams/${TID}/members` },
];

describe.each(SM12A)("$name: unavailable permission", (c) => {
  it("403 is a capability_gap with the no-escalation guidance after exactly one request", async () => {
    responses.push(() =>
      json(
        {
          error: {
            code: "Forbidden",
            message: `Missing scope permissions on the request. Bearer ${TOKEN}`,
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
    expect(calls).toHaveLength(1);
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
    expect(env.error?.code).toBe("reauth_required");
    expect(env.error?.message).not.toMatch(/scope/i);
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
  });
});

describe.each(SM12A.filter((c) => c.paged))("$name: permission lost mid-listing", (c) => {
  it("a 403 or 401 on a followed page fails the call instead of returning the first page", async () => {
    for (const [status, code] of [
      [403, "capability_gap"],
      [401, "reauth_required"],
    ] as const) {
      calls = [];
      responses.push(() =>
        json({
          value: [{ "@odata.type": AAD_T, id: "1" }],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0${c.next}?$skiptoken=abc`,
        }),
      );
      responses.push(() => json({ error: { code: "x", message: TOKEN } }, status));
      const { env, text } = await call(c.name, { ...c.args, top: 50 });
      expect(env.ok, `${c.name} ${status}`).toBe(false);
      expect(env.data).toBeUndefined();
      expect(env.error?.code).toBe(code);
      expect(calls).toHaveLength(2);
      expect(text).not.toContain(TOKEN);
    }
  });
});
