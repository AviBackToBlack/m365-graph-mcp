import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { projectChat, projectChatMember } from "./chats.js";
import { createHandlers } from "./index.js";

/**
 * SM-11 part A: list-chats, get-chat, list-chat-members. Includes the capability cases
 * (403 -> capability_gap with the no-escalation guidance after exactly one request; 401 ->
 * reauth_required) mirroring group-capability.test.ts.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const LIST = "list-chats";
const GET = "get-chat";
const MEMBERS = "list-chat-members";
const CID = "19:b8577894a63548969c5c92bb9c80c5e1@thread.v2";
const OTHER_CID = "19:ebe3857aa388434bab0cad9d2e09f4de@thread.v2";
const ONE_ON_ONE =
  "19:f17e0000-0000-4000-8000-000000000017_f17e0000-0000-4000-8000-000000000011@unq.gbl.spaces";
const AAD_T = "#microsoft.graph.aadUserConversationMember";
const GUEST_T = "#microsoft.graph.anonymousGuestConversationMember";
const PREVIEW = "lastMessagePreview";
const ORDER = "lastMessagePreview/createdDateTime";

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

const chat = (over: Record<string, unknown> = {}) => ({
  id: CID,
  topic: "Demo",
  createdDateTime: "2021-04-06T19:49:52.431Z",
  lastUpdatedDateTime: "2021-04-06T19:54:04.306Z",
  chatType: "group",
  webUrl: "https://teams.microsoft.com/l/chat/x/0?tenantId=t",
  tenantId: "f17e0000-0000-4000-8000-000000000014",
  isHiddenForAllMembers: false,
  viewpoint: { isHidden: false, lastMessageReadDateTime: "2021-05-06T23:55:07.191Z" },
  ...over,
});

describe("catalog and surface", () => {
  it("pins the three v1.0 GET entries and the Chat.ReadBasic scope (D18)", () => {
    const by = (n: string) => catalog.find((e) => e.toolName === n);
    const expected: Array<[string, string, string]> = [
      [LIST, "/me/chats", "Chat.ReadBasic"],
      [GET, "/chats/{chat-id}", "Chat.ReadBasic"],
      // D18: the documented least privilege, shared by all three part A tools.
      [MEMBERS, "/chats/{chat-id}/members", "Chat.ReadBasic"],
    ];
    for (const [name, path, scope] of expected) {
      expect(by(name)?.pathPattern).toBe(path);
      expect(by(name)?.method).toBe("get");
      expect(by(name)?.["apiVersion"]).toBeUndefined();
      expect(by(name)?.["workScopes"]).toEqual([scope]);
    }
  });

  it("describes the typed contracts, not raw Graph syntax (D19)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    const tip = (n: string) => String(catalog.find((x) => x.toolName === n)?.["llmTip"]);
    for (const n of [LIST, GET, MEMBERS]) {
      expect(desc(n)).toMatch(/Requires Chat\.ReadBasic/);
      expect(desc(n) + tip(n)).not.toMatch(/get-online-meeting|ChatMember\.Read|Use \$/);
    }
    expect(desc(LIST)).toMatch(/at most 50 chats are requested from Graph per page/);
    expect(desc(LIST)).toMatch(/expand supports only lastMessagePreview/);
    expect(desc(LIST)).toMatch(/no message body, event detail or attachments/);
    expect(desc(LIST)).toMatch(/lastMessagePreview\/createdDateTime in descending order/);
    expect(desc(LIST)).toMatch(/use list-chat-members/);
    expect(desc(GET)).toMatch(/Takes no query options/);
    expect(desc(GET)).toMatch(/onlineMeetingInfo and meeting join URLs are not exposed/);
    expect(desc(MEMBERS)).toMatch(/no Graph \$top is sent/);
    expect(desc(MEMBERS)).toMatch(/other member types expose only id and @odata\.type/);
    expect(desc(MEMBERS)).toMatch(/Membership ids are opaque/);
  });

  it("exposes only the typed options", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = (n: string) =>
      Object.keys(
        (tools.find((t) => t.name === n)?.inputSchema.properties ?? {}) as Record<string, unknown>,
      ).sort();
    expect(props(LIST)).toEqual(["expand", "orderby", "top"]);
    expect(props(GET)).toEqual(["chat-id"]);
    expect(props(MEMBERS)).toEqual(["chat-id", "top"]);
  });
});

describe(LIST, () => {
  it("sends only $top (no $select: Graph documents none) to /me/chats", async () => {
    responses.push(() => json({ value: [chat()] }));
    const { env } = await call(LIST);
    expect(env.ok).toBe(true);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe("/v1.0/me/chats");
    expect([...parsed().searchParams.keys()]).toEqual(["$top"]);
    expect(parsed().searchParams.get("$top")).toBe("25");
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect((env.data as unknown[])[0]).toMatchObject({ id: CID, chatType: "group" });
  });

  it("caps the per-request $top at 50 and follows validated pages up to the limit", async () => {
    const link = "https://graph.microsoft.com/v1.0/me/chats?$top=50&$skiptoken=abc";
    responses.push(() => json({ value: [chat()], "@odata.nextLink": link }));
    responses.push(() => json({ value: [chat({ id: OTHER_CID })] }));
    const { env } = await call(LIST, { top: 100 });
    expect(env.ok).toBe(true);
    expect(parsed().searchParams.get("$top")).toBe("50");
    expect(calls).toHaveLength(2);
    expect(parsed(1).pathname).toBe("/v1.0/me/chats");
    expect((env.data as unknown[]).length).toBe(2);
  });

  it("sends the fixed descending order and the typed expansion", async () => {
    responses.push(() => json({ value: [] }));
    await call(LIST, {
      orderby: [{ property: ORDER, direction: "desc" }],
      expand: [PREVIEW],
    });
    expect(parsed().searchParams.get("$orderby")).toBe(`${ORDER} desc`);
    expect(parsed().searchParams.get("$expand")).toBe(PREVIEW);
  });

  it("rejects everything outside the typed contract with zero network calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      { top: 0 },
      { top: 101 },
      { top: 1.5 },
      { filter: "chatType eq 'oneOnOne'" },
      { $filter: "chatType eq 'oneOnOne'" },
      { search: "x" },
      { select: ["topic"] },
      { skip: 1 },
      { count: true },
      { expand: ["members"] },
      { expand: ["members", PREVIEW] },
      { expand: [] },
      { expand: ["installedApps"] },
      { orderby: [{ property: ORDER }] }, // bare property is ascending
      { orderby: [{ property: ORDER, direction: "asc" }] },
      { orderby: [{ property: "lastUpdatedDateTime", direction: "desc" }] },
      { orderby: [{ property: "createdDateTime", direction: "desc" }] },
      { orderby: [{ property: "topic", direction: "desc" }] },
      {
        orderby: [
          { property: ORDER, direction: "desc" },
          { property: ORDER, direction: "desc" },
        ],
      },
      { orderby: [] },
      { orderby: `${ORDER} desc` },
      { headers: { ConsistencyLevel: "eventual" } },
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

  it("projects to the allowlist: everything else, including content, is dropped", async () => {
    responses.push(() =>
      json({
        "@odata.context": "x",
        value: [
          chat({
            onlineMeetingInfo: { joinWebUrl: "https://teams.example/join-secret" },
            members: [{ id: "m", displayName: "member-secret" }],
            installedApps: [{ id: "app-secret" }],
            createdBy: { user: { id: "creator-secret" } },
            migrationMode: "inProgress",
            lastMessagePreview: { id: "1", body: { content: "preview-secret" } },
            viewpoint: { isHidden: true, lastMessageReadDateTime: "t", extra: "vp-secret" },
            topic: { nested: "object-secret" },
            "@odata.id": "evil",
          }),
          "not-an-object",
          ["array"],
        ],
      }),
    );
    const { env, text } = await call(LIST);
    expect(env.data).toEqual([
      {
        id: CID,
        createdDateTime: "2021-04-06T19:49:52.431Z",
        lastUpdatedDateTime: "2021-04-06T19:54:04.306Z",
        chatType: "group",
        webUrl: "https://teams.microsoft.com/l/chat/x/0?tenantId=t",
        tenantId: "f17e0000-0000-4000-8000-000000000014",
        isHiddenForAllMembers: false,
        viewpoint: { isHidden: true, lastMessageReadDateTime: "t" },
      },
    ]);
    for (const s of ["secret", "evil", "inProgress"]) expect(text).not.toContain(s);
  });

  it("keeps lastMessagePreview as metadata only, and only when it was requested", async () => {
    const withPreview = chat({
      lastMessagePreview: {
        id: "1622853091207",
        createdDateTime: "2021-06-05T00:31:31.207Z",
        isDeleted: false,
        messageType: "message",
        body: { contentType: "html", content: "<p>body-secret</p>" },
        eventDetail: { "@odata.type": "x", initiator: { user: { id: "event-secret" } } },
        from: {
          device: { id: "device-secret" },
          application: null,
          user: { id: "u1", displayName: "Nick Fury", userIdentityType: "aadUser", tenantId: "t" },
        },
      },
    });
    responses.push(() => json({ value: [withPreview] }));
    const asked = await call(LIST, { expand: [PREVIEW] });
    expect((asked.env.data as Array<Record<string, unknown>>)[0]?.["lastMessagePreview"]).toEqual({
      id: "1622853091207",
      createdDateTime: "2021-06-05T00:31:31.207Z",
      isDeleted: false,
      messageType: "message",
      from: { application: null, user: { id: "u1", displayName: "Nick Fury" } },
    });
    for (const s of ["body-secret", "event-secret", "device-secret", "userIdentityType"]) {
      expect(asked.text).not.toContain(s);
    }
    // Over-returned by Graph although not requested: dropped.
    responses.push(() => json({ value: [withPreview] }));
    const notAsked = await call(LIST);
    expect((notAsked.env.data as Array<Record<string, unknown>>)[0]).not.toHaveProperty(PREVIEW);
    expect(notAsked.text).not.toContain("body-secret");
  });

  it("keeps a null preview and a system-event preview (from null) without leaking content", async () => {
    responses.push(() =>
      json({
        value: [
          chat({ lastMessagePreview: null }),
          chat({
            lastMessagePreview: {
              id: "2",
              messageType: "systemEventMessage",
              from: null,
              body: { content: "<systemEventMessage/>" },
            },
          }),
        ],
      }),
    );
    const { env, text } = await call(LIST, { expand: [PREVIEW] });
    const items = env.data as Array<Record<string, unknown>>;
    expect(items[0]?.["lastMessagePreview"]).toBeNull();
    expect(items[1]?.["lastMessagePreview"]).toEqual({
      id: "2",
      messageType: "systemEventMessage",
      from: null,
    });
    expect(text).not.toContain("systemEventMessage/>");
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "backend failure" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(LIST)).env.error?.code, JSON.stringify(body)).toBe("graph_error");
    }
  });

  it("refuses a nextLink that leaves /me/chats, the origin or the version (fail closed)", async () => {
    for (const bad of [
      "https://evil.example/v1.0/me/chats?$skiptoken=1",
      "https://graph.microsoft.com/v1.0/chats?$skiptoken=1",
      `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(CID)}/members`,
      `https://graph.microsoft.com/v1.0/me/chats/${encodeURIComponent(CID)}`,
      "https://graph.microsoft.com/v1.0/users/f17e0000-0000-4000-8000-000000000010/chats",
      "https://graph.microsoft.com/beta/me/chats?$skiptoken=1",
      "http://graph.microsoft.com/v1.0/me/chats?$skiptoken=1",
      "https://graph.microsoft.com:444/v1.0/me/chats?$skiptoken=1",
      "https://graph.microsoft.com.evil.example/v1.0/me/chats",
    ]) {
      calls = [];
      responses.push(() => json({ value: [chat()], "@odata.nextLink": bad }));
      const { env } = await call(LIST, { top: 5 });
      expect(env.ok, bad).toBe(false);
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("maps 401, 403 and 404/400 to the shared error conventions", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [400, "graph_error"],
      [404, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      const { env, text } = await call(LIST);
      expect(env.error?.code, String(status)).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe(GET, () => {
  it("requests /chats/{id} with no query and no headers, and projects to the allowlist", async () => {
    responses.push(() =>
      json(
        chat({
          "@odata.context": "ctx",
          onlineMeetingInfo: { joinWebUrl: "https://teams.example/join-secret" },
          members: [{ id: "m", displayName: "member-secret" }],
          lastMessagePreview: { id: "1", body: { content: "preview-secret" } },
          createdBy: { user: { id: "creator-secret" } },
        }),
      ),
    );
    const { env, text } = await call(GET, { "chat-id": CID });
    expect(env.ok).toBe(true);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe(`/v1.0/chats/${encodeURIComponent(CID)}`);
    expect([...parsed().searchParams.keys()]).toEqual([]);
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(env.data).toEqual(chat());
    for (const s of ["secret", "ctx"]) expect(text).not.toContain(s);
  });

  it("accepts a one-on-one chat id and the 48:notes self chat", async () => {
    for (const id of [ONE_ON_ONE, "48:notes", "19:meeting_MjdhNjM4YzUtYzExZi00OTFk@thread.v2"]) {
      calls = [];
      responses.push(() => json(chat({ id })));
      const { env } = await call(GET, { "chat-id": id });
      expect(env.ok, id).toBe(true);
      expect(calls, id).toHaveLength(1);
    }
  });

  it("rejects options and unknown inputs with zero network calls", async () => {
    for (const extra of [
      { select: ["topic"] },
      { expand: ["members"] },
      { expand: [PREVIEW] },
      { top: 1 },
      { filter: "x" },
      { unknownKey: 1 },
    ]) {
      calls = [];
      const { env } = await call(GET, { "chat-id": CID, ...extra });
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });

  it("fails closed on a non-object 200 body", async () => {
    for (const body of [[], "x", null]) {
      responses.push(() => json(body));
      expect((await call(GET, { "chat-id": CID })).env.error?.code, JSON.stringify(body)).toBe(
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
      const { env, text } = await call(GET, { "chat-id": CID });
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
    visibleHistoryStartDateTime: "0001-01-01T00:00:00Z",
    userId: "f17e0000-0000-4000-8000-000000000011",
    email: "john@contoso.com",
    tenantId: "f17e0000-0000-4000-8000-000000000014",
    ...over,
  });

  it("requests /chats/{id}/members with NO query (no $top, no $select) and projects the users", async () => {
    responses.push(() => json({ "@odata.context": "ctx", "@odata.count": 1, value: [member()] }));
    const { env, text } = await call(MEMBERS, { "chat-id": CID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/chats/${encodeURIComponent(CID)}/members`);
    expect([...parsed().searchParams.keys()]).toEqual([]);
    expect(env.data).toEqual([member()]);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
    expect(text).not.toContain("ctx");
  });

  it("projects per @odata.type: other types id + type only, missing type id only", () => {
    expect(projectChatMember(member({ extra: "x", roles: ["owner", 5, null] }))).toEqual(
      member({ roles: ["owner"] }),
    );
    expect(
      projectChatMember({
        "@odata.type": GUEST_T,
        id: "g",
        displayName: "Guest",
        email: "g@x.example",
        roles: ["guest"],
      }),
    ).toEqual({ "@odata.type": GUEST_T, id: "g" });
    expect(
      projectChatMember({ id: "x", displayName: "N", email: "e@x.example", roles: [] }),
    ).toEqual({ id: "x" });
    expect(
      projectChatMember({ "@odata.type": { evil: 1 }, id: "x", displayName: "N", userId: "u" }),
    ).toEqual({ id: "x" });
    expect(projectChatMember({ "@odata.type": AAD_T, id: { evil: 1 }, email: ["e"] })).toEqual({
      "@odata.type": AAD_T,
    });
    for (const bad of [null, "x", 1, ["a"], undefined]) {
      expect(projectChatMember(bad)).toBeUndefined();
    }
  });

  it("applies the member projection end to end", async () => {
    responses.push(() =>
      json({
        value: [
          member(),
          { "@odata.type": GUEST_T, id: "g", displayName: "guest-secret", email: "g@x.example" },
          { id: "typeless", displayName: "typeless-secret" },
          "not-an-object",
        ],
      }),
    );
    const { env, text } = await call(MEMBERS, { "chat-id": CID });
    expect(env.data).toEqual([member(), { "@odata.type": GUEST_T, id: "g" }, { id: "typeless" }]);
    expect(text).not.toContain("secret");
  });

  it("treats top as a local limit and follows pages without sending $top", async () => {
    const link = `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(CID)}/members?$skiptoken=abc`;
    responses.push(() => json({ value: [member({ id: "1" })], "@odata.nextLink": link }));
    responses.push(() => json({ value: [member({ id: "2" }), member({ id: "3" })] }));
    const { env } = await call(MEMBERS, { "chat-id": CID, top: 2 });
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
      { filter: "x" },
      { expand: ["x"] },
      { search: "x" },
      { skip: 1 },
      { unknownKey: 1 },
    ]) {
      calls = [];
      const { env } = await call(MEMBERS, { "chat-id": CID, ...extra });
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });

  it("refuses a nextLink for another chat, collection, host or version (fail closed)", async () => {
    for (const bad of [
      `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(OTHER_CID)}/members?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(CID)}/messages`,
      `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(CID)}`,
      "https://graph.microsoft.com/v1.0/me/chats",
      `https://evil.example/v1.0/chats/${encodeURIComponent(CID)}/members`,
      `https://graph.microsoft.com/beta/chats/${encodeURIComponent(CID)}/members`,
      `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(CID)}%2f..%2f${encodeURIComponent(OTHER_CID)}/members`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [member()], "@odata.nextLink": bad }));
      const { env } = await call(MEMBERS, { "chat-id": CID, top: 5 });
      expect(env.ok, bad).toBe(false);
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("accepts the OData key-syntax form Graph uses in links for the same chat", async () => {
    const link = `https://graph.microsoft.com/v1.0/chats('${encodeURIComponent(CID)}')/members?$skiptoken=abc`;
    responses.push(() => json({ value: [member({ id: "1" })], "@odata.nextLink": link }));
    responses.push(() => json({ value: [member({ id: "2" })] }));
    const { env } = await call(MEMBERS, { "chat-id": CID, top: 5 });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(2);
  });

  it("fails closed on a 200 page without a value array and maps errors", async () => {
    responses.push(() => json({ error: "backend failure" }));
    expect((await call(MEMBERS, { "chat-id": CID })).env.error?.code).toBe("graph_error");
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
      [400, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      expect((await call(MEMBERS, { "chat-id": CID })).env.error?.code, String(status)).toBe(code);
    }
  });
});

describe("projectChat", () => {
  it("keeps the allowlist only and yields undefined for non-objects", () => {
    expect(projectChat(chat({ extra: 1 }), false)).toEqual(chat());
    expect(projectChat({ id: CID, topic: { a: 1 }, viewpoint: "x" }, false)).toEqual({ id: CID });
    for (const bad of [null, "x", 1, [], undefined]) expect(projectChat(bad, true)).toBeUndefined();
  });
});

describe("chat-id injection (zero network calls)", () => {
  const bad = [
    "19:abc@thread.v2/messages",
    "19:x@thread.v2/messages",
    "19%3Aabc%40thread.v2%2Fmessages",
    "19:abc@thread.v2%2F..%2F..",
    "..",
    ".",
    "../chats",
    "%2e%2e",
    "19:abc@thread.v2?$expand=members",
    "19:abc@thread.v2#frag",
    "19:abc@thread.v2$ref",
    "$ref",
    "$count",
    "delta",
    "getAllMessages",
    " 19:abc@thread.v2",
    "19:abc@thread.v2 ",
    "19:abc @thread.v2",
    "19:abc@thread.v2\n",
    "19:abc@thread.v2\\x",
    `19:${"a".repeat(600)}@thread.v2`,
    "f17e0000-0000-4000-8000-000000000010",
    "user@contoso.com",
    "19:abc@evil.example",
    "19:abc@thread.v2@thread.v2",
    "19:@thread.v2",
    "48:notes/messages",
    "48:other",
    "microsoft.graph.chat",
    "",
  ];
  it.each([GET, MEMBERS])("%s rejects malformed or route-changing chat ids", async (name) => {
    for (const id of bad) {
      calls = [];
      const { env, text } = await call(name, { "chat-id": id });
      expect(env.ok, JSON.stringify(id)).toBe(false);
      expect(env.error?.code, JSON.stringify(id)).toBe("invalid_input");
      expect(calls, JSON.stringify(id)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("rejects a missing or non-string chat id", async () => {
    for (const name of [GET, MEMBERS]) {
      for (const args of [{}, { "chat-id": 5 }, { "chat-id": null }, { "chat-id": [CID] }]) {
        calls = [];
        const { env } = await call(name, args);
        expect(env.ok, `${name} ${JSON.stringify(args)}`).toBe(false);
        expect(calls).toHaveLength(0);
      }
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

  it("registers all three tools with a batch policy", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const n of [LIST, GET, MEMBERS]) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
    }
  });

  it("accepts the in-policy forms and projects like the single path", async () => {
    responses.push(
      ok({
        value: [
          chat({
            onlineMeetingInfo: { joinWebUrl: "join-secret" },
            lastMessagePreview: { id: "1" },
          }),
        ],
      }),
    );
    const a = await batch(`/me/chats?$top=50&$orderby=${encodeURIComponent(`${ORDER} desc`)}`);
    expect(a.env.ok).toBe(true);
    expect(a.text).not.toContain("secret");
    expect(a.text).not.toContain("lastMessagePreview");
    expect(calls[0]?.url).toContain("/v1.0/$batch");

    responses.push(
      ok({ value: [chat({ lastMessagePreview: { id: "1", body: { content: "b-secret" } } })] }),
    );
    const p = await batch(`/me/chats?$expand=${PREVIEW}`);
    expect(p.text).toContain('"lastMessagePreview":{"id":"1"}');
    expect(p.text).not.toContain("b-secret");

    for (const id of [encodeURIComponent(CID), CID]) {
      responses.push(ok(chat({ onlineMeetingInfo: { joinWebUrl: "join-secret" } })));
      const g = await batch(`/chats/${id}`);
      expect(g.env.ok, id).toBe(true);
      expect(g.text).not.toContain("secret");
      expect(bodyOf(g)).toEqual(chat());
    }

    responses.push(
      ok({
        value: [
          member(),
          { "@odata.type": GUEST_T, id: "g", displayName: "guest-secret" },
          { id: "typeless", email: "typeless-secret" },
        ],
      }),
    );
    const m = await batch(`/chats/${encodeURIComponent(CID)}/members`);
    expect(bodyOf(m)?.value).toEqual([
      member(),
      { "@odata.type": GUEST_T, id: "g" },
      { id: "typeless" },
    ]);
    expect(m.text).not.toContain("secret");
  });

  const member = () => ({
    "@odata.type": AAD_T,
    id: "bWVtYmVy",
    roles: [],
    displayName: "John Doe",
    userId: "u",
    email: "j@x.example",
  });

  it("returns the same page size on the batch path as on the single path", async () => {
    const page = {
      value: Array.from({ length: 50 }, (_, i) => chat({ id: `19:c${i}@thread.v2` })),
    };
    responses.push(() => json(page));
    const single = (await call(LIST, { top: 50 })).env.data as unknown[];
    responses.push(ok(page));
    const b = bodyOf(await batch("/me/chats?$top=50"));
    expect(single).toHaveLength(50);
    expect(b?.value).toHaveLength(50);
    expect(b?.hasMore).toBeUndefined();
    // Without $top the batch clips to the single tool's default (25), flagging hasMore.
    responses.push(ok(page));
    const d = bodyOf(await batch("/me/chats"));
    expect(d?.value).toHaveLength(25);
    expect(d?.hasMore).toBe(true);
    // Members: clipped to the local-limit default (100) and flagged.
    responses.push(
      ok({ value: Array.from({ length: 120 }, (_, i) => ({ "@odata.type": AAD_T, id: `m${i}` })) }),
    );
    const mm = bodyOf(await batch(`/chats/${encodeURIComponent(CID)}/members`));
    expect(mm?.value).toHaveLength(100);
    expect(mm?.hasMore).toBe(true);
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const enc = encodeURIComponent;
    const cid = enc(CID);
    const bad: Array<[string, Record<string, string>?]> = [
      ["/me/chats?$top=51"],
      ["/me/chats?$top=0"],
      ["/me/chats?$top=-1"],
      ["/me/chats?$top=5&$top=6"],
      ["/me/chats?$select=topic"],
      ["/me/chats?$filter=chatType%20eq%20'oneOnOne'"],
      ["/me/chats?$search=%22x%22"],
      ["/me/chats?$count=true"],
      ["/me/chats?$skip=1"],
      ["/me/chats?$expand=members"],
      ["/me/chats?$expand=members,lastMessagePreview"],
      [`/me/chats?$orderby=${enc(ORDER)}`],
      [`/me/chats?$orderby=${enc(`${ORDER} asc`)}`],
      [`/me/chats?$orderby=${enc("lastUpdatedDateTime desc")}`],
      [`/me/chats?$orderby=${enc("createdDateTime desc")}`],
      ["/me/chats", { ConsistencyLevel: "eventual" }],
      ["/me/chats", { Prefer: "odata.maxpagesize=1" }],
      ["/chats"],
      ["/me/chats/delta"],
      [`/me/chats/${cid}`],
      [`/chats/${cid}?$select=topic`],
      [`/chats/${cid}?$expand=members`],
      [`/chats/${cid}?$expand=${PREVIEW}`],
      [`/chats/${cid}?$top=1`],
      [`/chats/${cid}`, { ConsistencyLevel: "eventual" }],
      [`/chats/${cid}/members?$top=5`],
      [`/chats/${cid}/members?$select=email`],
      [`/chats/${cid}/members?$filter=x`],
      [`/chats/${cid}/members`, { ConsistencyLevel: "eventual" }],
      [`/chats/${cid}/members/microsoft.graph.aadUserConversationMember`],
      [`/chats/${cid}/members/$count`],
      [`/chats/${cid}/installedApps`],
      [`/chats/${cid}/tabs`],
      [`/chats/${enc(`${CID}/messages`)}`],
      [`/chats/${CID}/messages/${enc(`1/replies`)}`], // part B: the plain list path is now valid; its replies/injection forms are not
      ["/chats/..%2fme"],
      ["/chats/$ref"],
      ["/chats/getAllMessages"],
      ["/chats/19:abc@thread.v2%2f..%2fx/members"],
      ["/beta/me/chats"],
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
const SM11A: Case[] = [
  { name: LIST, args: {}, paged: true, next: "/me/chats" },
  { name: GET, args: { "chat-id": CID }, paged: false, next: "" },
  {
    name: MEMBERS,
    args: { "chat-id": CID },
    paged: true,
    next: `/chats/${encodeURIComponent(CID)}/members`,
  },
];

describe("SM-11 part A tool set", () => {
  it("covers exactly the chat/members tools implemented in this part", () => {
    const names = catalog
      .filter(
        (e) =>
          e.method === "get" &&
          (e.pathPattern === "/me/chats" ||
            /^\/chats\/\{chat-id\}(\/members)?$/.test(e.pathPattern)) &&
          e.toolName !== undefined,
      )
      .map((e) => e.toolName)
      .sort();
    expect(names).toEqual(SM11A.map((c) => c.name).sort());
  });
});

describe.each(SM11A)("$name: unavailable permission", (c) => {
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

describe.each(SM11A.filter((c) => c.paged))("$name: permission lost mid-listing", (c) => {
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
