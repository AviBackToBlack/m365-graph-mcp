import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import {
  GET_BODY_MAX_CHARS,
  LIST_BODY_MAX_CHARS,
  projectChatMessage,
  projectPinnedMessage,
} from "./chat-messages.js";
import { TRUNCATION_MARKER } from "./html-text.js";
import { createHandlers } from "./index.js";
import { probeUnimplementedGet } from "./unimplemented-get-fixture.js";

/**
 * SM-11 part B: list-chat-messages, get-chat-message, list-pinned-chat-messages. Replies are not
 * implemented (Microsoft documents them for channel messages only); hosted contents are part C.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const LIST = "list-chat-messages";
const GET = "get-chat-message";
const PINNED = "list-pinned-chat-messages";
const CID = "19:2da4c29f6d7041eca70b638b43d45437@thread.v2";
const OTHER_CID = "19:ebe3857aa388434bab0cad9d2e09f4de@thread.v2";
const MID = "1615971548136";
const ECID = encodeURIComponent(CID);

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

const msg = (over: Record<string, unknown> = {}) => ({
  id: MID,
  replyToId: null,
  etag: MID,
  messageType: "message",
  createdDateTime: "2021-03-17T08:59:08.136Z",
  lastModifiedDateTime: "2021-03-17T08:59:08.136Z",
  lastEditedDateTime: null,
  deletedDateTime: null,
  subject: null,
  summary: null,
  chatId: CID,
  importance: "normal",
  locale: "en-us",
  webUrl: null,
  channelIdentity: null,
  onBehalfOf: null,
  policyViolation: null,
  eventDetail: null,
  from: {
    application: null,
    device: null,
    user: {
      "@odata.type": "#microsoft.graph.teamworkUserIdentity",
      id: "f17e0000-0000-4000-8000-000000000012",
      displayName: "Robin Kline",
      userIdentityType: "aadUser",
      tenantId: "f17e0000-0000-4000-8000-000000000018",
    },
  },
  body: { contentType: "text", content: "Hello world" },
  attachments: [],
  mentions: [],
  reactions: [],
  messageHistory: [],
  ...over,
});

/** `msg()` after projection. */
const projected = (over: Record<string, unknown> = {}) => ({
  id: MID,
  replyToId: null,
  messageType: "message",
  createdDateTime: "2021-03-17T08:59:08.136Z",
  lastModifiedDateTime: "2021-03-17T08:59:08.136Z",
  lastEditedDateTime: null,
  deletedDateTime: null,
  subject: null,
  importance: "normal",
  chatId: CID,
  webUrl: null,
  from: {
    application: null,
    user: { id: "f17e0000-0000-4000-8000-000000000012", displayName: "Robin Kline" },
  },
  body: { contentType: "text", sourceContentType: "text", content: "Hello world" },
  mentions: [],
  attachments: [],
  reactions: [],
  ...over,
});

const htmlBody = (content: string) => ({ contentType: "html", content });
const bodyOf = (item: unknown) => (item as { body: { content: string } }).body.content;

describe("catalog and surface", () => {
  it("pins the three v1.0 GET entries and the Chat.Read scope (D20)", () => {
    const by = (n: string) => catalog.find((e) => e.toolName === n);
    const expected: Array<[string, string, string]> = [
      [LIST, "/chats/{chat-id}/messages", "Chat.Read"],
      [GET, "/chats/{chat-id}/messages/{chatMessage-id}", "Chat.Read"],
      [PINNED, "/chats/{chat-id}/pinnedMessages", "Chat.Read"],
    ];
    for (const [name, path, scope] of expected) {
      expect(by(name)?.pathPattern).toBe(path);
      expect(by(name)?.method).toBe("get");
      expect(by(name)?.["apiVersion"]).toBeUndefined();
      expect(by(name)?.["workScopes"]).toEqual([scope]);
    }
  });

  it("has no chat-message replies capability (D21: channel-only in the docs)", async () => {
    expect(catalog.some((e) => e.toolName === "list-chat-message-replies")).toBe(false);
    expect(
      catalog.some((e) => /^\/chats\/[^/]+\/messages\/[^/]+\/replies$/.test(e.pathPattern)),
    ).toBe(false);
    const tools = (await mcp.listTools()).tools;
    expect(tools.some((t) => t.name === "list-chat-message-replies")).toBe(false);
    // Channel replies stay a separate catalog capability (SM-12).
    expect(catalog.some((e) => e.toolName === "list-channel-message-replies")).toBe(true);
  });

  it("describes the typed contracts, not raw Graph syntax (D22)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    const tip = (n: string) => String(catalog.find((x) => x.toolName === n)?.["llmTip"]);
    for (const n of [LIST, GET, PINNED]) {
      expect(desc(n)).toMatch(/Requires Chat\.Read \(delegated/);
      expect(desc(n)).toMatch(/raw HTML is never returned/);
      expect(desc(n)).toMatch(/truncated: true/);
      expect(desc(n)).toMatch(
        /Attachment content and attachment URLs, reaction-user identities and policy details are not exposed/,
      );
      expect(desc(n) + tip(n)).not.toMatch(/list-meeting-transcripts|ChatMessage\.Read|Use \$/);
    }
    expect(desc(LIST)).toMatch(/at most 50 messages are requested from Graph per page/);
    expect(desc(LIST)).toMatch(/descending only, on createdDateTime or lastModifiedDateTime/);
    expect(desc(LIST)).toMatch(/modifiedAfter, modifiedBefore and createdBefore/);
    expect(desc(LIST)).toMatch(/conflicting range\/order combinations are rejected/);
    expect(desc(LIST)).toMatch(/capped at 2,000 characters/);
    expect(desc(GET)).toMatch(/Accepts no Graph query options/);
    expect(desc(GET)).toMatch(/capped at 20,000 characters/);
    expect(desc(PINNED)).toMatch(/no Graph \$top is sent/);
    expect(desc(PINNED)).toMatch(/only typed expansion is message/);
    expect(desc(PINNED)).toMatch(/2,000-character body cap/);
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

  it("exposes only the typed options", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = (n: string) =>
      Object.keys(
        (tools.find((t) => t.name === n)?.inputSchema.properties ?? {}) as Record<string, unknown>,
      ).sort();
    expect(props(LIST)).toEqual([
      "chat-id",
      "createdBefore",
      "modifiedAfter",
      "modifiedBefore",
      "orderby",
      "top",
    ]);
    expect(props(GET)).toEqual(["chat-id", "chatMessage-id"]);
    expect(props(PINNED)).toEqual(["chat-id", "expand", "top"]);
  });
});

describe(LIST, () => {
  it("requests /chats/{id}/messages with only $top: no $select, no Prefer header", async () => {
    responses.push(() => json({ value: [msg()] }));
    const { env } = await call(LIST, { "chat-id": CID });
    expect(env.ok).toBe(true);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe(`/v1.0/chats/${ECID}/messages`);
    expect([...parsed().searchParams.keys()]).toEqual(["$top"]);
    expect(parsed().searchParams.get("$top")).toBe("20");
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(env.data).toEqual([projected()]);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
  });

  it("caps the per-request $top at 50 and follows validated pages up to the limit", async () => {
    const link = `https://graph.microsoft.com/v1.0/chats/${CID}/messages?$top=50&$skiptoken=M2UyZDAw%3d%3d`;
    responses.push(() => json({ value: [msg({ id: "2" })], "@odata.nextLink": link }));
    responses.push(() => json({ value: [msg({ id: "1" })] }));
    const { env } = await call(LIST, { "chat-id": CID, top: 100 });
    expect(env.ok).toBe(true);
    expect(parsed().searchParams.get("$top")).toBe("50");
    expect(calls).toHaveLength(2);
    expect(parsed(1).pathname).toBe(`/v1.0/chats/${CID}/messages`);
    expect((env.data as unknown[]).length).toBe(2);
  });

  it("sends the descending order the documentation allows", async () => {
    for (const property of ["createdDateTime", "lastModifiedDateTime"]) {
      calls = [];
      responses.push(() => json({ value: [] }));
      await call(LIST, { "chat-id": CID, orderby: [{ property, direction: "desc" }], top: 2 });
      expect(parsed().searchParams.get("$orderby")).toBe(`${property} desc`);
      expect(parsed().searchParams.has("$filter")).toBe(false);
    }
  });

  it("turns typed ranges into the documented filter with the coupled orderby", async () => {
    const cases: Array<[Record<string, unknown>, string, string]> = [
      [
        { modifiedAfter: "2022-09-22T00:00:00Z", modifiedBefore: "2022-09-24T00:00:00.000Z" },
        "lastModifiedDateTime gt 2022-09-22T00:00:00.000Z and lastModifiedDateTime lt 2022-09-24T00:00:00.000Z",
        "lastModifiedDateTime desc",
      ],
      [
        { modifiedAfter: "2022-09-22T02:00:00+02:00" },
        "lastModifiedDateTime gt 2022-09-22T00:00:00.000Z",
        "lastModifiedDateTime desc",
      ],
      [
        // Seven fraction digits are fine when the sub-millisecond part is zero.
        { modifiedBefore: "2022-09-24T00:00:00.1230000Z" },
        "lastModifiedDateTime lt 2022-09-24T00:00:00.123Z",
        "lastModifiedDateTime desc",
      ],
      [
        { createdBefore: "2022-09-24T00:00:00Z" },
        "createdDateTime lt 2022-09-24T00:00:00.000Z",
        "createdDateTime desc",
      ],
      [
        {
          modifiedBefore: "2022-09-24T00:00:00Z",
          orderby: [{ property: "lastModifiedDateTime", direction: "desc" }],
        },
        "lastModifiedDateTime lt 2022-09-24T00:00:00.000Z",
        "lastModifiedDateTime desc",
      ],
    ];
    for (const [args, filter, orderby] of cases) {
      calls = [];
      responses.push(() => json({ value: [] }));
      const { env } = await call(LIST, { "chat-id": CID, ...args });
      expect(env.ok, JSON.stringify(args)).toBe(true);
      expect(parsed().searchParams.get("$filter"), JSON.stringify(args)).toBe(filter);
      expect(parsed().searchParams.get("$orderby"), JSON.stringify(args)).toBe(orderby);
    }
  });

  it("rejects a sub-millisecond cutoff instead of rounding the boundary", async () => {
    for (const key of ["modifiedAfter", "modifiedBefore", "createdBefore"]) {
      calls = [];
      const { env } = await call(LIST, { "chat-id": CID, [key]: "2022-09-24T00:00:00.1234567Z" });
      expect(env.error?.code, key).toBe("invalid_input");
      expect(calls, key).toHaveLength(0);
    }
  });

  it("rejects everything outside the typed contract with zero network calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      { top: 0 },
      { top: 101 },
      { top: 1.5 },
      { filter: "lastModifiedDateTime gt 2022-09-22T00:00:00.000Z" },
      { $filter: "x" },
      { select: ["body"] },
      { expand: ["replies"] },
      { search: "x" },
      { skip: 1 },
      { count: true },
      { bodyFormat: "html" },
      { bodyType: "html" },
      { headers: { Prefer: "outlook.body-content-type=html" } },
      { orderby: [{ property: "createdDateTime" }] }, // bare property is ascending
      { orderby: [{ property: "createdDateTime", direction: "asc" }] },
      { orderby: [{ property: "lastModifiedDateTime", direction: "asc" }] },
      { orderby: [{ property: "subject", direction: "desc" }] },
      { orderby: [{ property: "from/user/id", direction: "desc" }] },
      {
        orderby: [
          { property: "createdDateTime", direction: "desc" },
          { property: "lastModifiedDateTime", direction: "desc" },
        ],
      },
      { orderby: [] },
      { orderby: "createdDateTime desc" },
      // typed ranges
      { modifiedAfter: "yesterday" },
      { modifiedAfter: "2022-09-22T00:00:00" }, // no zone
      { modifiedAfter: "2022-02-30T00:00:00Z" },
      { modifiedAfter: "2022-09-22T00:00:00Z' or 1 eq 1" },
      { modifiedBefore: 5 },
      { createdBefore: "" },
      { createdAfter: "2022-09-22T00:00:00Z" },
      { modifiedAfter: "2022-09-24T00:00:00Z", modifiedBefore: "2022-09-22T00:00:00Z" },
      { modifiedAfter: "2022-09-22T00:00:00Z", modifiedBefore: "2022-09-22T00:00:00Z" },
      { createdBefore: "2022-09-24T00:00:00Z", modifiedBefore: "2022-09-24T00:00:00Z" },
      { createdBefore: "2022-09-24T00:00:00Z", modifiedAfter: "2022-09-20T00:00:00Z" },
      {
        modifiedAfter: "2022-09-22T00:00:00Z",
        orderby: [{ property: "createdDateTime", direction: "desc" }],
      },
      {
        createdBefore: "2022-09-22T00:00:00Z",
        orderby: [{ property: "lastModifiedDateTime", direction: "desc" }],
      },
      { unknownKey: 1 },
    ];
    for (const args of bad) {
      calls = [];
      const { env, text } = await call(LIST, { "chat-id": CID, ...args });
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code, JSON.stringify(args)).toBe("invalid_input");
      expect(calls, JSON.stringify(args)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("projects to the allowlist: unknown, nested and content-bearing fields are dropped", async () => {
    responses.push(() =>
      json({
        "@odata.context": "ctx-secret",
        "@odata.count": 99,
        value: [
          msg({
            "@odata.type": "#microsoft.graph.chatMessage",
            etag: "etag-secret",
            summary: "summary-secret",
            policyViolation: { dlpAction: "policy-secret" },
            onBehalfOf: { user: { id: "behalf-secret" } },
            channelIdentity: { teamId: "team-secret" },
            messageHistory: [{ actions: "history-secret" }],
            unknownProp: "unknown-secret",
            subject: "Release plan",
            webUrl: "https://teams.microsoft.com/l/message/x",
            from: {
              application: { id: "app1", displayName: "Bot", applicationIdentityType: "bot" },
              device: { id: "device-secret" },
              conversation: { id: "conv-secret" },
              user: null,
            },
            attachments: [
              {
                id: "att1",
                contentType: "reference",
                name: "plan.docx",
                contentUrl: "https://contoso.sharepoint.com/url-secret",
                thumbnailUrl: "https://thumb.example/thumb-secret",
                content: '{"originalMessageContent":"forwarded-secret"}',
                teamsAppId: "app-secret",
              },
              "not-an-object",
            ],
            mentions: [
              {
                id: 0,
                mentionText: "Jane",
                mentioned: {
                  device: { id: "mdevice-secret" },
                  user: {
                    id: "u1",
                    displayName: "Jane",
                    userIdentityType: "aadUser",
                    tenantId: "tenant-secret",
                  },
                  conversation: {
                    id: CID,
                    displayName: "Everyone",
                    conversationIdentityType: "chat",
                  },
                },
              },
              { id: 1, mentionText: "x", mentioned: "nope" },
            ],
            reactions: [
              {
                reactionType: "like",
                createdDateTime: "t",
                user: { user: { id: "reactor-secret" } },
                reactionContentUrl: "https://react.example/rc-secret",
              },
              { reactionType: "like", user: { user: { id: "reactor2-secret" } } },
              { reactionType: "custom", displayName: "party", reactionContentUrl: "u" },
              { notAReaction: true },
            ],
            eventDetail: {
              "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail",
              chatDisplayName: "name-secret",
              initiator: { user: { id: "initiator-secret" } },
            },
          }),
          "not-an-object",
          ["array"],
          null,
        ],
      }),
    );
    const { env, text } = await call(LIST, { "chat-id": CID });
    expect(env.data).toEqual([
      projected({
        subject: "Release plan",
        webUrl: "https://teams.microsoft.com/l/message/x",
        from: { application: { id: "app1", displayName: "Bot" }, user: null },
        attachments: [{ id: "att1", contentType: "reference", name: "plan.docx" }],
        mentions: [
          {
            id: 0,
            mentionText: "Jane",
            mentioned: {
              user: { id: "u1", displayName: "Jane" },
              conversation: { id: CID, displayName: "Everyone", conversationIdentityType: "chat" },
            },
          },
          { id: 1, mentionText: "x" },
        ],
        reactions: [
          { reactionType: "like", count: 2 },
          { reactionType: "custom", displayName: "party", count: 1 },
        ],
        eventDetail: { "@odata.type": "#microsoft.graph.chatRenamedEventMessageDetail" },
      }),
    ]);
    expect(text).not.toMatch(/secret/);
    for (const s of ["@odata.context", "@odata.count", "userIdentityType", "etag"]) {
      expect(text).not.toContain(s);
    }
  });

  it("normalizes HTML bodies to plain text and never returns raw HTML", async () => {
    const hosted = `https://graph.microsoft.com/v1.0/chats/${CID}/messages/${MID}/hostedContents/aWQ9QQ==/$value`;
    responses.push(() =>
      json({
        value: [
          msg({
            body: htmlBody(
              `<div><p>Hi&nbsp;<at id="0">Jane</at> &amp; team</p><script>alert(1)</script><attachment id="42"></attachment> <img src="${hosted}"> <a href="https://example.com/x">link</a></div>`,
            ),
          }),
          msg({
            id: "2",
            body: { contentType: "text", content: "plain <b>not html</b> &amp; text" },
          }),
        ],
      }),
    );
    const { env, text } = await call(LIST, { "chat-id": CID });
    const items = env.data as Array<{ body: Record<string, unknown> }>;
    expect(items[0]?.body).toEqual({
      contentType: "text",
      sourceContentType: "html",
      content:
        "Hi @Jane & team\n[attachment: 42] [image hostedContentId=aWQ9QQ==] link (https://example.com/x)",
    });
    // Text bodies are never interpreted.
    expect(items[1]?.body).toEqual({
      contentType: "text",
      sourceContentType: "text",
      content: "plain <b>not html</b> &amp; text",
    });
    expect(text).not.toContain("alert(1)");
    expect(text).not.toContain("<div>");
  });

  it("caps list bodies at 2000 characters with a visible marker and a truncated flag", async () => {
    responses.push(() =>
      json({
        value: [
          msg({ body: htmlBody(`<p>${"word ".repeat(1000)}</p>`) }),
          msg({ id: "2", body: { contentType: "text", content: "x".repeat(5000) } }),
          msg({ id: "3", body: htmlBody("<p>short</p>") }),
        ],
      }),
    );
    const { env } = await call(LIST, { "chat-id": CID });
    const items = env.data as Array<{ body: Record<string, unknown> }>;
    for (const i of [0, 1]) {
      expect(items[i]?.body["truncated"], String(i)).toBe(true);
      expect(bodyOf(items[i])).toHaveLength(LIST_BODY_MAX_CHARS + TRUNCATION_MARKER.length);
      expect(bodyOf(items[i]).endsWith(TRUNCATION_MARKER)).toBe(true);
    }
    expect(items[2]?.body).not.toHaveProperty("truncated");
  });

  it("keeps deleted messages and system events without leaking content", async () => {
    responses.push(() =>
      json({
        value: [
          msg({
            id: "3",
            messageType: "unknownFutureValue",
            from: null,
            body: htmlBody("<systemEventMessage/>"),
            eventDetail: { "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail" },
          }),
          msg({
            id: "4",
            deletedDateTime: "2021-03-18T00:00:00.000Z",
            body: { contentType: "html", content: "" },
          }),
        ],
      }),
    );
    const { env } = await call(LIST, { "chat-id": CID });
    const [event, deleted] = env.data as Array<Record<string, unknown>>;
    expect(event).toMatchObject({
      messageType: "unknownFutureValue",
      from: null,
      body: { sourceContentType: "html", content: "" },
      eventDetail: { "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail" },
    });
    expect(deleted).toMatchObject({ deletedDateTime: "2021-03-18T00:00:00.000Z" });
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "backend failure" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(LIST, { "chat-id": CID })).env.error?.code, JSON.stringify(body)).toBe(
        "graph_error",
      );
    }
  });

  it("refuses a nextLink for another chat, message, collection, host or version (fail closed)", async () => {
    for (const bad of [
      `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(OTHER_CID)}/messages?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/chats/${OTHER_CID}/messages?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages/${MID}`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages/${MID}/replies`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages/${MID}/hostedContents`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/pinnedMessages`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/members`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}`,
      `https://graph.microsoft.com/v1.0/me/chats/${ECID}/messages`,
      `https://graph.microsoft.com/v1.0/users/f17e0000-0000-4000-8000-000000000010/chats/${ECID}/messages`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}%2f..%2f${encodeURIComponent(OTHER_CID)}/messages`,
      `https://evil.example/v1.0/chats/${ECID}/messages?$skiptoken=1`,
      `https://graph.microsoft.com.evil.example/v1.0/chats/${ECID}/messages`,
      `https://graph.microsoft.com/beta/chats/${ECID}/messages?$skiptoken=1`,
      `http://graph.microsoft.com/v1.0/chats/${ECID}/messages?$skiptoken=1`,
      `https://graph.microsoft.com:444/v1.0/chats/${ECID}/messages?$skiptoken=1`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [msg()], "@odata.nextLink": bad }));
      const { env } = await call(LIST, { "chat-id": CID, top: 5 });
      expect(env.ok, bad).toBe(false);
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("accepts Graph's documented nextLink forms for the same chat", async () => {
    for (const link of [
      `https://graph.microsoft.com/v1.0/chats/${CID}/messages?$top=2&$skiptoken=M2UyZDAw%3d%3d`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages?$top=2&$skiptoken=abc`,
      `https://graph.microsoft.com/v1.0/chats('${ECID}')/messages?$top=2&$skiptoken=abc`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [msg({ id: "2" })], "@odata.nextLink": link }));
      responses.push(() => json({ value: [msg({ id: "1" })] }));
      const { env } = await call(LIST, { "chat-id": CID, top: 5 });
      expect(env.ok, link).toBe(true);
      expect(calls, link).toHaveLength(2);
      expect((env.data as unknown[]).length).toBe(2);
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
      const { env, text } = await call(LIST, { "chat-id": CID });
      expect(env.error?.code, String(status)).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe(GET, () => {
  it("requests /chats/{id}/messages/{mid} with no query and no headers, and projects", async () => {
    responses.push(() =>
      json(
        msg({
          "@odata.context": "ctx-secret",
          etag: "etag-secret",
          body: htmlBody("<div>Hi <b>there</b><br>second&nbsp;line</div>"),
        }),
      ),
    );
    const { env, text } = await call(GET, { "chat-id": CID, "chatMessage-id": MID });
    expect(env.ok).toBe(true);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe(`/v1.0/chats/${ECID}/messages/${MID}`);
    expect([...parsed().searchParams.keys()]).toEqual([]);
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(env.data).toEqual(
      projected({
        body: { contentType: "text", sourceContentType: "html", content: "Hi there\nsecond line" },
      }),
    );
    expect(text).not.toContain("secret");
  });

  it("caps the body at 20000 characters (far above the list cap)", async () => {
    responses.push(() => json(msg({ body: htmlBody(`<p>${"word ".repeat(10_000)}</p>`) })));
    const { env } = await call(GET, { "chat-id": CID, "chatMessage-id": MID });
    const body = (env.data as { body: Record<string, unknown> }).body;
    expect(body["truncated"]).toBe(true);
    expect(bodyOf(env.data)).toHaveLength(GET_BODY_MAX_CHARS + TRUNCATION_MARKER.length);
    responses.push(() => json(msg({ body: htmlBody(`<p>${"word ".repeat(1000)}</p>`) })));
    const mid = await call(GET, { "chat-id": CID, "chatMessage-id": MID });
    expect(bodyOf(mid.env.data).length).toBeGreaterThan(LIST_BODY_MAX_CHARS);
    expect((mid.env.data as { body: object }).body).not.toHaveProperty("truncated");
  });

  it("rejects options and unknown inputs with zero network calls", async () => {
    for (const extra of [
      { select: ["body"] },
      { expand: ["replies"] },
      { top: 1 },
      { filter: "x" },
      { bodyFormat: "html" },
      { unknownKey: 1 },
    ]) {
      calls = [];
      const { env } = await call(GET, { "chat-id": CID, "chatMessage-id": MID, ...extra });
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });

  it("fails closed on a non-object 200 body", async () => {
    for (const body of [[], "x", null]) {
      responses.push(() => json(body));
      expect(
        (await call(GET, { "chat-id": CID, "chatMessage-id": MID })).env.error?.code,
        JSON.stringify(body),
      ).toBe("graph_error");
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
      const { env, text } = await call(GET, { "chat-id": CID, "chatMessage-id": MID });
      expect(env.error?.code, String(status)).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe(PINNED, () => {
  const pin = (over: Record<string, unknown> = {}) => ({
    "@odata.type": "#microsoft.graph.pinnedChatMessageInfo",
    id: "1616964509832",
    ...over,
  });

  it("requests /chats/{id}/pinnedMessages with NO query by default and returns ids", async () => {
    responses.push(() =>
      json({ "@odata.context": "ctx-secret", value: [pin(), pin({ id: "2" })] }),
    );
    const { env, text } = await call(PINNED, { "chat-id": CID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/chats/${ECID}/pinnedMessages`);
    expect([...parsed().searchParams.keys()]).toEqual([]);
    expect(env.data).toEqual([pin(), pin({ id: "2" })]);
    expect(env.meta).toMatchObject({ count: 2, hasMore: false });
    expect(text).not.toContain("secret");
  });

  it("expands the message on request and normalizes it like list messages", async () => {
    responses.push(() =>
      json({
        value: [
          pin({
            message: msg({
              body: htmlBody("<p>Pinned&nbsp;<b>note</b></p>"),
              etag: "etag-secret",
              policyViolation: { x: "policy-secret" },
            }),
            extra: "extra-secret",
          }),
          pin({ id: "9", message: null }),
        ],
      }),
    );
    const { env, text } = await call(PINNED, { "chat-id": CID, expand: ["message"] });
    expect(parsed().searchParams.get("$expand")).toBe("message");
    expect(parsed().searchParams.has("$top")).toBe(false);
    expect(env.data).toEqual([
      pin({
        message: projected({
          body: { contentType: "text", sourceContentType: "html", content: "Pinned note" },
        }),
      }),
      pin({ id: "9", message: null }),
    ]);
    expect(text).not.toMatch(/secret/);
  });

  it("drops an unrequested expansion Graph returned anyway", async () => {
    responses.push(() =>
      json({ value: [pin({ message: msg({ body: htmlBody("body-secret") }) })] }),
    );
    const { env, text } = await call(PINNED, { "chat-id": CID });
    expect(env.data).toEqual([pin()]);
    expect(text).not.toContain("secret");
  });

  it("treats top as a local limit and follows pages without sending $top", async () => {
    const link = `https://graph.microsoft.com/v1.0/chats/${CID}/pinnedMessages?$skiptoken=abc`;
    responses.push(() => json({ value: [pin({ id: "1" })], "@odata.nextLink": link }));
    responses.push(() => json({ value: [pin({ id: "2" }), pin({ id: "3" })] }));
    const { env } = await call(PINNED, { "chat-id": CID, top: 2 });
    expect(calls).toHaveLength(2);
    expect(parsed(0).searchParams.has("$top")).toBe(false);
    expect((env.data as unknown[]).length).toBe(2);
    expect(env.meta).toMatchObject({ count: 2, hasMore: true });
  });

  it("rejects options with zero network calls", async () => {
    for (const extra of [
      { top: 0 },
      { top: 101 },
      { expand: ["members"] },
      { expand: ["message", "replies"] },
      { expand: [] },
      { select: ["id"] },
      { filter: "x" },
      { orderby: [{ property: "createdDateTime", direction: "desc" }] },
      { modifiedAfter: "2022-09-22T00:00:00Z" },
      { skip: 1 },
      { unknownKey: 1 },
    ]) {
      calls = [];
      const { env } = await call(PINNED, { "chat-id": CID, ...extra });
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });

  it("refuses a nextLink for another chat, collection, host or version (fail closed)", async () => {
    for (const bad of [
      `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(OTHER_CID)}/pinnedMessages?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages/${MID}`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/pinnedMessages/${MID}`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/members`,
      `https://evil.example/v1.0/chats/${ECID}/pinnedMessages`,
      `https://graph.microsoft.com/beta/chats/${ECID}/pinnedMessages`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [pin()], "@odata.nextLink": bad }));
      const { env } = await call(PINNED, { "chat-id": CID, top: 5 });
      expect(env.ok, bad).toBe(false);
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("fails closed on a 200 page without a value array and maps errors", async () => {
    responses.push(() => json({ error: "backend failure" }));
    expect((await call(PINNED, { "chat-id": CID })).env.error?.code).toBe("graph_error");
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
      [400, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "x", message: "m" } }, status));
      expect((await call(PINNED, { "chat-id": CID })).env.error?.code, String(status)).toBe(code);
    }
  });
});

describe("projectChatMessage / projectPinnedMessage", () => {
  it("treats any body contentType other than exactly text as HTML (fail closed)", () => {
    const markup = "<b>x</b><script>y</script>";
    for (const contentType of [undefined, "HTML", "Text", "unknownFutureValue", 5]) {
      const out = projectChatMessage(msg({ body: { contentType, content: markup } }), 100) as {
        body: { content: string; sourceContentType: string };
      };
      expect(out.body.content, String(contentType)).toBe("x");
      // The declared type is reported as a fixed value, never echoed raw.
      expect(out.body.sourceContentType).toBe("unknown");
    }
    const html = projectChatMessage(msg({ body: { contentType: "html", content: markup } }), 100);
    expect((html as { body: { sourceContentType: string } }).body.sourceContentType).toBe("html");
    const text = projectChatMessage(msg({ body: { contentType: "text", content: markup } }), 100);
    expect((text as { body: { content: string } }).body.content).toBe(markup);
  });

  it("groups reactions by an unambiguous (type, name) key", () => {
    const reactions = [
      { reactionType: "a\u0000b", displayName: "c" },
      { reactionType: "a", displayName: "b\u0000c" },
      { reactionType: "like", displayName: "" },
      { reactionType: "like" },
    ];
    const out = projectChatMessage(msg({ reactions }), 100) as { reactions: unknown[] };
    expect(out.reactions).toHaveLength(4);
  });

  it("stops scanning attachments once the nested cap is reached", () => {
    const many = [
      ...Array.from({ length: 60 }, (_, i) => ({ id: `a${i}`, contentType: "x", content: "c" })),
      ...Array.from({ length: 10_000 }, () => "not-an-object"),
    ];
    const out = projectChatMessage(msg({ attachments: many }), 100) as { attachments: unknown[] };
    expect(out.attachments).toHaveLength(50);
    expect(out.attachments[0]).toEqual({ id: "a0", contentType: "x" });
  });

  it("keeps the allowlist only and yields undefined for non-objects", () => {
    expect(projectChatMessage(msg({ extra: 1 }), 100)).toEqual(projected());
    expect(
      projectChatMessage(
        {
          id: { evil: 1 },
          subject: ["x"],
          from: "x",
          body: "x",
          mentions: "x",
          attachments: {},
          reactions: 5,
          eventDetail: "x",
        },
        100,
      ),
    ).toEqual({});
    for (const bad of [null, "x", 1, [], undefined]) {
      expect(projectChatMessage(bad, 100)).toBeUndefined();
      expect(projectPinnedMessage(bad, true)).toBeUndefined();
    }
  });

  it("bounds nested arrays", () => {
    const many = Array.from({ length: 200 }, (_, i) => i);
    const out = projectChatMessage(
      msg({
        mentions: many.map((i) => ({ id: i, mentionText: "m" })),
        attachments: many.map((i) => ({ id: String(i), name: "n" })),
        reactions: many.map((i) => ({ reactionType: `r${i}` })),
      }),
      100,
    ) as Record<string, unknown[]>;
    expect(out["mentions"]).toHaveLength(50);
    expect(out["attachments"]).toHaveLength(50);
    expect(out["reactions"]).toHaveLength(50);
  });

  it("normalizes a body without a string content to empty text", () => {
    expect(projectChatMessage({ body: { contentType: "html", content: { a: 1 } } }, 100)).toEqual({
      body: { contentType: "text", sourceContentType: "html", content: "" },
    });
    expect(projectChatMessage({ body: null }, 100)).toEqual({ body: null });
  });

  it("only surfaces a hosted-content id for the message's own chat and id", () => {
    const own = `https://graph.microsoft.com/v1.0/chats/${CID}/messages/${MID}/hostedContents/abc/$value`;
    const other = own.replace(MID, "1");
    const out = (src: string, over: Record<string, unknown> = {}) =>
      bodyOf(projectChatMessage(msg({ body: htmlBody(`<img src="${src}">`), ...over }), 100));
    expect(out(own)).toBe("[image hostedContentId=abc]");
    expect(out(other)).toBe("[image]");
    expect(out(own, { chatId: OTHER_CID })).toBe("[image]");
    expect(out(own, { id: "7" })).toBe("[image]");
  });
});

describe("path id injection (zero network calls)", () => {
  const badChat = [
    "19:abc@thread.v2/messages",
    "19:x@thread.v2/messages",
    "19%3Aabc%40thread.v2%2Fmessages",
    "19:abc@thread.v2%2F..%2F..",
    "..",
    "../chats",
    "%2e%2e",
    "19:abc@thread.v2?$top=1",
    "19:abc@thread.v2#frag",
    "$ref",
    "delta",
    " 19:abc@thread.v2",
    "19:abc@thread.v2\n",
    `19:${"a".repeat(600)}@thread.v2`,
    "f17e0000-0000-4000-8000-000000000010",
    "user@contoso.com",
    "",
  ];
  const badMessage = [
    "",
    "abc",
    "1615971548136/replies",
    "1615971548136%2Freplies",
    "1615971548136/hostedContents",
    "1615971548136?$expand=replies",
    "1615971548136#x",
    "..",
    "%2e%2e",
    "$value",
    "$ref",
    "delta",
    "$count",
    "-1",
    "1.5",
    "1e5",
    " 1615971548136",
    "1615971548136 ",
    "1615971548136\n",
    "0x10",
    "1".repeat(21),
    "microsoft.graph.chatMessage",
    "replies",
    "hostedContents",
  ];

  it.each([LIST, GET, PINNED])("%s rejects malformed or route-changing chat ids", async (name) => {
    for (const id of badChat) {
      calls = [];
      const args: Record<string, unknown> = { "chat-id": id };
      if (name === GET) args["chatMessage-id"] = MID;
      const { env, text } = await call(name, args);
      expect(env.ok, JSON.stringify(id)).toBe(false);
      expect(env.error?.code, JSON.stringify(id)).toBe("invalid_input");
      expect(calls, JSON.stringify(id)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("get-chat-message rejects malformed or route-changing message ids", async () => {
    for (const id of badMessage) {
      calls = [];
      const { env, text } = await call(GET, { "chat-id": CID, "chatMessage-id": id });
      expect(env.ok, JSON.stringify(id)).toBe(false);
      expect(env.error?.code, JSON.stringify(id)).toBe("invalid_input");
      expect(calls, JSON.stringify(id)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("rejects missing or non-string ids", async () => {
    for (const args of [
      { name: LIST, a: {} },
      { name: LIST, a: { "chat-id": 5 } },
      { name: PINNED, a: { "chat-id": null } },
      { name: GET, a: { "chat-id": CID } },
      { name: GET, a: { "chatMessage-id": MID } },
      { name: GET, a: { "chat-id": CID, "chatMessage-id": 1615971548136 } },
      { name: GET, a: { "chat-id": CID, "chatMessage-id": [MID] } },
    ]) {
      calls = [];
      const { env } = await call(args.name, args.a);
      expect(env.ok, JSON.stringify(args)).toBe(false);
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
  const batchBody = (r: { env: Env }) =>
    (
      r.env.data as {
        responses: Array<{
          body?: { value?: unknown[]; hasMore?: boolean } & Record<string, unknown>;
        }>;
      }
    ).responses[0]?.body;
  const enc = encodeURIComponent;

  it("registers all three tools with a batch policy", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const n of [LIST, GET, PINNED]) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
    }
  });

  it("projects and normalizes bodies exactly like the single path (list, get, pinned)", async () => {
    const raw = msg({
      etag: "etag-secret",
      policyViolation: { x: "policy-secret" },
      body: htmlBody("<div>Hi&nbsp;<at id='0'>Jane</at><script>alert(1)</script></div>"),
      attachments: [{ id: "a", name: "n", contentUrl: "https://x.example/url-secret" }],
    });
    responses.push(() => json({ value: [raw] }));
    const single = (await call(LIST, { "chat-id": CID })).env.data;
    responses.push(ok({ value: [raw], "@odata.context": "ctx-secret" }));
    const l = await batch(`/chats/${ECID}/messages?$top=20`);
    expect(l.env.ok).toBe(true);
    expect(batchBody(l)?.value).toEqual(single);
    expect(l.text).not.toMatch(/secret|alert\(1\)|<div>/);

    for (const id of [ECID, CID]) {
      responses.push(ok(raw));
      const g = await batch(`/chats/${id}/messages/${MID}`);
      expect(g.env.ok, id).toBe(true);
      expect(batchBody(g)).toEqual(
        projectChatMessage(raw, GET_BODY_MAX_CHARS) as Record<string, unknown>,
      );
      expect(g.text).not.toMatch(/secret|alert\(1\)|<div>/);
    }

    const pin = {
      "@odata.type": "#microsoft.graph.pinnedChatMessageInfo",
      id: MID,
      message: raw,
      extra: "extra-secret",
    };
    responses.push(ok({ value: [pin] }));
    const p = await batch(`/chats/${ECID}/pinnedMessages?$expand=message`);
    expect(batchBody(p)?.value).toEqual([projectPinnedMessage(pin, true)]);
    expect(p.text).not.toMatch(/secret|alert\(1\)|<div>/);
    responses.push(ok({ value: [pin] }));
    const q = await batch(`/chats/${ECID}/pinnedMessages`);
    expect(batchBody(q)?.value).toEqual([
      { "@odata.type": "#microsoft.graph.pinnedChatMessageInfo", id: MID },
    ]);
  });

  it("applies the same body caps in batch as on the single path", async () => {
    const big = msg({ body: htmlBody(`<p>${"word ".repeat(6000)}</p>`) });
    responses.push(ok({ value: [big] }));
    const l = batchBody(await batch(`/chats/${ECID}/messages`));
    expect(bodyOf(l?.value?.[0])).toHaveLength(LIST_BODY_MAX_CHARS + TRUNCATION_MARKER.length);
    responses.push(ok(big));
    const g = batchBody(await batch(`/chats/${ECID}/messages/${MID}`));
    expect(bodyOf(g)).toHaveLength(GET_BODY_MAX_CHARS + TRUNCATION_MARKER.length);
  });

  it("accepts the typed range only with its coupled orderby, in the single path's exact form", async () => {
    const good = [
      "lastModifiedDateTime gt 2022-09-22T00:00:00.000Z and lastModifiedDateTime lt 2022-09-24T00:00:00.000Z",
      "lastModifiedDateTime gt 2022-09-22T00:00:00.000Z",
      "lastModifiedDateTime lt 2022-09-24T00:00:00.000Z",
    ];
    for (const filter of good) {
      calls = [];
      responses.push(ok({ value: [] }));
      const r = await batch(
        `/chats/${ECID}/messages?$top=2&$orderby=${enc("lastModifiedDateTime desc")}&$filter=${enc(filter)}`,
      );
      expect(r.env.ok, filter).toBe(true);
    }
    calls = [];
    responses.push(ok({ value: [] }));
    const created = await batch(
      `/chats/${ECID}/messages?$orderby=${enc("createdDateTime desc")}&$filter=${enc("createdDateTime lt 2022-09-24T00:00:00.000Z")}`,
    );
    expect(created.env.ok).toBe(true);
  });

  it("returns the same page size on the batch path as on the single path", async () => {
    const page = { value: Array.from({ length: 50 }, (_, i) => msg({ id: String(1000 + i) })) };
    responses.push(() => json(page));
    const single = (await call(LIST, { "chat-id": CID, top: 50 })).env.data as unknown[];
    responses.push(ok(page));
    const b = batchBody(await batch(`/chats/${ECID}/messages?$top=50`));
    expect(single).toHaveLength(50);
    expect(b?.value).toHaveLength(50);
    expect(b?.hasMore).toBeUndefined();
    responses.push(ok(page));
    const d = batchBody(await batch(`/chats/${ECID}/messages`));
    expect(d?.value).toHaveLength(20);
    expect(d?.hasMore).toBe(true);
    responses.push(ok({ value: Array.from({ length: 40 }, (_, i) => ({ id: String(i) })) }));
    const pp = batchBody(await batch(`/chats/${ECID}/pinnedMessages`));
    expect(pp?.value).toHaveLength(25);
    expect(pp?.hasMore).toBe(true);
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const cid = ECID;
    const lm = "lastModifiedDateTime";
    const ts = "2022-09-22T00:00:00.000Z";
    const bad: Array<[string, Record<string, string>?]> = [
      [`/chats/${cid}/messages?$top=51`],
      [`/chats/${cid}/messages?$top=0`],
      [`/chats/${cid}/messages?$top=5&$top=6`],
      [`/chats/${cid}/messages?$select=body`],
      [`/chats/${cid}/messages?$search=%22x%22`],
      [`/chats/${cid}/messages?$count=true`],
      [`/chats/${cid}/messages?$skip=1`],
      [`/chats/${cid}/messages?$expand=replies`],
      [`/chats/${cid}/messages?$orderby=${enc("createdDateTime")}`],
      [`/chats/${cid}/messages?$orderby=${enc("createdDateTime asc")}`],
      [`/chats/${cid}/messages?$orderby=${enc("subject desc")}`],
      [`/chats/${cid}/messages`, { Prefer: 'outlook.body-content-type="html"' }],
      [`/chats/${cid}/messages`, { Prefer: "include-unknown-enum-members" }],
      [`/chats/${cid}/messages`, { ConsistencyLevel: "eventual" }],
      // raw or malformed filters
      [`/chats/${cid}/messages?$filter=${enc(`${lm} gt ${ts}`)}`], // no coupled orderby
      [
        `/chats/${cid}/messages?$orderby=${enc(`${lm} desc`)}&$filter=${enc(`createdDateTime lt ${ts}`)}`,
      ],
      [
        `/chats/${cid}/messages?$orderby=${enc("createdDateTime desc")}&$filter=${enc(`${lm} lt ${ts}`)}`,
      ],
      [
        `/chats/${cid}/messages?$orderby=${enc("createdDateTime desc")}&$filter=${enc(`createdDateTime gt ${ts}`)}`,
      ],
      [
        `/chats/${cid}/messages?$orderby=${enc(`${lm} desc`)}&$filter=${enc(`${lm} gt ${ts} or ${lm} lt ${ts}`)}`,
      ],
      [
        `/chats/${cid}/messages?$orderby=${enc(`${lm} desc`)}&$filter=${enc(`${lm} gt ${ts} and subject eq 'x'`)}`,
      ],
      [
        `/chats/${cid}/messages?$orderby=${enc(`${lm} desc`)}&$filter=${enc(`${lm} gt 2022-09-22T00:00:00Z`)}`,
      ],
      [
        `/chats/${cid}/messages?$orderby=${enc(`${lm} desc`)}&$filter=${enc(`${lm} gt 2022-09-22T00:00:00.000+01:00`)}`,
      ],
      [
        `/chats/${cid}/messages?$orderby=${enc(`${lm} desc`)}&$filter=${enc(`${lm} gt 2022-02-30T00:00:00.000Z`)}`,
      ],
      [
        `/chats/${cid}/messages?$orderby=${enc(`${lm} desc`)}&$filter=${enc(`${lm} gt 2022-09-24T00:00:00.000Z and ${lm} lt ${ts}`)}`,
      ],
      [
        `/chats/${cid}/messages?$orderby=${enc(`${lm} desc`)}&$filter=${enc(`${lm} gt ${ts}`)}&$filter=${enc(`${lm} gt ${ts}`)}`,
      ],
      [`/chats/${cid}/messages?$orderby=${enc(`${lm} desc`)}&$filter=`],
      // replies and hosted-content bytes are not reachable (hosted-content LISTS: see the part C tests)
      [`/chats/${cid}/messages/${MID}/replies`],
      [`/chats/${cid}/messages/${MID}/replies?$top=5`],
      [`/chats/${cid}/messages/${MID}/hostedContents/abc/$value`],
      // get message
      [`/chats/${cid}/messages/${MID}?$select=body`],
      [`/chats/${cid}/messages/${MID}?$expand=replies`],
      [`/chats/${cid}/messages/${MID}?$top=1`],
      [`/chats/${cid}/messages/${MID}`, { Prefer: 'outlook.body-content-type="html"' }],
      [`/chats/${cid}/messages/${enc("1/replies")}`],
      [`/chats/${cid}/messages/abc`],
      [`/chats/${cid}/messages/$count`],
      [`/chats/${cid}/messages/delta`],
      [`/chats/${cid}/messages/delta()`],
      ["/chats/getAllMessages"],
      [`/me/chats/${cid}/messages`],
      [`/users/f17e0000-0000-4000-8000-000000000010/chats/${cid}/messages`],
      [`/chats/${enc(`${CID}/x`)}/messages`],
      [`/chats/..%2fme/messages`],
      // pinned
      [`/chats/${cid}/pinnedMessages?$top=5`],
      [`/chats/${cid}/pinnedMessages?$expand=members`],
      [`/chats/${cid}/pinnedMessages?$expand=message,replies`],
      [`/chats/${cid}/pinnedMessages?$select=id`],
      [`/chats/${cid}/pinnedMessages?$filter=x`],
      [`/chats/${cid}/pinnedMessages/${MID}`],
      [`/chats/${cid}/pinnedMessages`, { ConsistencyLevel: "eventual" }],
      [`/beta/chats/${cid}/messages`],
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
const SM11B: Case[] = [
  { name: LIST, args: { "chat-id": CID }, paged: true, next: `/chats/${ECID}/messages` },
  { name: GET, args: { "chat-id": CID, "chatMessage-id": MID }, paged: false, next: "" },
  {
    name: PINNED,
    args: { "chat-id": CID },
    paged: true,
    next: `/chats/${ECID}/pinnedMessages`,
  },
];

describe.each(SM11B)("$name: unavailable permission", (c) => {
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

describe.each(SM11B.filter((c) => c.paged))("$name: permission lost mid-listing", (c) => {
  it("a 403 or 401 on a followed page fails the call instead of returning the first page", async () => {
    for (const [status, code] of [
      [403, "capability_gap"],
      [401, "reauth_required"],
    ] as const) {
      calls = [];
      responses.push(() =>
        json({
          value: [{ id: "1" }],
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
