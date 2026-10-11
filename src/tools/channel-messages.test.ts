import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { GET_BODY_MAX_CHARS, LIST_BODY_MAX_CHARS, projectChannelMessage } from "./chat-messages.js";
import { TRUNCATION_MARKER } from "./html-text.js";
import { createHandlers } from "./index.js";
import { probeUnimplementedGet } from "./unimplemented-get-fixture.js";

/**
 * SM-12 part B: list-channel-messages, get-channel-message, list-channel-message-replies and
 * list-channel-message-hosted-contents. Mirrors chat-messages.test.ts / chat-hosted-contents.test.ts
 * for the channel scope (typed `top` only, no range/orderby/expand, channel-scoped hosted-content
 * markers, same-team/channel/message pagination).
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const LIST = "list-channel-messages";
const GET = "get-channel-message";
const REPLIES = "list-channel-message-replies";
const HOSTED = "list-channel-message-hosted-contents";
const TID = "f17e0000-0000-4000-8000-000000000020";
const OTHER_TID = "f17e0000-0000-4000-8000-000000000009";
const CID = "19:4a95f7d8db4c4e7fae857bcebe0623e6@thread.tacv2";
const OTHER_CID = "19:982abbfca323a582f0a6d00ae2deca@thread.tacv2";
const MID = "1616963377068";
const RID = "1616989753153";
const ECID = encodeURIComponent(CID);
const enc = encodeURIComponent;
const BASE = `/teams/${TID}/channels/${CID}`;
const EBASE = `/teams/${TID}/channels/${ECID}`;

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

const hostedSrc = (message = MID, team = TID, channel = CID, id = "abc") =>
  `https://graph.microsoft.com/v1.0/teams/${team}/channels/${channel}/messages/${message}/hostedContents/${id}/$value`;

const msg = (over: Record<string, unknown> = {}) => ({
  id: MID,
  replyToId: null,
  etag: MID,
  messageType: "message",
  createdDateTime: "2021-03-28T20:29:37.068Z",
  lastModifiedDateTime: "2021-03-28T20:29:37.068Z",
  lastEditedDateTime: null,
  deletedDateTime: null,
  subject: "Release notes",
  summary: "summary-secret",
  chatId: null,
  importance: "normal",
  locale: "en-us",
  webUrl: "https://teams.microsoft.com/l/message/x/1?groupId=g",
  policyViolation: null,
  eventDetail: null,
  onBehalfOf: null,
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
  body: { contentType: "text", content: "Hello channel" },
  channelIdentity: { teamId: TID, channelId: CID },
  attachments: [],
  mentions: [],
  reactions: [],
  messageHistory: [],
  ...over,
});

/** `msg()` after projection (no `summary`, `chatId`, `locale`, `etag`, ...). */
const projected = (over: Record<string, unknown> = {}) => ({
  id: MID,
  replyToId: null,
  messageType: "message",
  createdDateTime: "2021-03-28T20:29:37.068Z",
  lastModifiedDateTime: "2021-03-28T20:29:37.068Z",
  lastEditedDateTime: null,
  deletedDateTime: null,
  subject: "Release notes",
  importance: "normal",
  webUrl: "https://teams.microsoft.com/l/message/x/1?groupId=g",
  from: {
    application: null,
    user: { id: "f17e0000-0000-4000-8000-000000000012", displayName: "Robin Kline" },
  },
  channelIdentity: { teamId: TID, channelId: CID },
  body: { contentType: "text", sourceContentType: "text", content: "Hello channel" },
  mentions: [],
  attachments: [],
  reactions: [],
  ...over,
});

const htmlBody = (content: string) => ({ contentType: "html", content });
const bodyOf = (item: unknown) => (item as { body: { content: string } }).body.content;
const ids = { "team-id": TID, "channel-id": CID };

interface Case {
  name: string;
  args: Record<string, unknown>;
  paged: boolean;
  next: string;
  item: Record<string, unknown>;
}
const CASES: Case[] = [
  { name: LIST, args: ids, paged: true, next: `${BASE}/messages`, item: msg() },
  {
    name: REPLIES,
    args: { ...ids, "chatMessage-id": MID },
    paged: true,
    next: `${BASE}/messages/${MID}/replies`,
    item: msg({ id: RID, replyToId: MID }),
  },
  {
    name: HOSTED,
    args: { ...ids, "chatMessage-id": MID },
    paged: true,
    next: `${BASE}/messages/${MID}/hostedContents`,
    item: { id: "aWQ9", contentBytes: null, contentType: null },
  },
  {
    name: GET,
    args: { ...ids, "chatMessage-id": MID },
    paged: false,
    next: "",
    item: msg(),
  },
];

describe("catalog and surface", () => {
  it("pins four v1.0 GET entries with ChannelMessage.Read.All", () => {
    for (const c of CASES) {
      const e = catalog.find((x) => x.toolName === c.name);
      expect(e?.method, c.name).toBe("get");
      expect(e?.["apiVersion"], c.name).toBeUndefined();
      expect(e?.["workScopes"], c.name).toEqual(["ChannelMessage.Read.All"]);
    }
  });

  it("describes the typed contracts, not raw Graph syntax (D26)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    const tip = (n: string) => String(catalog.find((x) => x.toolName === n)?.["llmTip"]);
    for (const n of [LIST, GET, REPLIES, HOSTED]) {
      expect(desc(n)).toMatch(/channel-id is an opaque Teams channel\/thread identifier/);
      expect(desc(n)).toMatch(/work\/school accounts only/);
      expect(desc(n)).toMatch(/Requires ChannelMessage\.Read\.All \(delegated\)/);
      expect(desc(n) + tip(n)).not.toMatch(
        /download-bytes|\$value|base64|contentBytes|contentType/,
      );
    }
    for (const n of [LIST, GET, REPLIES]) {
      expect(desc(n)).toMatch(/raw HTML is never returned/);
      expect(desc(n)).toMatch(/channelIdentity \(teamId, channelId\) is exposed/);
      expect(desc(n)).toMatch(/expanded replies are not exposed/);
    }
    expect(desc(LIST)).toMatch(/root messages .*\(not replies\)/);
    expect(desc(LIST)).toMatch(/at most 50 messages are requested from Graph per page/);
    expect(desc(LIST)).toMatch(/entire reply chain/);
    expect(desc(LIST)).toMatch(/use list-channel-message-replies for replies/);
    expect(desc(LIST)).toMatch(/capped at 2,000 characters/);
    expect(desc(GET)).toMatch(/Accepts no Graph query options/);
    expect(desc(GET)).toMatch(/capped at 20,000 characters/);
    expect(desc(REPLIES)).toMatch(/at most 50 replies are requested from Graph per page/);
    expect(desc(HOSTED)).toMatch(/Reference enumeration only/);
    expect(desc(HOSTED)).toMatch(/outside this project's capability surface/);
    expect(desc(HOSTED)).toMatch(/no Graph \$top or other OData query parameter is sent/);
  });

  it("exposes only the typed options (no $expand, $filter, $orderby, $search, select)", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = (n: string) =>
      Object.keys(
        (tools.find((t) => t.name === n)?.inputSchema.properties ?? {}) as Record<string, unknown>,
      ).sort();
    expect(props(LIST)).toEqual(["channel-id", "team-id", "top"]);
    expect(props(REPLIES)).toEqual(["channel-id", "chatMessage-id", "team-id", "top"]);
    expect(props(GET)).toEqual(["channel-id", "chatMessage-id", "team-id"]);
    expect(props(HOSTED)).toEqual(["channel-id", "chatMessage-id", "team-id", "top"]);
  });
});

describe(LIST, () => {
  it("requests /teams/{t}/channels/{c}/messages with only $top and projects to the allowlist", async () => {
    responses.push(() => json({ value: [msg()], "@odata.context": "ctx-secret" }));
    const { env, text } = await call(LIST, ids);
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe(`/v1.0/teams/${TID}/channels/${ECID}/messages`);
    expect([...parsed().searchParams.keys()]).toEqual(["$top"]);
    expect(parsed().searchParams.get("$top")).toBe("20");
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(env.data).toEqual([projected()]);
    expect(text).not.toMatch(/secret|etag|locale|messageHistory|"chatId"/);
  });

  it("caps $top at 50 per Graph page and follows validated pages up to the limit", async () => {
    const page = (n: number) => Array.from({ length: n }, (_, i) => msg({ id: String(1000 + i) }));
    const link = `https://graph.microsoft.com/v1.0${BASE}/messages?$top=50&$skiptoken=abc`;
    responses.push(() => json({ value: page(50), "@odata.nextLink": link }));
    responses.push(() => json({ value: page(50) }));
    const { env } = await call(LIST, { ...ids, top: 70 });
    expect(env.ok).toBe(true);
    expect(parsed(0).searchParams.get("$top")).toBe("50");
    expect(calls).toHaveLength(2);
    expect((env.data as unknown[]).length).toBe(70);
  });

  it("drops unrequested expansions, unknown and nested content", async () => {
    const raw = msg({
      replies: [
        msg({ id: RID, replyToId: MID, body: { contentType: "text", content: "reply-secret" } }),
      ],
      "replies@odata.nextLink": "https://evil.example/x",
      "@odata.type": "#microsoft.graph.chatMessage",
      attachments: [
        {
          id: "a",
          contentType: "reference",
          name: "doc",
          contentUrl: "https://x/url-secret",
          content: "content-secret",
        },
      ],
      reactions: [
        {
          reactionType: "like",
          createdDateTime: "t",
          reactionContentUrl: "https://x/r-secret",
          user: { user: { id: "u-secret" } },
        },
        { reactionType: "like", user: { user: { id: "u2-secret" } } },
      ],
      eventDetail: {
        "@odata.type": "#microsoft.graph.channelAddedEventMessageDetail",
        channelDisplayName: "name-secret",
      },
      mentions: [
        {
          id: 0,
          mentionText: "Jane",
          mentioned: {
            user: { id: "u1", displayName: "Jane", userIdentityType: "aadUser" },
            extra: "x-secret",
          },
        },
      ],
      channelIdentity: { teamId: TID, channelId: CID, other: "ci-secret" },
    });
    responses.push(() => json({ value: [raw] }));
    const { env, text } = await call(LIST, ids);
    expect(env.ok).toBe(true);
    expect(text).not.toMatch(/secret/);
    expect(env.data).toEqual([
      projected({
        attachments: [{ id: "a", contentType: "reference", name: "doc" }],
        reactions: [{ reactionType: "like", count: 2 }],
        eventDetail: { "@odata.type": "#microsoft.graph.channelAddedEventMessageDetail" },
        mentions: [
          { id: 0, mentionText: "Jane", mentioned: { user: { id: "u1", displayName: "Jane" } } },
        ],
      }),
    ]);
  });

  it("normalizes HTML bodies to plain text and caps list bodies at 2000 characters", async () => {
    const html = htmlBody("<div>Hi&nbsp;<at id='0'>Jane</at><script>alert(1)</script></div>");
    const long = htmlBody(`<p>${"x".repeat(5000)}</p>`);
    responses.push(() => json({ value: [msg({ body: html }), msg({ id: "2", body: long })] }));
    const { env, text } = await call(LIST, ids);
    const [a, b] = env.data as Array<{ body: Record<string, unknown> }>;
    expect(a?.body["contentType"]).toBe("text");
    expect(a?.body["sourceContentType"]).toBe("html");
    expect(a?.body["content"]).not.toMatch(/<|alert\(1\)/);
    expect(text).not.toMatch(/<div>|<script>/i);
    expect(b?.body["truncated"]).toBe(true);
    expect((b?.body["content"] as string).endsWith(TRUNCATION_MARKER)).toBe(true);
    expect((b?.body["content"] as string).length).toBeLessThanOrEqual(
      LIST_BODY_MAX_CHARS + TRUNCATION_MARKER.length,
    );
  });

  it("rejects everything outside the typed contract with zero network calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      { top: 0 },
      { top: 101 },
      { top: 1.5 },
      { top: "5" },
      { expand: "replies" },
      { $expand: "replies" },
      { filter: "x" },
      { $filter: "lastModifiedDateTime gt 2021-01-01T00:00:00.000Z" },
      { orderby: "createdDateTime desc" },
      { search: "x" },
      { select: "body" },
      { modifiedAfter: "2021-01-01T00:00:00Z" },
      { skip: 1 },
      { count: true },
      { headers: { Prefer: "include-unknown-enum-members" } },
      { extra: 1 },
    ];
    for (const o of bad) {
      calls = [];
      const { env } = await call(LIST, { ...ids, ...o });
      expect(env.ok, JSON.stringify(o)).toBe(false);
      expect(env.error?.code, JSON.stringify(o)).toBe("invalid_input");
      expect(calls, JSON.stringify(o)).toHaveLength(0);
    }
  });
});

describe(GET, () => {
  it("requests the message with no query and no headers and projects with the 20000 cap", async () => {
    const body = htmlBody(`<p>${"y".repeat(GET_BODY_MAX_CHARS + 500)}</p>`);
    responses.push(() => json(msg({ body })));
    const { env, text } = await call(GET, { ...ids, "chatMessage-id": MID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0${EBASE}/messages/${MID}`);
    expect([...parsed().searchParams.keys()]).toEqual([]);
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    const out = env.data as { body: { content: string; truncated?: boolean } };
    expect(out.body.truncated).toBe(true);
    expect(out.body.content.length).toBeGreaterThan(LIST_BODY_MAX_CHARS * 5);
    expect(text).not.toContain("summary-secret");
  });

  it("returns a reply-shaped message as is (replyToId kept)", async () => {
    responses.push(() => json(msg({ id: RID, replyToId: MID })));
    const { env } = await call(GET, { ...ids, "chatMessage-id": RID });
    expect(env.data).toEqual(projected({ id: RID, replyToId: MID }));
  });

  it("rejects options with zero network calls and fails closed on a non-object body", async () => {
    for (const o of [{ top: 1 }, { expand: "replies" }, { select: "body" }, { extra: 1 }]) {
      calls = [];
      const { env } = await call(GET, { ...ids, "chatMessage-id": MID, ...o });
      expect(env.error?.code, JSON.stringify(o)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
    for (const body of [[], "x", null]) {
      responses.push(() => json(body));
      const { env } = await call(GET, { ...ids, "chatMessage-id": MID });
      expect(env.ok, JSON.stringify(body)).toBe(false);
    }
  });
});

describe(REPLIES, () => {
  it("requests .../replies with only $top, projects replies and scopes markers to the reply", async () => {
    const reply = msg({
      id: RID,
      replyToId: MID,
      body: htmlBody(
        `<img src="https://graph.microsoft.com/v1.0/teams/${TID}/channels/${CID}/messages/${MID}/replies/${RID}/hostedContents/rid1/$value">`,
      ),
    });
    responses.push(() => json({ value: [reply] }));
    const { env } = await call(REPLIES, { ...ids, "chatMessage-id": MID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0${EBASE}/messages/${MID}/replies`);
    expect([...parsed().searchParams.keys()]).toEqual(["$top"]);
    expect((env.data as unknown[]).length).toBe(1);
    expect(bodyOf((env.data as unknown[])[0])).toBe("[image hostedContentId=rid1]");
    expect((env.data as Array<{ replyToId: string }>)[0]?.replyToId).toBe(MID);
  });

  it("rejects options with zero network calls", async () => {
    for (const o of [{ top: 0 }, { top: 101 }, { expand: "x" }, { orderby: "createdDateTime" }]) {
      calls = [];
      const { env } = await call(REPLIES, { ...ids, "chatMessage-id": MID, ...o });
      expect(env.error?.code, JSON.stringify(o)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });
});

describe(HOSTED, () => {
  it("lists references only: id and a string contentType; contentBytes is always dropped", async () => {
    responses.push(() =>
      json({
        value: [
          { id: "aWQ9a", contentBytes: "SECRET-BYTES", contentType: null },
          {
            id: "aWQ9b",
            contentType: "image/png",
            contentBytes: "SECRET-BYTES2",
            "@odata.mediaReadLink": "m-secret",
          },
          { contentBytes: "no-id-secret" },
        ],
      }),
    );
    const { env, text } = await call(HOSTED, { ...ids, "chatMessage-id": MID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0${EBASE}/messages/${MID}/hostedContents`);
    expect(parsed().search).toBe("");
    expect(env.data).toEqual([{ id: "aWQ9a" }, { id: "aWQ9b", contentType: "image/png" }]);
    expect(text).not.toMatch(/SECRET|secret/);
  });

  it("rejects options with zero network calls", async () => {
    for (const o of [{ top: 0 }, { select: "id" }, { expand: "x" }, { value: true }]) {
      calls = [];
      const { env } = await call(HOSTED, { ...ids, "chatMessage-id": MID, ...o });
      expect(env.error?.code, JSON.stringify(o)).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    }
  });
});

describe("channel hosted-content markers", () => {
  const out = (src: string, over: Record<string, unknown> = {}) =>
    bodyOf(projectChannelMessage(msg({ body: htmlBody(`<img src="${src}">`), ...over }), 100));

  it("surfaces an id only for the message's own team, channel and message", () => {
    expect(out(hostedSrc())).toBe("[image hostedContentId=abc]");
    expect(out(hostedSrc().replace("/v1.0/", "/1.0/"))).toBe("[image hostedContentId=abc]");
    expect(out(hostedSrc(MID, TID, ECID))).toBe("[image hostedContentId=abc]");
    expect(out(hostedSrc() + "?x=1#f")).toBe("[image hostedContentId=abc]");
    // another message, team or channel
    expect(out(hostedSrc("1"))).toBe("[image]");
    expect(out(hostedSrc(MID, OTHER_TID))).toBe("[image]");
    expect(out(hostedSrc(MID, TID, OTHER_CID))).toBe("[image]");
    expect(out(hostedSrc(MID, TID, enc(OTHER_CID)))).toBe("[image]");
    expect(out(hostedSrc(MID, TID, `${ECID}%2f..%2f${enc(OTHER_CID)}`))).toBe("[image]");
    expect(out(hostedSrc(MID, TID, "%E0%A4%A"))).toBe("[image]");
    // the message's own scope changes the answer
    expect(out(hostedSrc(), { channelIdentity: { teamId: OTHER_TID, channelId: CID } })).toBe(
      "[image]",
    );
    expect(out(hostedSrc(), { channelIdentity: { teamId: TID, channelId: OTHER_CID } })).toBe(
      "[image]",
    );
    expect(out(hostedSrc(), { id: "7" })).toBe("[image]");
  });

  it("never yields an id without a usable channelIdentity or from a chat URL", () => {
    expect(out(hostedSrc(), { channelIdentity: null })).toBe("[image]");
    expect(out(hostedSrc(), { channelIdentity: { teamId: TID } })).toBe("[image]");
    expect(out(hostedSrc(), { channelIdentity: { teamId: 5, channelId: CID } })).toBe("[image]");
    const chatUrl = `https://graph.microsoft.com/v1.0/chats/${CID}/messages/${MID}/hostedContents/abc/$value`;
    expect(out(chatUrl)).toBe("[image]");
    // a channel projection ignores any chatId: scope comes only from channelIdentity
    expect(out(hostedSrc(), { chatId: CID })).toBe("[image hostedContentId=abc]");
  });

  it("rejects malformed shapes, hosts, versions and byte paths other than the exact one", () => {
    const own = hostedSrc();
    for (const src of [
      own.replace("https://graph.microsoft.com", "https://graph.microsoft.com.evil.example"),
      own.replace("https://", "http://"),
      own.replace("/v1.0/", "/beta/"),
      own.replace("/$value", ""),
      own.replace("/$value", "/$value/x"),
      own.replace("/hostedContents/", "/hostedContents/x/hostedContents/"),
      own.replace("/teams/", "/teamz/"),
      own.replace("/channels/", "/channel/"),
      own.replace(`/messages/${MID}`, `/messages/${MID}/replies/${RID}`), // reply form needs the reply message
      own.replace("hostedContents/abc", "hostedContents/a%2fb"),
      own.replace("hostedContents/abc", "hostedContents/"),
      own.replace(`/messages/${MID}`, `/messages/${MID}x`),
      `https://evil.example/?u=${own}`,
      " " + own,
    ]) {
      expect(out(src), src).toBe("[image]");
    }
  });

  it("does not take an id from alt text or other attributes", () => {
    const body =
      `<img alt="hostedContentId=spoof" src="https://evil.example/x">` +
      `<img alt="[image hostedContentId=spoof2]" src="${hostedSrc(MID, OTHER_TID)}">` +
      `<img src="${hostedSrc(MID, TID, OTHER_CID)}" alt="hostedContentId=spoof3">`;
    const text = bodyOf(projectChannelMessage(msg({ body: htmlBody(body) }), 2000));
    expect(text).not.toMatch(/hostedContentId=spoof\b(?!:)/);
    expect(text).not.toMatch(/hostedContentId=spoof[23]\b(?!:)/);
    expect(text.match(/\[image hostedContentId=/g)).toBeNull();
  });

  it("accepts a reply's own image only when the reply and its parent both match", () => {
    const replyUrl = (parent: string, reply: string) =>
      `https://graph.microsoft.com/v1.0/teams/${TID}/channels/${CID}/messages/${parent}/replies/${reply}/hostedContents/rr/$value`;
    const reply = { id: RID, replyToId: MID };
    expect(out(replyUrl(MID, RID), reply)).toBe("[image hostedContentId=rr]");
    expect(out(replyUrl("1", RID), reply)).toBe("[image]");
    expect(out(replyUrl(MID, "1"), reply)).toBe("[image]");
    // a root message (no replyToId) never takes the reply form
    expect(out(replyUrl(MID, MID))).toBe("[image]");
    expect(out(replyUrl(MID, RID), { id: RID, replyToId: null })).toBe("[image]");
    // and a reply takes neither its parent's direct form nor its own: only the /replies/ form
    // pins both the reply and its parent
    expect(out(hostedSrc(MID), reply)).toBe("[image]");
    expect(out(hostedSrc(RID), reply)).toBe("[image]");
  });

  it("matches the team GUID case-insensitively but the opaque channel id exactly", () => {
    const url = (team: string, channel: string) =>
      `https://graph.microsoft.com/v1.0/teams/${team}/channels/${channel}/messages/${MID}/hostedContents/abc/$value`;
    expect(out(url(TID.toUpperCase(), CID))).toBe("[image hostedContentId=abc]");
    expect(out(url(TID, CID.toUpperCase()))).toBe("[image]");
  });

  it("yields undefined for non-objects and keeps the chat projection untouched", () => {
    expect(projectChannelMessage(null, 10)).toBeUndefined();
    expect(projectChannelMessage("x", 10)).toBeUndefined();
    expect(projectChannelMessage([], 10)).toBeUndefined();
  });
});

describe("pagination fails closed (same team, channel, message and collection)", () => {
  const lists = CASES.filter((c) => c.paged);
  it.each(lists)("$name refuses a foreign nextLink", async (c) => {
    const own = c.next;
    const cidPart = (cid: string) => `/teams/${TID}/channels/${cid}`;
    const bad = [
      // other team / channel
      `${own.replace(TID, OTHER_TID)}`,
      `${own.replace(CID, OTHER_CID)}`,
      `${own.replace(CID, enc(OTHER_CID))}`,
      // other message / collection / nested paths
      `${cidPart(CID)}/messages/1`,
      `${cidPart(CID)}/messages/${MID}`,
      `${cidPart(CID)}/messages/${MID}/replies/${RID}`,
      `${cidPart(CID)}/messages/${MID}/hostedContents/abc/$value`,
      `${cidPart(CID)}/messages/${MID}/replies/${RID}/hostedContents`,
      `${cidPart(CID)}/messages/delta()`,
      `${cidPart(CID)}/tabs`,
      `${cidPart(CID)}/filesFolder`,
      `${cidPart(CID)}/members`,
      `${cidPart(CID)}`,
      `/teams/${TID}/channels`,
      `/teams/${TID}`,
      `/teams/${TID}/channels/getAllMessages`,
      `/chats/${CID}/messages`,
      `/me/chats/${CID}/messages`,
      `${own}/%24value`,
      `${own}/$value`,
      `${own}%2f..%2f${enc(OTHER_CID)}`,
    ];
    for (const path of bad) {
      calls = [];
      responses.push(() =>
        json({
          value: [c.item],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0${path}?$skiptoken=1`,
        }),
      );
      const { env } = await call(c.name, { ...c.args, top: 5 });
      expect(env.ok, `${c.name} ${path}`).toBe(false);
      expect(env.error?.code, `${c.name} ${path}`).toBe("invalid_input");
      expect(calls, `${c.name} ${path}`).toHaveLength(1);
    }
    for (const link of [
      `https://evil.example/v1.0${own}?$skiptoken=1`,
      `https://graph.microsoft.com.evil.example/v1.0${own}`,
      `https://graph.microsoft.com/beta${own}?$skiptoken=1`,
      `http://graph.microsoft.com/v1.0${own}?$skiptoken=1`,
      `https://graph.microsoft.com:444/v1.0${own}?$skiptoken=1`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [c.item], "@odata.nextLink": link }));
      const { env } = await call(c.name, { ...c.args, top: 5 });
      expect(env.error?.code, link).toBe("invalid_input");
      expect(calls, link).toHaveLength(1);
    }
  });

  it.each(lists)(
    "$name follows Graph's documented nextLink forms for the same scope",
    async (c) => {
      for (const link of [
        `https://graph.microsoft.com/v1.0${c.next}?$skiptoken=abc`,
        `https://graph.microsoft.com/v1.0${c.next.replace(CID, ECID)}?$top=2&$skiptoken=a%3d%3d`,
      ]) {
        calls = [];
        responses.push(() => json({ value: [c.item], "@odata.nextLink": link }));
        responses.push(() => json({ value: [c.item] }));
        const { env } = await call(c.name, { ...c.args, top: 5 });
        expect(env.ok, `${c.name} ${link}`).toBe(true);
        expect(calls).toHaveLength(2);
        expect((env.data as unknown[]).length).toBe(2);
      }
    },
  );

  it.each(lists)("$name fails closed on a 200 page without a value array", async (c) => {
    for (const body of [{ error: "backend failure" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(c.name, c.args)).env.error?.code, JSON.stringify(body)).toBe(
        "graph_error",
      );
    }
  });
});

describe("path id injection (zero network calls)", () => {
  const badTeam = [
    "",
    "not-a-guid",
    `${TID}/channels`,
    `${TID}%2Fchannels`,
    "..",
    "%2e%2e",
    `${TID}?$top=1`,
    `${TID}#x`,
    " " + TID,
    `${TID}\n`,
    "me",
  ];
  const badChannel = [
    "",
    `${CID}/messages`,
    "19%3Aabc%40thread.tacv2%2Fmessages",
    `${CID}%2F..%2F..`,
    "..",
    "$ref",
    "getAllMessages",
    `${CID}?$top=1`,
    `${CID}#frag`,
    ` ${CID}`,
    `19:${"a".repeat(600)}@thread.tacv2`,
    TID,
  ];
  const badMessage = [
    "",
    "abc",
    `${MID}/replies`,
    `${MID}%2Freplies`,
    `${MID}/hostedContents`,
    `${MID}?$expand=replies`,
    `${MID}#x`,
    "..",
    "%2e%2e",
    "$value",
    "$ref",
    "delta",
    "$count",
    "-1",
    "1.5",
    "1e5",
    ` ${MID}`,
    `${MID}\n`,
    "1".repeat(21),
    "replies",
    "hostedContents",
  ];
  const needsMsg = (n: string) => n !== LIST;

  it.each(CASES.map((c) => c.name))("%s rejects malformed team and channel ids", async (name) => {
    for (const [key, list] of [
      ["team-id", badTeam],
      ["channel-id", badChannel],
    ] as const) {
      for (const v of list) {
        calls = [];
        const args: Record<string, unknown> = { ...ids, [key]: v };
        if (needsMsg(name)) args["chatMessage-id"] = MID;
        const { env, text } = await call(name, args);
        expect(env.ok, `${name} ${key} ${JSON.stringify(v)}`).toBe(false);
        expect(env.error?.code, `${name} ${key} ${JSON.stringify(v)}`).toBe("invalid_input");
        expect(calls, `${name} ${key} ${JSON.stringify(v)}`).toHaveLength(0);
        expect(text).not.toContain(TOKEN);
      }
    }
  });

  it.each([GET, REPLIES, HOSTED])(
    "%s rejects malformed or route-changing message ids",
    async (name) => {
      for (const id of badMessage) {
        calls = [];
        const { env } = await call(name, { ...ids, "chatMessage-id": id });
        expect(env.ok, `${name} ${JSON.stringify(id)}`).toBe(false);
        expect(env.error?.code, `${name} ${JSON.stringify(id)}`).toBe("invalid_input");
        expect(calls, `${name} ${JSON.stringify(id)}`).toHaveLength(0);
      }
      for (const bad of [undefined, null, 5, {}, [MID]]) {
        calls = [];
        const { env } = await call(name, { ...ids, "chatMessage-id": bad });
        expect(env.ok).toBe(false);
        expect(calls).toHaveLength(0);
      }
    },
  );

  it("is unreachable for byte downloads and other routes (not in the catalog)", async () => {
    expect(catalog.some((e) => /\$value|download-bytes/.test(e.pathPattern + e.toolName))).toBe(
      false,
    );
    const tools = (await mcp.listTools()).tools.map((t) => t.name);
    expect(tools).not.toContain("download-bytes");
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

  it("registers all four tools with a batch policy", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    for (const n of [LIST, GET, REPLIES, HOSTED]) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
    }
  });

  it("projects and normalizes bodies exactly like the single path (list, get, replies, hosted)", async () => {
    const raw = msg({
      etag: "etag-secret",
      policyViolation: { x: "policy-secret" },
      body: htmlBody(
        `<div>Hi&nbsp;<script>alert(1)</script><img src="${hostedSrc()}"><img src="${hostedSrc(MID, OTHER_TID)}"></div>`,
      ),
      attachments: [{ id: "a", name: "n", contentUrl: "https://x.example/url-secret" }],
      replies: [{ id: "r-secret" }],
    });
    responses.push(() => json({ value: [raw] }));
    const single = (await call(LIST, ids)).env.data;
    responses.push(ok({ value: [raw], "@odata.context": "ctx-secret" }));
    const l = await batch(`${BASE}/messages?$top=20`.replace(CID, ECID));
    expect(l.env.ok).toBe(true);
    expect(batchBody(l)?.value).toEqual(single);
    expect(l.text).not.toMatch(/secret|alert\(1\)|<div>/);
    expect(JSON.stringify(single)).toContain("[image hostedContentId=abc]");

    for (const id of [ECID, CID]) {
      responses.push(ok(raw));
      const g = await batch(`/teams/${TID}/channels/${id}/messages/${MID}`);
      expect(g.env.ok, id).toBe(true);
      expect(batchBody(g)).toEqual(projectChannelMessage(raw, GET_BODY_MAX_CHARS));
      expect(g.text).not.toMatch(/secret|alert\(1\)|<div>/);
    }

    const reply = msg({ id: RID, replyToId: MID, etag: "etag-secret" });
    responses.push(() => json({ value: [reply] }));
    const sr = (await call(REPLIES, { ...ids, "chatMessage-id": MID })).env.data;
    responses.push(ok({ value: [reply] }));
    const r = await batch(`/teams/${TID}/channels/${ECID}/messages/${MID}/replies`);
    expect(batchBody(r)?.value).toEqual(sr);
    expect(r.text).not.toMatch(/secret/);

    const hosted = [
      { id: "h1", contentType: "image/png", contentBytes: "BYTES-secret" },
      { id: "h2" },
    ];
    responses.push(() => json({ value: hosted }));
    const sh = (await call(HOSTED, { ...ids, "chatMessage-id": MID })).env.data;
    responses.push(ok({ value: hosted }));
    const h = await batch(`/teams/${TID}/channels/${ECID}/messages/${MID}/hostedContents`);
    expect(batchBody(h)?.value).toEqual(sh);
    expect(sh).toEqual([{ id: "h1", contentType: "image/png" }, { id: "h2" }]);
    expect(h.text).not.toMatch(/secret/);
  });

  it("applies the same list body cap in batch as on the single path", async () => {
    const raw = msg({ body: htmlBody(`<p>${"z".repeat(5000)}</p>`) });
    responses.push(ok({ value: [raw] }));
    const l = await batch(`/teams/${TID}/channels/${ECID}/messages`);
    const body = (batchBody(l)?.value?.[0] as { body: { truncated?: boolean } }).body;
    expect(body.truncated).toBe(true);
  });

  it("returns the same page size on the batch path as on the single path", async () => {
    const many = Array.from({ length: 30 }, (_, i) => msg({ id: String(100 + i) }));
    responses.push(ok({ value: many }));
    const l = await batch(`/teams/${TID}/channels/${ECID}/messages?$top=3`);
    expect(batchBody(l)?.value).toHaveLength(3);
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const c = `/teams/${TID}/channels/${ECID}`;
    const m = `${c}/messages`;
    const bad: Array<[string, Record<string, string>?]> = [
      [`${m}?$top=51`],
      [`${m}?$top=0`],
      [`${m}?$top=5&$top=6`],
      [`${m}?$expand=replies`],
      [`${m}?$select=body`],
      [`${m}?$search=%22x%22`],
      [`${m}?$count=true`],
      [`${m}?$skip=1`],
      [`${m}?$orderby=${enc("createdDateTime desc")}`],
      [`${m}?$filter=${enc("lastModifiedDateTime gt 2021-01-01T00:00:00.000Z")}`],
      [m, { Prefer: 'outlook.body-content-type="html"' }],
      [m, { Prefer: "include-unknown-enum-members" }],
      [m, { ConsistencyLevel: "eventual" }],
      [`${m}/${MID}?$select=body`],
      [`${m}/${MID}?$expand=replies`],
      [`${m}/${MID}?$top=1`],
      [`${m}/${enc("1/replies")}`],
      [`${m}/abc`],
      [`${m}/delta()`],
      [`${m}/${MID}/replies?$expand=x`],
      [`${m}/${MID}/replies?$top=51`],
      [`${m}/${MID}/replies/${RID}`],
      [`${m}/${MID}/replies/${RID}/hostedContents`],
      [`${m}/${MID}/hostedContents?$top=5`],
      [`${m}/${MID}/hostedContents?$select=contentBytes`],
      [`${m}/${MID}/hostedContents/abc`],
      [`${m}/${MID}/hostedContents/abc/$value`],
      [`/teams/getAllMessages`],
      [`/teams/${TID}/channels/getAllMessages`],
      [`/teams/${TID}/channels/${enc(`${CID}/x`)}/messages`],
      [`/teams/not-a-guid/channels/${ECID}/messages`],
      [`/groups/${TID}/team/channels/${ECID}/messages`],
      [`/beta/teams/${TID}/channels/${ECID}/messages`],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });
});

describe("graph-batch: catalogued but unimplemented GET", () => {
  it("is refused with zero requests (synthetic entry; real catalog GETs are all implemented)", async () => {
    const probe = await probeUnimplementedGet(catalog);
    expect(probe.catalogued).toBe(true);
    expect(probe.hasHandler).toBe(false);
    expect(probe.batch.ok).toBe(false);
    expect(probe.batch.error?.code).toBe("invalid_input");
    expect(probe.requests).toBe(0);
  });
});

describe.each(CASES)("$name: capability and errors", (c) => {
  it("403 is a capability_gap with the no-escalation guidance after exactly one request", async () => {
    responses.push(() =>
      json({ error: { code: "Forbidden", message: `Missing scope. Bearer ${TOKEN}` } }, 403),
    );
    const { env, text } = await call(c.name, c.args);
    expect(env.ok).toBe(false);
    expect(env.data).toBeUndefined();
    expect(env.error?.code).toBe("capability_gap");
    expect(env.error?.message).toContain("do not broaden scopes automatically");
    expect(env.error?.details).toMatchObject({ status: 403, graphCode: "Forbidden" });
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
  });

  it("401 is reauth_required after one request and never mentions scopes", async () => {
    responses.push(() =>
      json({ error: { code: "InvalidAuthenticationToken", message: TOKEN } }, 401),
    );
    const { env, text } = await call(c.name, c.args);
    expect(env.error?.code).toBe("reauth_required");
    expect(env.error?.message).not.toMatch(/scope/i);
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
  });

  it("400 and 404 are graph_error", async () => {
    for (const status of [400, 404]) {
      responses.push(() => json({ error: { code: "NotFound", message: "m" } }, status));
      const { env } = await call(c.name, c.args);
      expect(env.error?.code, String(status)).toBe("graph_error");
    }
  });

  it.runIf(c.paged)(
    "a 403 or 401 on a followed page fails the call, not the first page",
    async () => {
      for (const [status, code] of [
        [403, "capability_gap"],
        [401, "reauth_required"],
      ] as const) {
        calls = [];
        responses.push(() =>
          json({
            value: [c.item],
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
    },
  );
});
