import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { projectChatMessage, projectHostedContent } from "./chat-messages.js";
import { createHandlers } from "./index.js";

/**
 * SM-11 part C: list-chat-message-hosted-contents (v1.0 GET, references only). There is no byte
 * download: no `/hostedContents/{id}` or `/$value` path is in the catalog.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const NAME = "list-chat-message-hosted-contents";
const CID = "19:2da4c29f6d7041eca70b638b43d45437@thread.v2";
const OTHER_CID = "19:ebe3857aa388434bab0cad9d2e09f4de@thread.v2";
const MID = "1615971548136";
const ECID = encodeURIComponent(CID);
const HID =
  "aWQ9eF8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNix0eXBlPTEsdXJsPWh0dHBzOi8vdXMtYXBpLmFzbS5za3lwZS5jb20vdjEvb2JqZWN0cy8wLXd1cy1kOS1lNTRmNjM1NWYxYmJkNGQ3ZTNmNGJhZmU4NTI5MTBmNi92aWV3cy9pbWdv";
const ARGS = { "chat-id": CID, "chatMessage-id": MID };

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

/** What Graph documents: contentBytes and contentType are always null in the list. */
const hosted = (over: Record<string, unknown> = {}) => ({
  id: HID,
  contentBytes: null,
  contentType: null,
  ...over,
});

describe("catalog and surface", () => {
  it("pins the v1.0 GET entry; no hosted-content byte path exists in the catalog", () => {
    const e = catalog.find((x) => x.toolName === NAME);
    expect(e?.pathPattern).toBe("/chats/{chat-id}/messages/{chatMessage-id}/hostedContents");
    expect(e?.method).toBe("get");
    expect(e?.["apiVersion"]).toBeUndefined();
    // D23: the documented chat-path scope.
    expect(e?.["workScopes"]).toEqual(["Chat.Read"]);
    expect(catalog.some((x) => x.toolName === "download-bytes")).toBe(false);
    expect(catalog.some((x) => /\$value|hostedContents\/\{/.test(x.pathPattern))).toBe(false);
  });

  it("describes the reference-only contract (D24)", async () => {
    const tool = (await mcp.listTools()).tools.find((t) => t.name === NAME);
    const desc = tool?.description ?? "";
    const tip = String(catalog.find((x) => x.toolName === NAME)?.["llmTip"]);
    expect(desc).toMatch(/Reference enumeration only/);
    expect(desc).toMatch(/\[image hostedContentId=<id>\]/);
    expect(desc).toMatch(/cannot retrieve hosted-content bytes/);
    expect(desc).toMatch(/no Graph \$top or other OData query parameter is sent/);
    expect(desc).toMatch(/same chat, message and hosted-content collection/);
    expect(desc).toMatch(/never converted into an empty successful result/);
    expect(desc).toMatch(/Work\/school accounts only\. Requires Chat\.Read/);
    for (const text of [desc, tip]) {
      expect(text).not.toMatch(/download-bytes|\$value|base64|contentType|ChatMessage\.Read/);
    }
  });

  it("exposes only chat-id, chatMessage-id and a local top", async () => {
    const tool = (await mcp.listTools()).tools.find((t) => t.name === NAME);
    expect(Object.keys(tool?.inputSchema.properties ?? {}).sort()).toEqual([
      "chat-id",
      "chatMessage-id",
      "top",
    ]);
  });
});

describe(NAME, () => {
  it("requests the collection with NO query and no special headers, returns id references", async () => {
    responses.push(() =>
      json({ "@odata.context": "ctx-secret", "@odata.count": 1, value: [hosted()] }),
    );
    const { env, text } = await call(NAME, ARGS);
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe(`/v1.0/chats/${ECID}/messages/${MID}/hostedContents`);
    expect([...parsed().searchParams.keys()]).toEqual([]);
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(env.data).toEqual([{ id: HID }]);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
    expect(text).not.toContain("ctx-secret");
  });

  it("drops contentBytes, media links, annotations and unknown fields", async () => {
    responses.push(() =>
      json({
        value: [
          hosted({
            contentBytes: "AAAA-bytes-secret",
            contentType: "image/png",
            "@odata.mediaContentType": "image/png",
            "@odata.mediaReadLink": "https://media.example/read-secret",
            "@odata.type": "#microsoft.graph.chatMessageHostedContent",
            "@odata.etag": "etag-secret",
            "@microsoft.graph.temporaryId": "temp-secret",
            webUrl: "https://x.example/url-secret",
          }),
          hosted({ id: "2", contentType: { nested: "type-secret" } }),
          "not-an-object",
          null,
          ["array"],
        ],
      }),
    );
    const { env, text } = await call(NAME, ARGS);
    expect(env.data).toEqual([{ id: HID, contentType: "image/png" }, { id: "2" }]);
    expect(text).not.toMatch(/secret/);
    expect(text).not.toContain("contentBytes");
    expect(text).not.toContain("@odata");
  });

  it("treats top as a local limit, sends no $top and follows validated pages", async () => {
    const link = `https://graph.microsoft.com/v1.0/chats/${CID}/messages/${MID}/hostedContents?$skiptoken=abc`;
    responses.push(() => json({ value: [hosted({ id: "1" })], "@odata.nextLink": link }));
    responses.push(() => json({ value: [hosted({ id: "2" }), hosted({ id: "3" })] }));
    const { env } = await call(NAME, { ...ARGS, top: 2 });
    expect(calls).toHaveLength(2);
    expect(parsed(0).searchParams.has("$top")).toBe(false);
    expect(parsed(1).pathname).toBe(`/v1.0/chats/${CID}/messages/${MID}/hostedContents`);
    expect((env.data as unknown[]).length).toBe(2);
    expect(env.meta).toMatchObject({ count: 2, hasMore: true });
  });

  it("rejects options with zero network calls", async () => {
    for (const extra of [
      { top: 0 },
      { top: 101 },
      { top: 1.5 },
      { select: ["id"] },
      { expand: ["x"] },
      { filter: "x" },
      { $filter: "x" },
      { orderby: [{ property: "id", direction: "desc" }] },
      { skip: 1 },
      { count: true },
      { download: true },
      { target: `/chats/${CID}/messages/${MID}/hostedContents/${HID}/$value` },
      { id: HID },
      { headers: { Prefer: "x" } },
      { unknownKey: 1 },
    ]) {
      calls = [];
      const { env, text } = await call(NAME, { ...ARGS, ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("rejects route-changing chat and message ids with zero network calls", async () => {
    const badChat = [
      "19:abc@thread.v2/messages",
      "19:abc@thread.v2/messages/1/hostedContents",
      "19%3Aabc%40thread.v2%2Fmessages",
      "19:abc@thread.v2%2F..%2F..",
      "..",
      "%2e%2e",
      "19:abc@thread.v2?$top=1",
      "19:abc@thread.v2#frag",
      "$ref",
      " 19:abc@thread.v2",
      "f17e0000-0000-4000-8000-000000000010",
      "",
    ];
    const badMessage = [
      "",
      "abc",
      "1/hostedContents/x/$value",
      "1%2fhostedContents%2fx%2f%24value",
      `${MID}/hostedContents/${HID}`,
      `${MID}/hostedContents/${HID}/$value`,
      `${MID}/replies`,
      `${MID}?$expand=x`,
      `${MID}#x`,
      "..",
      "%2e%2e",
      "$value",
      "hostedContents",
      "-1",
      "1.5",
      `${MID}\n`,
      "1".repeat(21),
    ];
    const cases: Array<Record<string, unknown>> = [
      ...badChat.map((id) => ({ "chat-id": id, "chatMessage-id": MID })),
      ...badMessage.map((id) => ({ "chat-id": CID, "chatMessage-id": id })),
      { "chat-id": CID },
      { "chatMessage-id": MID },
      { "chat-id": 5, "chatMessage-id": MID },
      { "chat-id": CID, "chatMessage-id": Number(MID) },
      { "chat-id": CID, "chatMessage-id": [MID] },
    ];
    for (const args of cases) {
      calls = [];
      const { env, text } = await call(NAME, args);
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code, JSON.stringify(args)).toBe("invalid_input");
      expect(calls, JSON.stringify(args)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "backend failure" }, { value: null }, { value: {} }, []]) {
      responses.push(() => json(body));
      expect((await call(NAME, ARGS)).env.error?.code, JSON.stringify(body)).toBe("graph_error");
    }
  });

  it("refuses a nextLink for another chat, message, collection, path, host or version", async () => {
    const own = `/chats/${ECID}/messages/${MID}/hostedContents`;
    for (const bad of [
      `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(OTHER_CID)}/messages/${MID}/hostedContents?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages/1/hostedContents?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0${own}/${HID}`,
      `https://graph.microsoft.com/v1.0${own}/${HID}/$value`,
      `https://graph.microsoft.com/v1.0${own}/$value`,
      `https://graph.microsoft.com/v1.0${own}/$count`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages/${MID}`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages/${MID}/replies`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/pinnedMessages`,
      `https://graph.microsoft.com/v1.0/users/f17e0000-0000-4000-8000-000000000010${own}`,
      `https://graph.microsoft.com/v1.0/me${own}`,
      `https://graph.microsoft.com/v1.0/chats/${ECID}/messages/${MID}%2fx/hostedContents`,
      `https://evil.example/v1.0${own}?$skiptoken=1`,
      `https://graph.microsoft.com.evil.example/v1.0${own}`,
      `https://graph.microsoft.com/beta${own}?$skiptoken=1`,
      `http://graph.microsoft.com/v1.0${own}?$skiptoken=1`,
      `https://graph.microsoft.com:444/v1.0${own}?$skiptoken=1`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [hosted()], "@odata.nextLink": bad }));
      const { env } = await call(NAME, { ...ARGS, top: 5 });
      expect(env.ok, bad).toBe(false);
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("maps 401, 403, 404 and 400 to the shared error conventions", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
      [400, "graph_error"],
    ] as const) {
      calls = [];
      responses.push(() => json({ error: { code: "x", message: `m ${TOKEN}` } }, status));
      const { env, text } = await call(NAME, ARGS);
      expect(env.ok, String(status)).toBe(false);
      expect(env.data).toBeUndefined();
      expect(env.error?.code, String(status)).toBe(code);
      expect(calls, String(status)).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("403 is a capability_gap with no-escalation guidance, even without a JSON body", async () => {
    responses.push(() => new Response(`body-marker-xyz ${TOKEN}`, { status: 403 }));
    const { env, text } = await call(NAME, ARGS);
    expect(env.error?.code).toBe("capability_gap");
    expect(env.error?.message).toContain("do not broaden scopes automatically");
    expect(text).not.toContain(TOKEN);
    expect(text).not.toContain("body-marker-xyz");
  });

  it("a 401 or 403 on a followed page fails the call instead of returning the first page", async () => {
    for (const [status, code] of [
      [403, "capability_gap"],
      [401, "reauth_required"],
    ] as const) {
      calls = [];
      responses.push(() =>
        json({
          value: [{ id: "1" }],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0/chats/${ECID}/messages/${MID}/hostedContents?$skiptoken=abc`,
        }),
      );
      responses.push(() => json({ error: { code: "x", message: TOKEN } }, status));
      const { env, text } = await call(NAME, { ...ARGS, top: 50 });
      expect(env.ok).toBe(false);
      expect(env.data).toBeUndefined();
      expect(env.error?.code).toBe(code);
      expect(calls).toHaveLength(2);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe("projectHostedContent", () => {
  it("keeps only a string id and string contentType; non-objects and id-less items yield undefined", () => {
    expect(
      projectHostedContent(hosted({ contentBytes: "bytes", contentType: "image/png" })),
    ).toEqual({ id: HID, contentType: "image/png" });
    expect(projectHostedContent({ id: 5, contentType: ["x"], contentBytes: "b" })).toBeUndefined();
    expect(projectHostedContent({ contentType: "image/png" })).toBeUndefined();
    expect(projectHostedContent({ id: HID, contentType: ["x"] })).toEqual({ id: HID });
    for (const bad of [null, "x", 1, [], undefined]) {
      expect(projectHostedContent(bad)).toBeUndefined();
    }
  });

  it("lists ids in the same form htmlToText surfaces as [image hostedContentId=<id>]", () => {
    const src = `https://graph.microsoft.com/v1.0/chats/${CID}/messages/${MID}/hostedContents/${HID}/$value`;
    const message = {
      id: MID,
      chatId: CID,
      body: { contentType: "html", content: `<img src="${src}">` },
    };
    const body = (projectChatMessage(message, 1000) as { body: { content: string } }).body.content;
    const listed = projectHostedContent(hosted()) as { id: string };
    expect(body).toBe(`[image hostedContentId=${listed.id}]`);
  });
});

describe("graph-batch", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const ok = (body: unknown) => () => json({ responses: [{ id: "1", status: 200, body }] });
  const batchBody = (r: { env: Env }) =>
    (
      r.env.data as {
        responses: Array<{ body?: { value?: unknown[]; hasMore?: boolean } }>;
      }
    ).responses[0]?.body;
  const own = `/chats/${ECID}/messages/${MID}/hostedContents`;

  it("registers a batch policy", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    expect(createHandlers(graph, catalog).get(NAME)?.batchPolicy?.check).toBeTypeOf("function");
  });

  it("applies the same projection as the single path and never returns contentBytes", async () => {
    const raw = [
      hosted({
        contentBytes: "AAAA-bytes-secret",
        contentType: "image/png",
        "@odata.mediaReadLink": "https://media.example/read-secret",
      }),
      "not-an-object",
    ];
    responses.push(() => json({ value: raw }));
    const single = (await call(NAME, ARGS)).env.data;
    for (const url of [own, `/chats/${CID}/messages/${MID}/hostedContents`]) {
      responses.push(ok({ value: raw, "@odata.context": "ctx-secret" }));
      const r = await batch(url);
      expect(r.env.ok, url).toBe(true);
      expect(batchBody(r)?.value, url).toEqual(single);
      expect(r.text, url).not.toMatch(/secret|contentBytes|@odata/);
    }
  });

  it("fails closed on a batch body that is not a collection", async () => {
    for (const body of [{ value: "x" }, "x", null, [], { id: HID, contentBytes: "b-secret" }]) {
      responses.push(ok(body));
      const r = await batch(own);
      expect(r.text, JSON.stringify(body)).not.toMatch(/secret/);
    }
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const bad: Array<[string, Record<string, string>?]> = [
      [`${own}?$top=5`],
      [`${own}?$select=id`],
      [`${own}?$filter=x`],
      [`${own}?$expand=x`],
      [`${own}?$count=true`],
      [`${own}?$skip=1`],
      [own, { Prefer: "include-unknown-enum-members" }],
      [own, { ConsistencyLevel: "eventual" }],
      // bytes and single items are not reachable
      [`${own}/${HID}`],
      [`${own}/${HID}/$value`],
      [`${own}/${enc(HID)}/$value`],
      [`${own}/$value`],
      [`${own}/$count`],
      [`/chats/${ECID}/messages/${enc(`${MID}/hostedContents/${HID}/$value`)}`],
      [`/chats/${ECID}/messages/${enc(`${MID}/hostedContents`)}`],
      // routing
      [`/chats/${ECID}/messages/abc/hostedContents`],
      [`/chats/${enc(`${CID}/x`)}/messages/${MID}/hostedContents`],
      [`/me/chats/${ECID}/messages/${MID}/hostedContents`],
      [`/users/f17e0000-0000-4000-8000-000000000010/chats/${ECID}/messages/${MID}/hostedContents`],
      [`/beta/chats/${ECID}/messages/${MID}/hostedContents`],
      [`/chats/${ECID}/messages/${MID}/replies/1/hostedContents`],
      [`/teams/t1/channels/${ECID}/messages/${MID}/hostedContents`],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });

  it("returns the local-limit page size on the batch path", async () => {
    responses.push(ok({ value: Array.from({ length: 40 }, (_, i) => ({ id: String(i) })) }));
    const b = batchBody(await batch(own));
    expect(b?.value).toHaveLength(25);
    expect(b?.hasMore).toBe(true);
  });
});

const enc = encodeURIComponent;
