import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { TokenError } from "../auth/token-provider.js";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { RESERVED_INPUT_NAMES } from "../registry.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";
import { probeUnimplementedGet } from "./unimplemented-get-fixture.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const MSG_ID = "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAB=";
const FOLDER_ID =
  "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAFNAAA=";

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
let tokenError: Error | undefined;

const json = (body: unknown, status = 200): Response =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });

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
    tokenProvider: {
      getAccessToken: () => (tokenError ? Promise.reject(tokenError) : Promise.resolve(TOKEN)),
    },
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
  tokenError = undefined;
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

const LIST_TOOLS = [
  { name: "list-mail-messages", path: "/v1.0/me/messages", params: {} },
  { name: "list-mail-folders", path: "/v1.0/me/mailFolders", params: {} },
  {
    name: "list-mail-child-folders",
    path: `/v1.0/me/mailFolders/${FOLDER_ID}/childFolders`,
    params: { "mailFolder-id": FOLDER_ID },
  },
  {
    name: "list-mail-folder-messages",
    path: `/v1.0/me/mailFolders/${FOLDER_ID}/messages`,
    params: { "mailFolder-id": FOLDER_ID },
  },
  {
    name: "list-mail-attachments",
    path: `/v1.0/me/messages/${MSG_ID}/attachments`,
    params: { "message-id": MSG_ID },
  },
] as const;

describe("mail tool schemas", () => {
  it("advertises typed inputs with allowlists and no proxy-style parameters", async () => {
    const tools = (await mcp.listTools()).tools;
    const by = (n: string) => tools.find((t) => t.name === n);
    for (const name of [...LIST_TOOLS.map((t) => t.name), "get-mail-message"]) {
      const props = Object.keys(by(name)?.inputSchema.properties ?? {});
      expect(props.length).toBeGreaterThan(0);
      for (const p of props) expect(RESERVED_INPUT_NAMES).not.toContain(p.toLowerCase());
      for (const p of props) expect(p.startsWith("$")).toBe(false);
    }
    const schemaOf = (n: string) => JSON.stringify(by(n)?.inputSchema);
    expect(schemaOf("list-mail-messages")).not.toContain('"body"');
    expect(schemaOf("list-mail-folder-messages")).not.toContain('"uniqueBody"');
    expect(schemaOf("get-mail-message")).toContain('"body"');
    expect(schemaOf("list-mail-attachments")).not.toContain("contentBytes");
    expect(schemaOf("list-mail-messages")).toContain('"maximum":1000');
    expect(schemaOf("list-mail-folders")).toContain('"maximum":250');
    expect(by("get-mail-message")?.inputSchema.required).toEqual(["message-id"]);
  });

  it("describes list-mail-attachments as metadata only, with no byte workflow (D27)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = tools.find((t) => t.name === "list-mail-attachments")?.description ?? "";
    const entry = (await loadEndpointCatalog()).find((e) => e.toolName === "list-mail-attachments");
    const tip = String(entry?.["llmTip"]);
    expect(desc).toMatch(/metadata only/);
    expect(desc).toMatch(/top is a typed, bounded result limit \(1-100, default 25\)/);
    expect(desc).toMatch(/id, name, contentType, size, isInline and lastModifiedDateTime/);
    expect(desc).toMatch(/also the default projection/);
    expect(desc).toMatch(/Attachment bytes and content are never returned/);
    expect(desc).toMatch(/Requires Mail\.Read/);
    expect(entry?.["scopes"]).toEqual(["Mail.Read"]);
    for (const text of [desc, tip]) {
      expect(text).not.toMatch(/download-bytes|\$value|contentBytes|contentLocation|base64/);
    }
  });

  it("keeps the untrusted-content notice on every mail tool", async () => {
    const tools = (await mcp.listTools()).tools;
    for (const name of [...LIST_TOOLS.map((t) => t.name), "get-mail-message"]) {
      expect(tools.find((t) => t.name === name)?.description).toContain("untrusted data");
    }
  });
});

describe("single-request tools", () => {
  it.each(LIST_TOOLS)("$name: fixed path, compact default query, no token in result", async (t) => {
    responses.push(() => json({ value: [{ id: "1", subject: "hi" }] }));
    const { r, env, text } = await call(t.name, { ...t.params });
    expect(r.isError).toBe(false);
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().origin).toBe("https://graph.microsoft.com");
    expect(parsed().pathname).toBe(t.path);
    expect(parsed().searchParams.get("$top")).not.toBeNull();
    expect(parsed().searchParams.get("$select")).not.toBeNull();
    expect(parsed().searchParams.get("$expand")).toBeNull();
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(text).not.toContain(TOKEN);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
  });

  it("list-mail-messages default select is compact: no body", async () => {
    await call("list-mail-messages");
    const select = parsed().searchParams.get("$select")?.split(",") ?? [];
    expect(select).toContain("bodyPreview");
    expect(select).not.toContain("body");
    expect(select).not.toContain("uniqueBody");
    expect(parsed().searchParams.get("$top")).toBe("10");
  });

  it("builds filter, orderby, skip, select and top as data", async () => {
    await call("list-mail-messages", {
      top: 250,
      select: ["id", "subject"],
      filter: "receivedDateTime ge 2024-01-01T00:00:00Z and subject eq 'a&b=c'",
      orderby: [{ property: "receivedDateTime", direction: "desc" }],
      skip: 20,
    });
    const q = parsed().searchParams;
    expect(q.get("$top")).toBe("250");
    expect(q.get("$select")).toBe("id,subject");
    expect(q.get("$filter")).toBe(
      "receivedDateTime ge 2024-01-01T00:00:00Z and subject eq 'a&b=c'",
    );
    expect(q.get("$orderby")).toBe("receivedDateTime desc");
    expect(q.get("$skip")).toBe("20");
    // The hostile-looking `&` stayed inside one encoded value.
    expect([...q.keys()].sort()).toEqual(["$filter", "$orderby", "$select", "$skip", "$top"]);
  });

  it("quotes $search as KQL and sends no ConsistencyLevel", async () => {
    await call("list-mail-messages", { search: 'from:john AND subject:"q3 plan"' });
    expect(parsed().searchParams.get("$search")).toBe('"from:john AND subject:\\"q3 plan\\""');
    expect(calls[0]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(parsed().searchParams.get("$orderby")).toBeNull();
    await call("list-mail-folder-messages", { "mailFolder-id": "inbox", search: "budget" });
    expect(parsed(1).searchParams.get("$search")).toBe('"budget"');
  });

  it("includeHiddenFolders is passed as a boolean on folder tools only", async () => {
    await call("list-mail-folders", { includeHiddenFolders: true });
    expect(parsed().searchParams.get("includeHiddenFolders")).toBe("true");
    const bad = await call("list-mail-messages", { includeHiddenFolders: true });
    expect(bad.env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(1);
  });

  it("rejects out-of-contract arguments with invalid_input and zero network calls", async () => {
    const bad: Array<[string, Record<string, unknown>]> = [
      ["list-mail-messages", { top: 1001 }],
      ["list-mail-folder-messages", { "mailFolder-id": "inbox", top: 1001 }],
      ["list-mail-messages", { top: 0 }],
      [
        "list-mail-messages",
        { filter: "isRead eq false", orderby: [{ property: "receivedDateTime" }] },
      ],
      [
        "list-mail-messages",
        {
          filter: "subject eq 'receivedDateTime'",
          orderby: [{ property: "receivedDateTime", direction: "desc" }],
        },
      ],
      ["list-mail-messages", { select: ["body"] }],
      ["list-mail-messages", { select: ["contentBytes"] }],
      ["list-mail-messages", { orderby: [{ property: "body" }] }],
      ["list-mail-messages", { filter: "x".repeat(1001) }],
      ["list-mail-messages", { search: "x".repeat(501) }],
      ["list-mail-messages", { search: "x", orderby: [{ property: "subject" }] }],
      ["list-mail-messages", { search: "x", skip: 5 }],
      ["list-mail-messages", { search: "x", filter: "isRead eq true" }],
      ["list-mail-messages", { expand: "attachments" }],
      ["list-mail-messages", { url: "/me" }],
      ["list-mail-messages", { nextLink: "https://graph.microsoft.com/v1.0/me/messages" }],
      ["list-mail-folders", { top: 251 }],
      ["list-mail-folders", { search: "x" }],
      ["list-mail-attachments", { "message-id": MSG_ID, top: 101 }],
      ["list-mail-attachments", { "message-id": MSG_ID, select: ["contentBytes"] }],
      ["list-mail-attachments", { "message-id": MSG_ID, filter: "size gt 1" }],
      ["get-mail-message", { "message-id": MSG_ID, top: 5 }],
      ["get-mail-message", { "message-id": MSG_ID, bodyType: "markdown" }],
      ["get-mail-message", { "message-id": MSG_ID, select: ["internetMessageHeaders"] }],
    ];
    for (const [name, args] of bad) {
      const { env, text } = await call(name, args);
      expect(env.ok, `${name} ${JSON.stringify(args)}`).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });

  it("never echoes caller-chosen argument names", async () => {
    const { text } = await call("list-mail-messages", { "pasted-secret-key": 1 });
    expect(text).not.toContain("pasted-secret-key");
  });
});

describe("path parameters", () => {
  const hostile = [
    "",
    ".",
    "..",
    "a/b",
    "../me",
    "x%2Fy",
    "%2e%2e",
    "x y",
    "delta",
    "$batch",
    "$count",
    "messages/AAA=",
    "id()",
    "a:b",
    "inbox/childFolders",
    "AAA\\BBB",
    "microsoft.graph.foo",
    "?$expand=x",
    "#frag",
  ];

  it("rejects hostile mailFolder-id and message-id values before any request", async () => {
    for (const bad of hostile) {
      for (const [name, key] of [
        ["list-mail-child-folders", "mailFolder-id"],
        ["list-mail-folder-messages", "mailFolder-id"],
        ["get-mail-message", "message-id"],
        ["list-mail-attachments", "message-id"],
      ] as const) {
        const { env, text } = await call(name, { [key]: bad });
        expect(env.error?.code, `${name} ${JSON.stringify(bad)}`).toBe("invalid_input");
        expect(text).not.toContain(TOKEN);
      }
    }
    expect(calls).toHaveLength(0);
  });

  it("accepts well-known folder names case-insensitively and opaque ids with '='", async () => {
    for (const id of ["inbox", "Inbox", "sentitems", "deleteditems", "drafts", "archive"]) {
      const { env } = await call("list-mail-folder-messages", { "mailFolder-id": id });
      expect(env.ok, id).toBe(true);
    }
    expect(parsed(0).pathname).toBe("/v1.0/me/mailFolders/inbox/messages");
    calls = [];
    await call("get-mail-message", { "message-id": MSG_ID });
    expect(parsed().pathname).toBe(`/v1.0/me/messages/${MSG_ID}`);
    // "=" stays literal (Softeria #245: %3D in Exchange ids caused 404s).
    expect(calls[0]?.url).not.toContain("%3D");
  });

  it("encodes reserved characters in an otherwise valid id as one segment", async () => {
    // "+" and "-" and "_" are valid Exchange id characters; "+" must be encoded.
    await call("get-mail-message", { "message-id": "AAMk+AG_x-y=" });
    expect(parsed().pathname).toBe("/v1.0/me/messages/AAMk%2BAG_x-y=");
  });
});

describe("get-mail-message", () => {
  it("returns the body, defaults to text, honours html, and keeps the select allowlist", async () => {
    responses.push(() =>
      json({ id: "m", subject: "s", body: { contentType: "text", content: "hello" } }),
    );
    const a = await call("get-mail-message", { "message-id": MSG_ID });
    expect(a.env.ok).toBe(true);
    expect((a.env.data as { body: { content: string } }).body.content).toBe("hello");
    expect(calls[0]?.headers["Prefer"]).toBe('outlook.body-content-type="text"');
    expect(parsed().searchParams.get("$select")?.split(",")).toContain("body");
    expect(parsed().searchParams.get("$expand")).toBeNull();
    expect(a.text).not.toContain(TOKEN);

    await call("get-mail-message", {
      "message-id": MSG_ID,
      bodyType: "html",
      select: ["id", "uniqueBody"],
    });
    expect(calls[1]?.headers["Prefer"]).toBe('outlook.body-content-type="html"');
    expect(parsed(1).searchParams.get("$select")).toBe("id,uniqueBody");
  });

  it("maps Graph failures to envelope errors without leaking the token", async () => {
    responses.push(() =>
      json({ error: { code: "ErrorItemNotFound", message: `nope ${TOKEN}` } }, 404),
    );
    const a = await call("get-mail-message", { "message-id": MSG_ID });
    expect(a.env.error?.code).toBe("graph_error");
    expect(a.text).not.toContain(TOKEN);
    responses.push(() => json({ error: { code: "x", message: "y" } }, 401));
    expect((await call("get-mail-message", { "message-id": MSG_ID })).env.error?.code).toBe(
      "reauth_required",
    );
    tokenError = new TokenError("reauth_required", "expired");
    expect((await call("get-mail-message", { "message-id": MSG_ID })).env.error?.code).toBe(
      "reauth_required",
    );
  });
});

describe("list-mail-attachments", () => {
  it("requests metadata only and never returns contentBytes or other content", async () => {
    responses.push(() =>
      json({
        value: [
          {
            "@odata.type": "#microsoft.graph.fileAttachment",
            id: "a1",
            name: "report.pdf",
            contentType: "application/pdf",
            size: 1234,
            isInline: false,
            lastModifiedDateTime: "2024-01-01T00:00:00Z",
            contentBytes: "QUJDRA==",
            contentId: "cid",
            contentLocation: "http://evil",
          },
          {
            "@odata.type": "#microsoft.graph.itemAttachment",
            id: "a2",
            name: "fwd",
            size: 9,
            item: { body: "x" },
          },
          {
            "@odata.type": "#microsoft.graph.referenceAttachment",
            id: "a3",
            name: "link",
            size: 1,
            sourceUrl: "https://x",
          },
          "not-an-object",
        ],
      }),
    );
    const { env, text } = await call("list-mail-attachments", { "message-id": MSG_ID });
    expect(env.ok).toBe(true);
    const select = parsed().searchParams.get("$select")?.split(",") ?? [];
    expect(select).toEqual([
      "id",
      "name",
      "contentType",
      "size",
      "isInline",
      "lastModifiedDateTime",
    ]);
    expect(text).not.toContain("contentBytes");
    expect(text).not.toContain("QUJDRA");
    expect(text).not.toContain("evil");
    expect(text).not.toContain("sourceUrl");
    expect(text).not.toContain('"item"');
    const data = env.data as Array<Record<string, unknown>>;
    expect(data).toHaveLength(3);
    expect(data[0]).toEqual({
      "@odata.type": "#microsoft.graph.fileAttachment",
      id: "a1",
      name: "report.pdf",
      contentType: "application/pdf",
      size: 1234,
      isInline: false,
      lastModifiedDateTime: "2024-01-01T00:00:00Z",
    });
    expect(data[1]).toEqual({
      "@odata.type": "#microsoft.graph.itemAttachment",
      id: "a2",
      name: "fwd",
      size: 9,
    });
  });

  it("strips contentBytes even on later pages whose link dropped $select", async () => {
    responses.push(() =>
      json({
        value: [{ id: "a1", name: "x" }],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/me/messages/${MSG_ID}/attachments?$skip=1`,
      }),
    );
    responses.push(() => json({ value: [{ id: "a2", name: "y", contentBytes: "SECRETBYTES" }] }));
    const { env, text } = await call("list-mail-attachments", { "message-id": MSG_ID, top: 5 });
    expect((env.data as unknown[]).length).toBe(2);
    expect(text).not.toContain("SECRETBYTES");
  });
});

describe("list projection (single request)", () => {
  it.each(LIST_TOOLS)(
    "$name: drops unselected properties even when a page ignored select",
    async (t) => {
      responses.push(() =>
        json({
          value: [
            {
              id: "1",
              "@odata.etag": "W/1",
              "@odata.nextLink": "x",
              body: { content: "SECRETBODY" },
              uniqueBody: { content: "SECRETUNIQUE" },
              contentBytes: "SECRETBYTES",
              unknown: "SECRETX",
            },
          ],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
        }),
      );
      // A page whose link dropped $select (second page of a path the first request shares).
      const { env, text } = await call(t.name, { ...t.params, top: 1 });
      expect(env.ok).toBe(true);
      expect(text).not.toContain("SECRET");
      expect(text).not.toContain("nextLink");
      expect(env.data).toEqual([{ id: "1", "@odata.etag": "W/1" }]);
    },
  );

  it("projects items of later pages whose link dropped $select", async () => {
    responses.push(() =>
      json({
        value: [{ id: "1", subject: "a" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
      }),
    );
    responses.push(() =>
      json({ value: [{ id: "2", subject: "b", body: { content: "SECRETBODY" } }] }),
    );
    const { env, text } = await call("list-mail-messages", { top: 5 });
    expect((env.data as unknown[]).length).toBe(2);
    expect(text).not.toContain("SECRETBODY");
  });
});

describe("pagination", () => {
  const link = (skip: number) =>
    `https://graph.microsoft.com/v1.0/me/messages?$top=2&$skip=${skip}`;

  it("follows nextLink within the limit via the bound client and sets no hasMore at the end", async () => {
    responses.push(() => json({ value: [{ id: "1" }, { id: "2" }], "@odata.nextLink": link(2) }));
    responses.push(() => json({ value: [{ id: "3" }] }));
    const { env, text } = await call("list-mail-messages", { top: 10, select: ["id"] });
    expect(calls).toHaveLength(2);
    expect(parsed(1).pathname).toBe("/v1.0/me/messages");
    expect((env.data as unknown[]).length).toBe(3);
    expect(env.meta).toMatchObject({ count: 3, hasMore: false, pages: 2 });
    expect(text).not.toContain("nextLink");
  });

  it("stops at top, reports hasMore, and never returns a Graph link", async () => {
    responses.push(() => json({ value: [{ id: "1" }, { id: "2" }], "@odata.nextLink": link(2) }));
    const { env, text } = await call("list-mail-messages", { top: 2 });
    expect(calls).toHaveLength(1);
    expect((env.data as unknown[]).length).toBe(2);
    expect(env.meta?.["hasMore"]).toBe(true);
    expect(text).not.toContain("graph.microsoft.com");
    expect(text).not.toContain("nextLink");
    expect(text).not.toContain("$skip");
  });

  it("clips an oversized page to top and flags hasMore", async () => {
    responses.push(() => json({ value: [{ id: "1" }, { id: "2" }, { id: "3" }] }));
    const { env } = await call("list-mail-messages", { top: 2 });
    expect((env.data as unknown[]).length).toBe(2);
    expect(env.meta?.["hasMore"]).toBe(true);
  });

  it("stops at the page ceiling and reports hasMore", async () => {
    let n = 0;
    for (let i = 0; i < 20; i += 1) {
      responses.push(() => json({ value: [{ id: String((n += 1)) }], "@odata.nextLink": link(n) }));
    }
    const { env } = await call("list-mail-messages", { top: 250 });
    expect(calls.length).toBeLessThanOrEqual(10);
    expect(env.meta?.["hasMore"]).toBe(true);
  });

  it("reaches up to 1000 results: $top stays <= 250 per request and pages are followed", async () => {
    let n = 0;
    const page = (more: boolean) => () =>
      json({
        value: Array.from({ length: 250 }, () => ({ id: String((n += 1)) })),
        ...(more ? { "@odata.nextLink": link(n) } : {}),
      });
    responses.push(page(true), page(true), page(true), page(true));
    const { env } = await call("list-mail-messages", { top: 1000, search: "budget" });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(4);
    expect(parsed(0).searchParams.get("$top")).toBe("250");
    expect(parsed(0).searchParams.get("$search")).toBe('"budget"');
    expect((env.data as unknown[]).length).toBe(1000);
    expect(env.meta).toMatchObject({ count: 1000, hasMore: true, pages: 4 });
  });

  it("sends a small $top unchanged and clamps nothing below the page size", async () => {
    await call("list-mail-messages", { top: 100 });
    expect(parsed().searchParams.get("$top")).toBe("100");
  });

  it("refuses a nextLink that leaves the endpoint, origin or principal (fail closed)", async () => {
    for (const bad of [
      "https://evil.example/v1.0/me/messages?$skip=1",
      "https://graph.microsoft.com.evil.example/v1.0/me/messages",
      "https://graph.microsoft.com/v1.0/users/x/messages",
      "https://graph.microsoft.com/v1.0/me/events",
      "http://graph.microsoft.com/v1.0/me/messages?$skip=1",
    ]) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": bad }));
      const { env, text } = await call("list-mail-messages", { top: 5 });
      expect(env.ok).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("keeps the Prefer/ConsistencyLevel-free header set across pages", async () => {
    responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": link(1) }));
    responses.push(() => json({ value: [] }));
    await call("list-mail-messages", { top: 5, search: "x" });
    expect(calls[1]?.headers["ConsistencyLevel"]).toBeUndefined();
    expect(calls[1]?.headers["Authorization"]).toBe(`Bearer ${TOKEN}`);
  });
});

describe("graph-batch reachability (query policy)", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const sel = "$select=id,subject";

  const VALID = [
    `/me/messages?${sel}&$top=250`,
    `/me/messages?${sel}&$search=%22from%3Ajohn%22`,
    `/me/messages?${sel}&$filter=receivedDateTime%20ge%202024-01-01T00:00:00Z%20and%20isRead%20eq%20false&$orderby=receivedDateTime%20desc&$skip=5`,
    `/me/mailFolders`,
    `/me/mailFolders?$top=250&includeHiddenFolders=true`,
    `/me/mailFolders/inbox/childFolders?$select=id,displayName`,
    `/me/mailFolders/${FOLDER_ID}/childFolders`,
    `/me/mailFolders/inbox/messages?${sel}&$top=10`,
    `/me/messages/${MSG_ID}/attachments?$select=id,name,size&$top=100`,
  ];

  it("accepts in-policy subrequests for every mail tool", async () => {
    for (const url of VALID) {
      calls = [];
      responses.push(() => json({ responses: [{ id: "1", status: 200, body: { value: [] } }] }));
      const { env } = await batch(url);
      expect(env.ok, url).toBe(true);
      expect(calls).toHaveLength(1);
      expect(calls[0]?.url).toContain("/$batch");
      const sent = JSON.parse(calls[0]?.body ?? "{}") as { requests: Array<{ method: string }> };
      expect(sent.requests[0]?.method).toBe("GET");
    }
  });

  it("rejects subrequests that violate the same limits, with zero network calls", async () => {
    const bad = [
      `/me/messages`, // $select required (batch cannot inject the compact default)
      `/me/messages?$top=10`,
      `/me/messages?${sel}&$top=251`,
      `/me/messages?${sel}&$top=1000`,
      `/me/messages?${sel}&$filter=isRead%20eq%20false&$orderby=receivedDateTime%20desc`,
      `/me/messages?${sel}&$filter=subject%20eq%20'x'%20and%20receivedDateTime%20ge%202024-01-01T00:00:00Z&$orderby=receivedDateTime,subject`,
      `/me/messages/${MSG_ID}`, // Prefer required: batch cannot inject the text default
      `/me/messages/${MSG_ID}?$select=id,body`,
      `/me/messages?${sel}&$top=05`,
      `/me/messages?${sel}&$select=id`,
      `/me/messages?$select=id,body`,
      `/me/messages?$select=id,contentBytes`,
      `/me/messages?${sel}&$expand=attachments`,
      `/me/messages?${sel}&expand=attachments`,
      `/me/messages?${sel}&top=5`,
      `/me/messages?${sel}&$TOP=5`,
      `/me/messages?${sel}&$skiptoken=abc`,
      `/me/messages?${sel}&$count=true`,
      `/me/messages?${sel}&$format=json`,
      `/me/messages?${sel}&$search=from%3Ajohn`, // unquoted KQL
      `/me/messages?${sel}&$search=%22a%22%20OR%20%22b%22`,
      `/me/messages?${sel}&$search=%22x%22&$orderby=subject`,
      `/me/messages?${sel}&$search=%22x%22&$skip=1`,
      `/me/messages?${sel}&$search=%22x%22&$filter=isRead%20eq%20true`,
      `/me/messages?${sel}&$orderby=body`,
      `/me/messages?${sel}&$filter=${"a".repeat(1001)}`,
      `/me/mailFolders?$top=251`,
      `/me/mailFolders?search=x`,
      `/me/mailFolders/inbox/messages`,
      `/me/mailFolders/inbox/messages?$select=id&$expand=attachments`,
      `/me/mailFolders/delta/messages?${sel}`,
      `/me/mailFolders/$count/childFolders`,
      `/me/messages/${MSG_ID}?$top=1`,
      `/me/messages/${MSG_ID}?$expand=attachments`,
      `/me/messages/${MSG_ID}?$select=internetMessageHeaders`,
      `/me/messages/${MSG_ID}/attachments`, // would embed contentBytes
      `/me/messages/${MSG_ID}/attachments?$select=id,contentBytes`,
      `/me/messages/${MSG_ID}/attachments?$select=id&$top=101`,
      `/me/messages/${MSG_ID}/attachments?$select=id&$expand=microsoft.graph.itemattachment/item`,
      `/me/messages/${MSG_ID}/attachments/a1?$select=id`, // not a catalogued path
      `/me/messages/${MSG_ID}/attachments/a1/$value`,
      `/me/messages/${MSG_ID}/$value`,
      `/me/messages/delta`,
      `/me/messages/..%2Fevents`,
    ];
    for (const url of bad) {
      const { env, text } = await batch(url);
      expect(env.ok, url).toBe(false);
      expect(env.error?.code, url).toBe("invalid_input");
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });

  it("strips contentBytes from batch attachment sub-responses even if Graph ignored $select", async () => {
    responses.push(() =>
      json({
        responses: [
          {
            id: "1",
            status: 200,
            body: {
              "@odata.context": "ctx",
              value: [
                {
                  "@odata.type": "#microsoft.graph.fileAttachment",
                  id: "a1",
                  name: "n",
                  size: 3,
                  contentBytes: "SECRETBYTES",
                },
                "junk",
              ],
            },
          },
          { id: "2", status: 200, body: "not-json-base64" },
          { id: "3", status: 200, body: { value: "SECRETBYTES-shape" } },
        ],
      }),
    );
    const { env, text } = await call("graph-batch", {
      requests: [
        { id: "1", method: "GET", url: `/me/messages/${MSG_ID}/attachments?$select=id,name,size` },
        { id: "2", method: "GET", url: `/me/messages/${MSG_ID}/attachments?$select=id` },
        { id: "3", method: "GET", url: `/me/messages/${MSG_ID}/attachments?$select=id&$top=1` },
      ],
    });
    expect(env.ok).toBe(true);
    expect(text).not.toContain("SECRETBYTES");
    expect(text).not.toContain("contentBytes");
    const out = (env.data as { responses: Array<{ body?: { value?: unknown[] } }> }).responses;
    expect(out[0]?.body?.value).toEqual([
      { "@odata.type": "#microsoft.graph.fileAttachment", id: "a1", name: "n", size: 3 },
    ]);
    expect(out[2]?.body).toEqual({});
  });

  it("clips batch collections to $top (or the default) and drops continuation links", async () => {
    const att = (i: number) => ({ id: `a${i}`, name: "n", size: 1 });
    const msg = (i: number) => ({ id: `m${i}`, subject: "s" });
    responses.push(() =>
      json({
        responses: [
          {
            id: "1",
            status: 200,
            body: { value: [1, 2, 3].map(att), "@odata.nextLink": "https://graph.microsoft.com/x" },
          },
          { id: "2", status: 200, body: { value: Array.from({ length: 101 }, (_, i) => att(i)) } },
          { id: "3", status: 200, body: { value: [1, 2, 3].map(msg) } },
        ],
      }),
    );
    const { env, text } = await call("graph-batch", {
      requests: [
        { id: "1", method: "GET", url: `/me/messages/${MSG_ID}/attachments?$select=id&$top=1` },
        { id: "2", method: "GET", url: `/me/messages/${MSG_ID}/attachments?$select=id` },
        { id: "3", method: "GET", url: "/me/messages?$select=id,subject&$top=2" },
      ],
    });
    expect(env.ok).toBe(true);
    expect(text).not.toContain("nextLink");
    const out = (
      env.data as { responses: Array<{ body?: { value: unknown[]; hasMore?: boolean } }> }
    ).responses;
    expect(out[0]?.body?.value).toHaveLength(1);
    expect(out[0]?.body?.hasMore).toBe(true);
    expect(out[1]?.body?.value).toHaveLength(25); // attachment default, not all 101
    expect(out[1]?.body?.hasMore).toBe(true);
    expect(out[2]?.body?.value).toHaveLength(2);
  });

  it("rejects the beta-only mailFolder sizeInBytes property (v1.0 tools)", async () => {
    const { env } = await call("list-mail-folders", { select: ["id", "sizeInBytes"] });
    expect(env.ok).toBe(false);
    expect(calls).toHaveLength(0);
  });

  it("accepts only the canonical body-content-type Prefer on get-mail-message", async () => {
    const getUrls = [`/me/messages/${MSG_ID}`, `/me/messages/${MSG_ID}?$select=id,body`];
    for (const type of ["text", "html"]) {
      calls = [];
      responses.push(() => json({ responses: [{ id: "1", status: 200, body: { id: "m" } }] }));
      const { env } = await batch(getUrls[type === "text" ? 0 : 1] as string, {
        Prefer: `outlook.body-content-type="${type}"`,
      });
      expect(env.ok, type).toBe(true);
      expect(calls).toHaveLength(1);
    }
    calls = [];
    const bad: Array<[string, Record<string, string>]> = [
      [`/me/messages/${MSG_ID}`, { Prefer: "outlook.allow-unsafe-html" }],
      [`/me/messages/${MSG_ID}`, { Prefer: 'outlook.timezone="UTC"' }],
      [`/me/messages/${MSG_ID}`, { Prefer: 'outlook.body-content-type="markdown"' }],
      [`/me/messages/${MSG_ID}`, { Prefer: "outlook.body-content-type=text" }],
      [
        `/me/messages/${MSG_ID}`,
        { Prefer: 'outlook.body-content-type="text", outlook.allow-unsafe-html' },
      ],
      [`/me/messages/${MSG_ID}`, { Prefer: 'outlook.body-content-type="TEXT"' }],
      [`/me/messages/${MSG_ID}`, { ConsistencyLevel: "eventual" }],
      [`/me/messages?${sel}`, { Prefer: 'outlook.timezone="UTC"' }],
      [`/me/messages?${sel}`, { Prefer: 'outlook.body-content-type="text"' }],
      [`/me/messages?${sel}`, { ConsistencyLevel: "eventual" }],
      [`/me/mailFolders`, { ConsistencyLevel: "eventual" }],
      [`/me/mailFolders/inbox/childFolders`, { Prefer: 'outlook.timezone="UTC"' }],
      [`/me/mailFolders/inbox/messages?${sel}`, { ConsistencyLevel: "eventual" }],
      [`/me/messages/${MSG_ID}/attachments?$select=id`, { ConsistencyLevel: "eventual" }],
      [`/me/messages/${MSG_ID}/attachments?$select=id`, { Prefer: "outlook.allow-unsafe-html" }],
    ];
    for (const [url, headers] of bad) {
      const { env, text } = await batch(url, headers);
      expect(env.error?.code, `${url} ${JSON.stringify(headers)}`).toBe("invalid_input");
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });

  it("accepts a valid filter+orderby pair and a message list at the page size", async () => {
    responses.push(() => json({ responses: [{ id: "1", status: 200, body: { value: [] } }] }));
    const { env } = await batch(
      `/me/messages?${sel}&$top=250&$filter=receivedDateTime%20ge%202024-01-01T00:00:00Z%20and%20isRead%20eq%20false&$orderby=receivedDateTime%20desc`,
    );
    expect(env.ok).toBe(true);
  });

  it("projects message, folder and attachment sub-responses to the allowlist and drops collection metadata", async () => {
    const msg = {
      id: "m1",
      subject: "s",
      "@odata.etag": "W/1",
      body: { content: "SECRETBODY" },
      uniqueBody: { content: "SECRETUNIQUE" },
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
    };
    const collection = (item: unknown) => ({
      "@odata.context": "ctx",
      "@odata.nextLink": "https://evil.example/next",
      unknownKey: "SECRETMETA",
      value: [item],
    });
    responses.push(() =>
      json({
        responses: [
          { id: "1", status: 200, body: collection(msg) },
          {
            id: "2",
            status: 200,
            body: collection({ id: "f", displayName: "d", extra: "SECRETX" }),
          },
          { id: "3", status: 200, body: collection({ id: "a", contentBytes: "SECRETBYTES" }) },
          { id: "4", status: 200, body: collection({ id: "m", body: { content: "SECRETBODY2" } }) },
        ],
      }),
    );
    const { env, text } = await call("graph-batch", {
      requests: [
        { id: "1", method: "GET", url: `/me/messages?${sel}` },
        { id: "2", method: "GET", url: `/me/mailFolders` },
        { id: "3", method: "GET", url: `/me/messages/${MSG_ID}/attachments?$select=id` },
        { id: "4", method: "GET", url: `/me/mailFolders/inbox/messages?${sel}` },
      ],
    });
    expect(env.ok).toBe(true);
    for (const secret of ["SECRET", "evil.example", "nextLink", "@odata.context", "unknownKey"]) {
      expect(text).not.toContain(secret);
    }
    const out = (env.data as { responses: Array<{ body?: unknown }> }).responses;
    expect(out[0]?.body).toEqual({
      value: [{ id: "m1", subject: "s", "@odata.etag": "W/1" }],
      hasMore: true,
    });
    expect(out[1]?.body).toEqual({ value: [{ id: "f", displayName: "d" }], hasMore: true });
    expect(out[2]?.body).toEqual({ value: [{ id: "a" }], hasMore: true });
    expect(out[3]?.body).toEqual({ value: [{ id: "m" }], hasMore: true });
  });

  it("does not project get-mail-message sub-responses (body is intended there)", async () => {
    responses.push(() =>
      json({ responses: [{ id: "1", status: 200, body: { id: "m", body: { content: "hi" } } }] }),
    );
    const { env } = await batch(`/me/messages/${MSG_ID}`, {
      Prefer: 'outlook.body-content-type="text"',
    });
    const out = (env.data as { responses: Array<{ body?: unknown }> }).responses;
    expect(out[0]?.body).toEqual({ id: "m", body: { content: "hi" } });
  });

  it("does not echo rejected parameter names or values", async () => {
    const { text } = await batch(`/me/messages?${sel}&secretparam=abc123`);
    expect(text).not.toContain("secretparam");
    expect(text).not.toContain("abc123");
  });

  it("still refuses a catalogued GET whose tool is not implemented (synthetic entry), zero requests (batch and direct)", async () => {
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

describe("registration", () => {
  it("registers exactly the mail tools (SM-5 plus SM-6 shared) with a batch policy, before graph-batch", () => {
    const handlers = createHandlers(
      new GraphClient({ tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) }, catalog }),
      catalog,
    );
    const names = [...handlers.keys()];
    const mail = [
      "list-mail-messages",
      "list-mail-folders",
      "list-mail-child-folders",
      "list-mail-folder-messages",
      "get-mail-message",
      "list-mail-attachments",
      "list-shared-mailbox-messages",
      "list-shared-mailbox-folder-messages",
      "get-shared-mailbox-message",
    ];
    for (const n of mail) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
      expect(names.indexOf(n)).toBeLessThan(names.indexOf("graph-batch"));
    }
    // The delta tool is registered too, but has no batch policy (see mail-shared-delta.test.ts).
    // Calendar tools (SM-7) are covered in calendar.test.ts, the availability POST tools (SM-8B)
    // in calendar-availability.test.ts; everything else is mail.
    // Identity tools (SM-9 part A) are covered in identity.test.ts.
    const notMail = new Set([
      "graph-batch",
      "get-schedule",
      "find-meeting-times",
      "get-current-user",
      "get-my-profile",
      "get-mailbox-settings",
      // Directory tools (SM-9 part B) are covered in directory.test.ts.
      "list-users",
      "get-my-manager",
      "list-my-direct-reports",
      "get-user-manager",
      "list-user-direct-reports",
      // People and memberships (SM-9 part C) are covered in people.test.ts.
      "list-relevant-people",
      "list-my-memberships",
      // Group directory tools (SM-10 part A) are covered in groups.test.ts.
      "list-groups",
      "get-group",
      "list-group-members",
      "list-group-owners",
      // Group calendar tools (SM-10 part B) are covered in group-calendar.test.ts.
      "list-group-events",
      // Legacy group conversations/threads (SM-10 part C) are covered in group-conversations.test.ts.
      "list-group-conversations",
      "list-group-threads",
      // Teams chats (SM-11 part A) are covered in chats.test.ts.
      "list-chats",
      "get-chat",
      "list-chat-members",
      // Teams chat messages (SM-11 part B) are covered in chat-messages.test.ts.
      "list-chat-messages",
      "get-chat-message",
      "list-pinned-chat-messages",
      // Teams and channels (SM-12 part A) are covered in teams.test.ts.
      "get-team",
      "list-team-channels",
      "get-team-channel",
      "list-team-members",
      // Teams channel messages (SM-12 part B) are covered in channel-messages.test.ts.
      "list-channel-messages",
      "get-channel-message",
      "list-channel-message-replies",
      "list-channel-message-hosted-contents",
      // Teams tabs and files folder (SM-12 part C) are covered in channel-tabs.test.ts.
      "list-channel-tabs",
      "get-channel-files-folder",
      // Teams chat hosted contents (SM-11 part C) are covered in chat-hosted-contents.test.ts.
      "list-chat-message-hosted-contents",
      // SharePoint sites (SM-13 part A) are covered in sharepoint-sites.test.ts.
      "search-sharepoint-sites",
      "get-sharepoint-site",
      "get-sharepoint-site-by-path",
      // SharePoint site drives (SM-13 part B) are covered in sharepoint-drives.test.ts.
      "list-sharepoint-site-drives",
      "get-sharepoint-site-drive-by-id",
      // SharePoint lists, items and columns (SM-13 part C) are covered in sharepoint-lists.test.ts.
      "list-sharepoint-site-lists",
      "get-sharepoint-site-list",
      "list-sharepoint-site-list-items",
      "get-sharepoint-site-list-item",
      "list-sharepoint-list-columns",
      "get-sharepoint-list-column",
      "get-my-presence",
      "get-user-presence",
      "list-trending-insights",
      // Microsoft Search query (SM-14 part B) is covered in search-query.test.ts.
      "search-query",
      // Copilot Retrieval (SM-14 part C) is covered in copilot-retrieve.test.ts.
      "copilot-retrieve",
    ]);
    expect(names.filter((n) => !notMail.has(n) && !/calendar/.test(n)).sort()).toEqual(
      [...mail, "list-mail-folder-messages-delta"].sort(),
    );
  });
});
