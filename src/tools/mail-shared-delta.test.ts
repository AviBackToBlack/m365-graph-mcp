import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const MSG_ID = "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAB=";
const FOLDER_ID =
  "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAFNAAA=";
const OTHER_FOLDER = "AAMkOTHERFOLDERoooooooooooooooooooooooooooooooooooooooooAAA=";
const UPN = "shared@contoso.com";
const GUID = "f17e0000-0000-4000-8000-000000000016";
const SKIP = "R0usmcCM996atia_sQz-4aBc9DeFgHiJkLmNoPqRsTuVwXyZ0123456789";
const DELTA_TOK = "ZmFrZS1kZWx0YS10b2tlbi1mb3ItdGVzdHM_0123456789abcdefABCDEF";

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
let handlers: ReturnType<typeof createHandlers>;

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
    tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
    catalog,
    fetch,
    sleep: () => Promise.resolve(),
  });
  handlers = createHandlers(graph, catalog);
  const server = createServer(catalog, handlers);
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

const SHARED_LIST = "list-shared-mailbox-messages";
const SHARED_FOLDER = "list-shared-mailbox-folder-messages";
const SHARED_GET = "get-shared-mailbox-message";
const DELTA = "list-mail-folder-messages-delta";

type SharedCase = [string, Record<string, unknown>];
const SHARED_CASES: SharedCase[] = [
  [SHARED_LIST, {}],
  [SHARED_FOLDER, { "mailFolder-id": "inbox" }],
  [SHARED_GET, { "message-id": MSG_ID }],
];

describe("catalog (permission assertions)", () => {
  const entry = (name: string) => catalog.find((e) => e.toolName === name);
  it("shared tools declare only the work-account Mail.Read.Shared scope", () => {
    for (const n of [SHARED_LIST, SHARED_FOLDER, SHARED_GET]) {
      expect(entry(n)?.["workScopes"], n).toEqual(["Mail.Read.Shared"]);
      expect(entry(n)?.["scopes"], n).toBeUndefined();
      expect(entry(n)?.method).toBe("get");
    }
  });
  it("the delta tool is the /me folder delta and needs Mail.Read", () => {
    expect(entry(DELTA)?.pathPattern).toBe("/me/mailFolders/{mailFolder-id}/messages/delta()");
    expect(entry(DELTA)?.["scopes"]).toEqual(["Mail.Read"]);
  });
});

describe("shared mailbox tools", () => {
  it("list-shared-mailbox-messages: path, defaults, UPN encoding", async () => {
    await call(SHARED_LIST, { "user-id": UPN });
    expect(parsed().pathname).toBe("/v1.0/users/shared%40contoso.com/messages");
    expect(parsed().searchParams.get("$top")).toBe("10");
    expect(parsed().searchParams.get("$select")).toBe(
      "id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments",
    );
    expect(calls[0]?.headers["Authorization"]).toBe(`Bearer ${TOKEN}`);
  });

  it("accepts a GUID user-id and follows a same-principal nextLink", async () => {
    responses.push(() =>
      json({
        value: [{ id: "1" }],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/users/${GUID}/messages?$skip=1`,
      }),
    );
    responses.push(() => json({ value: [{ id: "2" }] }));
    const { env } = await call(SHARED_LIST, { "user-id": GUID, top: 5 });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(2);
    expect(parsed(1).pathname).toBe(`/v1.0/users/${GUID}/messages`);
    expect(env.meta).toMatchObject({ count: 2, hasMore: false });
  });

  it("accepts OData key syntax in a same-principal nextLink", async () => {
    responses.push(() =>
      json({
        value: [{ id: "1" }],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/users('${GUID}')/messages?$skip=1`,
      }),
    );
    const { env } = await call(SHARED_LIST, { "user-id": GUID, top: 5 });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(2);
  });

  it("refuses a nextLink that switches mailbox (user-id is bound), incl. a UPN->GUID rewrite", async () => {
    for (const bad of [
      `https://graph.microsoft.com/v1.0/users/other%40contoso.com/messages?$skip=1`,
      `https://graph.microsoft.com/v1.0/me/messages?$skip=1`,
      // Pinned: a UPN rewritten to a GUID cannot be proven to be the same mailbox, so fail closed.
      `https://graph.microsoft.com/v1.0/users/${GUID}/messages?$skip=1`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": bad }));
      const { env, text } = await call(SHARED_LIST, { "user-id": UPN, top: 5 });
      expect(env.ok, bad).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("list-shared-mailbox-folder-messages: path and the shared SM-5 query rules", async () => {
    await call(SHARED_FOLDER, {
      "user-id": UPN,
      "mailFolder-id": "inbox",
      top: 100,
      search: "from:john",
      select: ["id", "subject"],
    });
    expect(parsed().pathname).toBe("/v1.0/users/shared%40contoso.com/mailFolders/inbox/messages");
    expect(parsed().searchParams.get("$top")).toBe("100");
    expect(parsed().searchParams.get("$search")).toBe('"from:john"');
    expect(parsed().searchParams.get("$select")).toBe("id,subject");
  });

  it("reuses the list limits and allowlists (no bodies, no attachments, no raw OData)", async () => {
    const cases: SharedCase[] = [
      [SHARED_LIST, { top: 1001 }],
      [SHARED_LIST, { select: ["body"] }],
      [SHARED_LIST, { select: ["contentBytes"] }],
      [SHARED_LIST, { expand: "attachments" }],
      [SHARED_LIST, { search: "x", orderby: [{ property: "subject" }] }],
      [SHARED_LIST, { search: "x", skip: 1 }],
      [SHARED_LIST, { filter: "a".repeat(1001) }],
      [SHARED_FOLDER, { "mailFolder-id": "inbox", url: "/me/messages" }],
      [SHARED_GET, { "message-id": MSG_ID, select: ["attachments"] }],
      [SHARED_GET, { "message-id": MSG_ID, expand: "attachments" }],
      [SHARED_GET, { "message-id": MSG_ID, bodyType: "unsafe" }],
    ];
    for (const [name, args] of cases) {
      const { env } = await call(name, { "user-id": UPN, ...args });
      expect(env.ok, `${name} ${JSON.stringify(args)}`).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("projects list items to the allowlist (bodies and attachments never leave)", async () => {
    responses.push(() =>
      json({
        value: [
          {
            id: "1",
            subject: "s",
            body: { content: "secret" },
            attachments: [{ contentBytes: "x" }],
          },
        ],
      }),
    );
    const { env, text } = await call(SHARED_LIST, { "user-id": UPN });
    expect(env.data).toEqual([{ id: "1", subject: "s" }]);
    expect(text).not.toContain("secret");
  });

  it("get-shared-mailbox-message: path, text body Prefer, no attachments in default select", async () => {
    responses.push(() => json({ id: "m", subject: "s" }));
    const { env } = await call(SHARED_GET, { "user-id": UPN, "message-id": MSG_ID });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/users/shared%40contoso.com/messages/${MSG_ID}`);
    expect(calls[0]?.headers["Prefer"]).toBe('outlook.body-content-type="text"');
    expect(parsed().searchParams.get("$select")).not.toContain("attachments");
    await call(SHARED_GET, { "user-id": UPN, "message-id": MSG_ID, bodyType: "html" });
    expect(calls[1]?.headers["Prefer"]).toBe('outlook.body-content-type="html"');
  });

  it("requires user-id (and the other path ids) with zero network calls", async () => {
    const cases: SharedCase[] = [
      ...SHARED_CASES,
      [SHARED_FOLDER, { "user-id": UPN }],
      [SHARED_GET, { "user-id": UPN }],
    ];
    for (const [name, args] of cases) {
      const { env } = await call(name, args);
      expect(env.ok, name).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("rejects hostile user-id values (incl. me) with zero network calls", async () => {
    const hostile = [
      "me", // alias of the personal-mailbox tools, not a principal; never a shared mailbox
      "ME",
      "Me",
      "delta",
      "$batch",
      "..",
      "../me",
      "a@b.com/../x",
      "a@b.com/messages",
      "shared%40contoso.com", // pre-encoded: would double-encode or smuggle %2F
      "a%2Fb@contoso.com",
      "a@b.com?$select=body",
      "a@b.com#x",
      "a@b.com\n",
      "a b@contoso.com",
      " ",
      "f17e0000-0000-4000-8000-000000000016/messages",
      "microsoft.graph.getByIds",
      "users('x')",
      "@contoso.com",
      "shared@contoso",
    ];
    for (const id of hostile) {
      for (const [name, extra] of SHARED_CASES) {
        const { env } = await call(name, { "user-id": id, ...extra });
        expect(env.ok, `${name} ${JSON.stringify(id)}`).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
      }
    }
    expect(calls).toHaveLength(0);
  });

  it("maps 403 from a shared mailbox to capability_gap (single call, no retry, no token)", async () => {
    for (const [name, args] of SHARED_CASES) {
      calls = [];
      responses.push(() =>
        json(
          { error: { code: "ErrorAccessDenied", message: `Access is denied. Bearer ${TOKEN}` } },
          403,
        ),
      );
      const { env, text } = await call(name, { "user-id": UPN, ...args });
      expect(env.ok).toBe(false);
      expect(env.error?.code).toBe("capability_gap");
      expect(env.error?.message).toContain("do not broaden scopes automatically");
      expect(env.error?.details).toMatchObject({ status: 403, graphCode: "ErrorAccessDenied" });
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("401 stays reauth_required and 404 stays a plain graph_error for shared tools", async () => {
    responses.push(() => json({ error: { code: "InvalidAuthenticationToken" } }, 401));
    expect((await call(SHARED_LIST, { "user-id": UPN })).env.error?.code).toBe("reauth_required");
    responses.push(() => json({ error: { code: "ErrorInvalidUser" } }, 404));
    expect((await call(SHARED_LIST, { "user-id": UPN })).env.error?.code).toBe("graph_error");
  });
});

describe("shared tools via graph-batch (same contract as SM-5)", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const PREFER = { Prefer: 'outlook.body-content-type="text"' };

  it("accepts in-policy subrequests", async () => {
    const ok: Array<[string, Record<string, string>?]> = [
      [`/users/${UPN}/messages?$select=id,subject&$top=250`],
      [`/users/${GUID}/messages?$select=id`],
      [`/users/${UPN}/mailFolders/inbox/messages?$select=id&$top=10`],
      [`/users/${UPN}/mailFolders/${FOLDER_ID}/messages?$select=id`],
      [`/users/${UPN}/messages/${MSG_ID}`, PREFER],
      [`/users/${UPN}/messages/${MSG_ID}?$select=id,body`, PREFER],
    ];
    for (const [url, headers] of ok) {
      calls = [];
      responses.push(() => json({ responses: [{ id: "1", status: 200, body: { value: [] } }] }));
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(true);
      expect(calls).toHaveLength(1);
    }
  });

  it("rejects out-of-policy and hostile subrequests with zero network calls", async () => {
    const bad: Array<[string, Record<string, string>?]> = [
      [`/users/${UPN}/messages`], // $select required in batch
      [`/users/${UPN}/messages?$select=id&$top=251`],
      [`/users/${UPN}/messages?$select=body`],
      [`/users/${UPN}/messages?$select=id&$expand=attachments`],
      [`/users/${UPN}/messages?$select=id&$skiptoken=abc`],
      [`/users/${UPN}/mailFolders/inbox/messages`],
      [`/users/${UPN}/messages/${MSG_ID}`], // Prefer required
      [`/users/${UPN}/messages/${MSG_ID}?$expand=attachments`, PREFER],
      [`/users/me/messages?$select=id`],
      [`/users/me/mailFolders/inbox/messages?$select=id`],
      [`/users/${UPN}/messages/delta?$select=id`],
      [`/users/${UPN}/mailFolders/inbox/messages/delta?$select=id`],
      [`/users/${UPN}/messages/${MSG_ID}/attachments?$select=id`],
      [`/users/${UPN}/messages/${MSG_ID}/$value`, PREFER],
      [`/users/a@b.com%2F..%2Fme/messages?$select=id`],
    ];
    for (const [url, headers] of bad) {
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(env.error?.code, url).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });
});

describe("list-mail-folder-messages-delta", () => {
  const nextLink = (tok = SKIP, folder = "inbox") =>
    `https://graph.microsoft.com/v1.0/me/mailFolders('${folder}')/messages/delta?$skiptoken=${tok}`;
  const deltaLink = (tok = DELTA_TOK, folder = "inbox") =>
    `https://graph.microsoft.com/v1.0/me/mailFolders('${folder}')/messages/delta?$deltatoken=${tok}`;
  const page =
    (value: unknown[], links: Record<string, string> = {}) =>
    () =>
      json({ "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#x", value, ...links });
  const inbox = { "mailFolder-id": "inbox" };

  describe("requests", () => {
    it("initial call: fixed path, default $select, page size via Prefer, no $top", async () => {
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      const { env } = await call(DELTA, inbox);
      expect(env.ok).toBe(true);
      expect(calls).toHaveLength(1);
      expect(parsed().pathname).toBe("/v1.0/me/mailFolders/inbox/messages/delta()");
      expect(parsed().searchParams.get("$select")).toBe(
        "id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments",
      );
      expect(parsed().searchParams.has("$top")).toBe(false);
      expect(parsed().searchParams.has("$skiptoken")).toBe(false);
      expect(calls[0]?.headers["Prefer"]).toBe("odata.maxpagesize=100");
      expect(calls[0]?.method).toBe("GET");
    });

    it("sends select, changeType and receivedAfter on the initial call; caps page size at 250", async () => {
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      await call(DELTA, {
        "mailFolder-id": FOLDER_ID,
        select: ["id", "subject"],
        changeType: "created",
        receivedAfter: "2024-01-31T00:00:00Z",
        top: 1000,
      });
      expect(parsed().pathname).toBe(`/v1.0/me/mailFolders/${FOLDER_ID}/messages/delta()`);
      expect(parsed().searchParams.get("$select")).toBe("id,subject");
      expect(parsed().searchParams.get("changeType")).toBe("created");
      expect(parsed().searchParams.get("$filter")).toBe("receivedDateTime ge 2024-01-31T00:00:00Z");
      expect(calls[0]?.headers["Prefer"]).toBe("odata.maxpagesize=250");
    });

    it("rejects invalid initial options without a network call", async () => {
      const bad: Array<Record<string, unknown>> = [
        { select: ["body"] },
        { select: ["uniqueBody"] },
        { select: [] },
        { changeType: "moved" },
        { receivedAfter: "2024-01-01" },
        { receivedAfter: "2024-13-45T00:00:00Z" },
        { receivedAfter: "2024-02-30T00:00:00Z" }, // Date.parse would roll this to March 1
        { receivedAfter: "2023-02-29T00:00:00Z" },
        { receivedAfter: "2024-04-31T00:00:00Z" },
        { receivedAfter: "2024-01-01T24:00:00Z" },
        { receivedAfter: "2024-01-01T00:00:00Z' or 1 eq 1" },
        { top: 0 },
        { top: 1001 },
        { top: 1.5 },
        { filter: "isRead eq false" },
        { search: "x" },
        { expand: "attachments" },
        { url: "https://graph.microsoft.com/v1.0/me/messages/delta" },
        { nextLink: "https://graph.microsoft.com/v1.0/x" },
        { deltaLink: "https://graph.microsoft.com/v1.0/x" },
      ];
      for (const args of bad) {
        const { env } = await call(DELTA, { ...inbox, ...args });
        expect(env.ok, JSON.stringify(args)).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
      }
      expect(calls).toHaveLength(0);
    });

    it("rejects hostile or reserved mailFolder-id values with zero network calls", async () => {
      for (const id of ["delta", "$count", "..", "inbox/messages", "a b", "inbox%2Fx", "x()", ""]) {
        const { env } = await call(DELTA, { "mailFolder-id": id });
        expect(env.ok, id).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
      }
      expect((await call(DELTA, {})).env.ok).toBe(false);
      expect(calls).toHaveLength(0);
    });
  });

  describe("rounds and continuation", () => {
    it("follows pages inside one call and returns deltaToken when the round completes", async () => {
      responses.push(page([{ id: "1" }, { id: "2" }], { "@odata.nextLink": nextLink() }));
      responses.push(page([{ id: "3" }], { "@odata.deltaLink": deltaLink() }));
      const { env, text } = await call(DELTA, { ...inbox, top: 500 });
      expect(env.ok).toBe(true);
      expect(calls).toHaveLength(2);
      expect((env.data as unknown[]).length).toBe(3);
      expect(env.meta).toMatchObject({
        count: 3,
        pages: 2,
        roundComplete: true,
        hasMore: false,
        deltaToken: DELTA_TOK,
      });
      expect(env.meta?.["skipToken"]).toBeUndefined();
      // The follow-up request is rebuilt: fixed path (OUR folder) plus only the skip token.
      expect(parsed(1).pathname).toBe("/v1.0/me/mailFolders/inbox/messages/delta()");
      expect([...parsed(1).searchParams.keys()]).toEqual(["$skiptoken"]);
      expect(parsed(1).searchParams.get("$skiptoken")).toBe(SKIP);
      expect(calls[1]?.headers["Prefer"]).toBe("odata.maxpagesize=250");
      expect(calls[1]?.headers["Authorization"]).toBe(`Bearer ${TOKEN}`);
      expect(text).not.toContain("graph.microsoft.com");
      expect(text).not.toMatch(/nextLink|deltaLink/);
      expect(text).not.toContain(TOKEN);
    });

    it("stops at top on a page boundary, returns skipToken, and never cuts a page", async () => {
      responses.push(
        page([{ id: "1" }, { id: "2" }, { id: "3" }], { "@odata.nextLink": nextLink() }),
      );
      const { env } = await call(DELTA, { ...inbox, top: 1 });
      expect(calls).toHaveLength(1);
      expect((env.data as unknown[]).length).toBe(3); // whole page kept: nothing is dropped
      expect(env.meta).toMatchObject({ roundComplete: false, hasMore: true, skipToken: SKIP });
      expect(env.meta?.["deltaToken"]).toBeUndefined();
      expect(calls[0]?.headers["Prefer"]).toBe("odata.maxpagesize=1");
    });

    it("stops at the page ceiling with a skipToken", async () => {
      let n = 0;
      for (let i = 0; i < 20; i += 1) {
        responses.push(() =>
          json({ value: [{ id: String((n += 1)) }], "@odata.nextLink": nextLink(`${SKIP}${n}`) }),
        );
      }
      const { env } = await call(DELTA, { ...inbox, top: 1000 });
      expect(calls).toHaveLength(10);
      expect(env.meta).toMatchObject({ hasMore: true, roundComplete: false, pages: 10 });
      expect(env.meta?.["skipToken"]).toBe(`${SKIP}10`);
    });

    it("resumes a round with skipToken: only that parameter, fixed path, deltaToken at the end", async () => {
      responses.push(page([{ id: "9" }], { "@odata.deltaLink": deltaLink() }));
      const { env } = await call(DELTA, { ...inbox, skipToken: SKIP });
      expect(env.meta).toMatchObject({ roundComplete: true, deltaToken: DELTA_TOK });
      expect(parsed().pathname).toBe("/v1.0/me/mailFolders/inbox/messages/delta()");
      expect([...parsed().searchParams.keys()]).toEqual(["$skiptoken"]);
      expect(parsed().searchParams.get("$skiptoken")).toBe(SKIP);
    });

    it("starts the next round with deltaToken: only that parameter", async () => {
      responses.push(
        page([{ id: "5", "@removed": { reason: "deleted" } }], {
          "@odata.deltaLink": deltaLink(`${DELTA_TOK}2`),
        }),
      );
      const { env } = await call(DELTA, { ...inbox, deltaToken: DELTA_TOK });
      expect([...parsed().searchParams.keys()]).toEqual(["$deltatoken"]);
      expect(parsed().searchParams.get("$deltatoken")).toBe(DELTA_TOK);
      expect(env.meta?.["deltaToken"]).toBe(`${DELTA_TOK}2`);
    });

    it("a returned token is accepted back as input (extract/validate symmetry)", async () => {
      responses.push(page([{ id: "1" }], { "@odata.nextLink": nextLink() }));
      const first = await call(DELTA, { ...inbox, top: 1 });
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      const second = await call(DELTA, { ...inbox, skipToken: first.env.meta?.["skipToken"] });
      expect(second.env.ok).toBe(true);
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      const third = await call(DELTA, { ...inbox, deltaToken: second.env.meta?.["deltaToken"] });
      expect(third.env.ok).toBe(true);
    });

    it("decodes percent-escapes and keeps a raw '+' when extracting, then re-encodes on send", async () => {
      responses.push(page([{ id: "1" }], { "@odata.nextLink": nextLink("ab%2Bcd%2Fef12%3Dghij") }));
      responses.push(page([], { "@odata.nextLink": nextLink("rawplus+token12345") }));
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      await call(DELTA, { ...inbox, top: 1000 });
      expect(calls[1]?.url).toContain("$skiptoken=ab%2Bcd%2Fef12%3Dghij");
      expect(calls[2]?.url).toContain("$skiptoken=rawplus%2Btoken12345");
    });

    it("never follows the folder in a link: a different folder id is ignored, ours is used", async () => {
      responses.push(page([{ id: "1" }], { "@odata.nextLink": nextLink(SKIP, OTHER_FOLDER) }));
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      const { env } = await call(DELTA, { "mailFolder-id": FOLDER_ID, top: 1000 });
      expect(env.ok).toBe(true);
      expect(parsed(1).pathname).toBe(`/v1.0/me/mailFolders/${FOLDER_ID}/messages/delta()`);
      expect(calls[1]?.url).not.toContain("OTHERFOLDER");
    });
  });

  describe("input validation (tokens)", () => {
    it("rejects both tokens, and a token combined with initial-only options", async () => {
      const bad: Array<Record<string, unknown>> = [
        { deltaToken: DELTA_TOK, skipToken: SKIP },
        { deltaToken: DELTA_TOK, select: ["id"] },
        { skipToken: SKIP, select: ["id"] },
        { deltaToken: DELTA_TOK, changeType: "created" },
        { skipToken: SKIP, receivedAfter: "2024-01-31T00:00:00Z" },
      ];
      for (const args of bad) {
        const { env } = await call(DELTA, { ...inbox, ...args });
        expect(env.ok, JSON.stringify(Object.keys(args))).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
      }
      expect(calls).toHaveLength(0);
    });

    it("rejects hostile, URL-shaped, short and overlong tokens without echoing them", async () => {
      const hostile = [
        "",
        "short",
        "https://graph.microsoft.com/v1.0/me/messages/delta?$deltatoken=abcdefgh",
        "abcdefgh&$select=body",
        "abcdefgh?x=1",
        "abcdefgh%2Fabc",
        "abcdefgh abc",
        "abcdefgh\nabc",
        "abcdefgh\u0000",
        "abcdefgh;DROP",
        "abcdefgh'abc",
        "abcdefgh<script>",
        "abcdefgh\\abc",
        "abcdefgh#frag",
        "a".repeat(4097),
      ];
      for (const kind of ["deltaToken", "skipToken"]) {
        for (const tok of hostile) {
          const { env, text } = await call(DELTA, { ...inbox, [kind]: tok });
          expect(env.ok, `${kind} ${JSON.stringify(tok.slice(0, 20))}`).toBe(false);
          expect(env.error?.code).toBe("invalid_input");
          if (tok.length > 8) expect(text).not.toContain(tok);
        }
      }
      expect(calls).toHaveLength(0);
    });
  });

  describe("forged or malicious Graph links", () => {
    const P = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta";
    const FORGED_NEXT = [
      `https://evil.example/v1.0/me/mailFolders/inbox/messages/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com.evil.example/v1.0/me/mailFolders/inbox/messages/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com@evil.example/v1.0/me/mailFolders/inbox/messages/delta?$skiptoken=${SKIP}`,
      `http://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com:8443/v1.0/me/mailFolders/inbox/messages/delta?$skiptoken=${SKIP}`,
      // other endpoint / other principal / other version
      `https://graph.microsoft.com/v1.0/me/messages/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/v1.0/me/calendar/events/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/v1.0/users/victim%40contoso.com/mailFolders/inbox/messages/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/beta/me/mailFolders/inbox/messages/delta?$skiptoken=${SKIP}`,
      `${P}/extra?$skiptoken=${SKIP}`,
      // extra / different query parameters
      `${P}?$skiptoken=${SKIP}&$select=body`,
      `${P}?$select=body&$skiptoken=${SKIP}`,
      `${P}?$skiptoken=${SKIP}&$skiptoken=${SKIP}2`,
      `${P}?$skiptoken=${SKIP}&$expand=attachments`,
      `${P}?$deltatoken=${DELTA_TOK}`, // wrong name for a nextLink
      `${P}?$skip=10`,
      P,
      `${P}?`,
      `${P}?$skiptoken=`,
      `${P}?skiptoken=${SKIP}`,
      `${P}?$skiptoken=${SKIP}#frag`,
      // hostile / overlong token values
      `${P}?$skiptoken=abc`,
      `${P}?$skiptoken=abcdefgh%00`,
      `${P}?$skiptoken=abcdefgh%26%24select%3Dbody`,
      `${P}?$skiptoken=${"a".repeat(4097)}`,
      `${P}?$skiptoken=%ZZabcdefgh`,
    ];

    it("rejects every forged nextLink with one request made and no raw link or token returned", async () => {
      for (const bad of FORGED_NEXT) {
        calls = [];
        responses.length = 0;
        responses.push(page([{ id: "1" }], { "@odata.nextLink": bad }));
        const { env, text } = await call(DELTA, { ...inbox, top: 1000 });
        expect(env.ok, bad.slice(0, 120)).toBe(false);
        expect(env.error?.code, bad.slice(0, 120)).toBe("invalid_input");
        expect(calls, bad.slice(0, 120)).toHaveLength(1);
        expect(text).not.toContain("evil.example");
        expect(text).not.toContain(TOKEN);
      }
    });

    it("rejects forged deltaLinks (wrong origin, user, extra query, wrong token name)", async () => {
      const bad = [
        `https://evil.example/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=${DELTA_TOK}`,
        `https://graph.microsoft.com/v1.0/users/victim%40contoso.com/mailFolders/inbox/messages/delta?$deltatoken=${DELTA_TOK}`,
        `${P}?$deltatoken=${DELTA_TOK}&$select=body`,
        `${P}?$skiptoken=${DELTA_TOK}`, // wrong name for a deltaLink
        `${P}?$deltatoken=latest`,
        `${P}?$deltatoken=${DELTA_TOK}&changeType=deleted`,
      ];
      for (const link of bad) {
        calls = [];
        responses.length = 0;
        responses.push(page([{ id: "1" }], { "@odata.deltaLink": link }));
        const { env, text } = await call(DELTA, inbox);
        expect(env.ok, link.slice(0, 120)).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
        expect(text).not.toContain("evil.example");
        expect(text).not.toContain("deltatoken");
      }
    });

    it("treats a non-string link value as no link at all (graph_error)", async () => {
      responses.push(() => json({ value: [], "@odata.nextLink": { href: "x" } }));
      const { env } = await call(DELTA, inbox);
      expect(env.ok).toBe(false);
      expect(env.error?.code).toBe("graph_error");
    });

    it("both links on one page, neither link, or a repeating skip token is a graph_error", async () => {
      responses.push(page([], { "@odata.nextLink": nextLink(), "@odata.deltaLink": deltaLink() }));
      expect((await call(DELTA, inbox)).env.error?.code).toBe("graph_error");
      responses.push(page([{ id: "1" }]));
      expect((await call(DELTA, inbox)).env.error?.code).toBe("graph_error");
      calls = [];
      responses.push(page([{ id: "1" }], { "@odata.nextLink": nextLink() }));
      responses.push(page([{ id: "2" }], { "@odata.nextLink": nextLink() }));
      const rep = await call(DELTA, { ...inbox, top: 1000 });
      expect(rep.env.error?.code).toBe("graph_error");
      expect(calls).toHaveLength(2);
    });

    it("never hands back the skip token it was resumed with, even when stopping at top", async () => {
      calls = [];
      responses.push(page([{ id: "1" }], { "@odata.nextLink": nextLink() }));
      const { env, text } = await call(DELTA, { ...inbox, skipToken: SKIP, top: 1 });
      expect(env.error?.code).toBe("graph_error");
      expect(text).not.toContain(SKIP);
      expect(calls).toHaveLength(1);
    });

    it("accepts a real leap-day receivedAfter", async () => {
      calls = [];
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      const { env } = await call(DELTA, { ...inbox, receivedAfter: "2024-02-29T00:00:00Z" });
      expect(env.ok).toBe(true);
    });
  });

  describe("results and errors", () => {
    it("keeps @removed (reduced to a known reason) and strips everything outside the allowlist", async () => {
      responses.push(
        page(
          [
            { id: "1", subject: "s", body: { content: "secret-body" }, "@odata.etag": "W/1" },
            { id: "2", "@removed": { reason: "deleted", extra: "x" } },
            { id: "3", "@removed": { reason: "changed" } },
          ],
          { "@odata.deltaLink": deltaLink() },
        ),
      );
      const { env, text } = await call(DELTA, inbox);
      expect(env.data).toEqual([
        { id: "1", subject: "s", "@odata.etag": "W/1" },
        { id: "2", "@removed": { reason: "deleted" } },
        { id: "3", "@removed": { reason: "changed" } },
      ]);
      expect(text).not.toContain("secret-body");
    });

    it("fails the whole page closed on any malformed change, never returning its token", async () => {
      const fresh = "ZnJlc2gtZGVsdGEtdG9rZW4tbmV2ZXItcmV0dXJuZWQ_0123456789";
      for (const bad of [
        "not-an-object",
        null,
        42,
        [],
        { subject: "no id" },
        { id: "" },
        { id: 7 },
        { id: "4", "@removed": { reason: "<script>" } },
        { id: "5", "@removed": "weird" },
        { id: "6", "@removed": {} },
        { id: "7", "@removed": null },
      ]) {
        calls = [];
        responses.length = 0;
        responses.push(page([{ id: "ok" }, bad], { "@odata.deltaLink": deltaLink(fresh) }));
        const { env, text } = await call(DELTA, inbox);
        expect(env.error?.code, JSON.stringify(bad)).toBe("graph_error");
        expect(text).not.toContain(fresh);
      }
    });

    it("projects resumed pages too, even if the token encoded a wider $select", async () => {
      responses.push(
        page([{ id: "1", body: { content: "secret-body" }, uniqueBody: "x" }], {
          "@odata.deltaLink": deltaLink(),
        }),
      );
      const { text } = await call(DELTA, { ...inbox, deltaToken: DELTA_TOK });
      expect(text).not.toContain("secret-body");
    });

    it("maps 410 Gone to restart_required", async () => {
      responses.push(() => json({ error: { code: "syncStateNotFound", message: "gone" } }, 410));
      const { env } = await call(DELTA, { ...inbox, deltaToken: DELTA_TOK });
      expect(env.ok).toBe(false);
      expect(env.error?.code).toBe("restart_required");
      expect(calls).toHaveLength(1);
    });

    it("maps a 410 on a mid-round page to restart_required too", async () => {
      responses.push(page([{ id: "1" }], { "@odata.nextLink": nextLink() }));
      responses.push(() => json({ error: { code: "syncStateNotFound" } }, 410));
      const { env } = await call(DELTA, { ...inbox, top: 1000 });
      expect(env.error?.code).toBe("restart_required");
    });

    it("maps a 4xx syncStateNotFound/resyncRequired error code to restart_required", async () => {
      for (const [status, code] of [
        [400, "syncStateNotFound"],
        [404, "SyncStateNotFound"],
        [400, "resyncRequired"],
      ] as const) {
        responses.push(() => json({ error: { code } }, status));
        const { env } = await call(DELTA, { ...inbox, skipToken: SKIP });
        expect(env.error?.code, `${status} ${code}`).toBe("restart_required");
      }
    });

    it("leaves other errors alone: plain 400 stays graph_error, a 5xx syncState is not a reset", async () => {
      responses.push(() => json({ error: { code: "BadRequest" } }, 400));
      expect((await call(DELTA, inbox)).env.error?.code).toBe("graph_error");
      responses.push(() => json({ error: { code: "syncStateNotFound" } }, 500));
      expect((await call(DELTA, inbox)).env.error?.code).toBe("graph_error");
    });

    it("maps 403 to capability_gap and 401 to reauth_required (no token echoed)", async () => {
      responses.push(() => json({ error: { code: "ErrorAccessDenied", message: TOKEN } }, 403));
      const gap = await call(DELTA, inbox);
      expect(gap.env.error?.code).toBe("capability_gap");
      expect(calls).toHaveLength(1);
      expect(gap.text).not.toContain(TOKEN);
      responses.push(() => json({ error: { code: "x" } }, 401));
      expect((await call(DELTA, inbox)).env.error?.code).toBe("reauth_required");
    });

    it("never returns a Graph URL or the access token", async () => {
      responses.push(
        page([{ id: "1", webLink: "https://outlook.office365.com/x" }], {
          "@odata.nextLink": nextLink(),
        }),
      );
      const { text } = await call(DELTA, { ...inbox, top: 1 });
      expect(text).not.toContain("graph.microsoft.com");
      expect(text).not.toContain("$skiptoken");
      expect(text).not.toContain(TOKEN);
    });
  });

  describe("batch", () => {
    it("has no batch policy and is unreachable through graph-batch", async () => {
      expect(handlers.get(DELTA)?.batchPolicy).toBeUndefined();
      for (const url of [
        "/me/mailFolders/inbox/messages/delta()?$select=id",
        "/me/mailFolders/inbox/messages/delta",
        "/me/mailFolders/inbox/messages/delta?$select=id&$deltatoken=abcdefgh",
        `/me/mailFolders/${FOLDER_ID}/messages/delta()?$skiptoken=${SKIP}`,
        "/users/a@b.com/mailFolders/inbox/messages/delta?$select=id",
      ]) {
        const { env } = await call("graph-batch", { requests: [{ id: "1", method: "GET", url }] });
        expect(env.ok, url).toBe(false);
        expect(env.error?.code, url).toBe("invalid_input");
      }
      expect(calls).toHaveLength(0);
    });
  });
});

describe("registration", () => {
  it("registers the shared and delta tools, with batch policies only on the shared ones", () => {
    const names = [...handlers.keys()];
    for (const n of [SHARED_LIST, SHARED_FOLDER, SHARED_GET]) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
      expect(names.indexOf(n)).toBeLessThan(names.indexOf("graph-batch"));
    }
    expect(handlers.has(DELTA)).toBe(true);
    expect(handlers.get(DELTA)?.batchPolicy).toBeUndefined();
    expect(names.indexOf(DELTA)).toBeLessThan(names.indexOf("graph-batch"));
  });

  it("does not expose url-like or link-like inputs on the delta tool", () => {
    const keys = Object.keys(handlers.get(DELTA)?.inputSchema ?? {});
    expect(keys.sort()).toEqual(
      ["changeType", "deltaToken", "receivedAfter", "select", "skipToken", "top"].sort(),
    );
  });
});
