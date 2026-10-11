import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const GUID = "f17e0000-0000-4000-8000-000000000016";
const UPN = "shared@contoso.com";
const SKIP = "R0usmcCM996atia_sQz-4aBc9DeFgHiJkLmNoPqRsTuVwXyZ0123456789";
const DELTA_TOK = "ZmFrZS1kZWx0YS10b2tlbi1mb3ItdGVzdHM_0123456789abcdefABCDEF";
const START = "2024-03-01T00:00:00Z";
const END = "2024-03-08T00:00:00Z";
const WIN = { startDateTime: START, endDateTime: END };

interface Call {
  url: string;
  method: string;
  headers: Record<string, string>;
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

const SHARED_EVENTS = "list-shared-calendar-events";
const SHARED_VIEW = "get-shared-calendar-view";
const VIEW_DELTA = "list-calendar-view-delta";
const EVENTS_DELTA = "list-calendar-events-delta";

type Case = [string, Record<string, unknown>];
const SHARED_CASES: Case[] = [
  [SHARED_EVENTS, {}],
  [SHARED_VIEW, { ...WIN }],
];

describe("catalog (permission assertions)", () => {
  const entry = (name: string) => catalog.find((e) => e.toolName === name);
  it("shared tools declare only the work-account Calendars.Read.Shared scope and no extended-property flag", () => {
    for (const n of [SHARED_EVENTS, SHARED_VIEW]) {
      expect(entry(n)?.["workScopes"], n).toEqual(["Calendars.Read.Shared"]);
      expect(entry(n)?.["supportsExpandExtendedProperties"], n).toBeUndefined();
      expect(entry(n)?.["supportsTimezone"], n).toBe(true);
      expect(entry(n)?.method).toBe("get");
    }
    expect(entry(SHARED_EVENTS)?.pathPattern).toBe("/users/{user-id}/calendar/events");
    expect(entry(SHARED_VIEW)?.pathPattern).toBe("/users/{user-id}/calendarView");
  });
  it("the calendarView delta is the /me v1.0 function; the events delta is beta-only (D2)", () => {
    expect(entry(VIEW_DELTA)?.pathPattern).toBe("/me/calendarView/delta()");
    expect(entry(VIEW_DELTA)?.["apiVersion"]).toBeUndefined();
    expect(entry(VIEW_DELTA)?.["supportsTimezone"]).toBe(true);
    // Microsoft documents /me/events/delta in beta only; covered in calendar-events-delta.test.ts.
    expect(entry(EVENTS_DELTA)?.pathPattern).toBe("/me/events/delta()");
    expect(entry(EVENTS_DELTA)?.["apiVersion"]).toBe("beta");
    expect(handlers.has(EVENTS_DELTA)).toBe(true);
  });
});

describe("shared calendar tools", () => {
  it("list-shared-calendar-events: path, SM-7 list defaults, UPN encoding, no extendedPropertyId", async () => {
    const { env } = await call(SHARED_EVENTS, { "user-id": UPN });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe("/v1.0/users/shared%40contoso.com/calendar/events");
    expect(parsed().searchParams.get("$top")).toBe("25");
    expect(parsed().searchParams.get("$select")).toContain("recurrence");
    expect(parsed().searchParams.get("$select")).not.toContain("body");
    expect(calls[0]?.headers["Prefer"]).toBeUndefined();
    expect(calls[0]?.headers["Authorization"]).toBe(`Bearer ${TOKEN}`);
    const keys = Object.keys(handlers.get(SHARED_EVENTS)?.inputSchema ?? {});
    expect(keys).toContain("timeZone");
    expect(keys).not.toContain("extendedPropertyId");
    const { env: bad } = await call(SHARED_EVENTS, {
      "user-id": UPN,
      extendedPropertyId: "String {8ECCC264-6880-4EBE-992F-8888D2EEAA1D} Name TestProperty",
    });
    expect(bad.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(1);
  });

  it("get-shared-calendar-view: path, canonical UTC window from offsets, timeZone Prefer", async () => {
    await call(SHARED_VIEW, {
      "user-id": GUID,
      startDateTime: "2024-03-01T09:00:00+01:00",
      endDateTime: "2024-03-08T00:00:00Z",
      timeZone: "Europe/Berlin",
    });
    expect(parsed().pathname).toBe(`/v1.0/users/${GUID}/calendarView`);
    expect(parsed().searchParams.get("startDateTime")).toBe("2024-03-01T08:00:00Z");
    expect(parsed().searchParams.get("endDateTime")).toBe("2024-03-08T00:00:00Z");
    expect(calls[0]?.headers["Prefer"]).toBe('outlook.timezone="Europe/Berlin"');
    expect(Object.keys(handlers.get(SHARED_VIEW)?.inputSchema ?? {})).not.toContain(
      "extendedPropertyId",
    );
  });

  it("reuses the SM-7 window contract (required, ordered, bounded) with zero calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      {},
      { startDateTime: START },
      { startDateTime: END, endDateTime: START },
      { startDateTime: "2024-01-01T00:00:00", endDateTime: END },
      { startDateTime: "2023-01-01T00:00:00Z", endDateTime: "2024-03-01T00:00:00Z" }, // > 366 days
      { ...WIN, timeZone: "Not/AZone" },
      { ...WIN, $select: "body" },
      { ...WIN, select: ["body"] },
    ];
    for (const args of bad) {
      const { env } = await call(SHARED_VIEW, { "user-id": UPN, ...args });
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("follows a same-principal nextLink and refuses one that switches user", async () => {
    responses.push(() =>
      json({
        value: [{ id: "1", subject: "a", body: { content: "SECRET" } }],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/users/${UPN}/calendar/events?$skiptoken=abc`,
      }),
    );
    responses.push(() => json({ value: [{ id: "2" }] }));
    const ok = await call(SHARED_EVENTS, { "user-id": UPN, top: 50 });
    expect(ok.env.ok).toBe(true);
    expect((ok.env.data as unknown[]).length).toBe(2);
    expect(ok.text).not.toContain("SECRET");
    calls = [];
    responses.push(() =>
      json({
        value: [{ id: "1" }],
        "@odata.nextLink":
          "https://graph.microsoft.com/v1.0/users/victim@contoso.com/calendar/events?$skiptoken=abc",
      }),
    );
    const bad = await call(SHARED_EVENTS, { "user-id": UPN, top: 50 });
    expect(bad.env.ok).toBe(false);
    expect(bad.env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(1);
  });

  it("requires user-id with zero network calls", async () => {
    for (const [name, args] of SHARED_CASES) {
      const { env } = await call(name, args);
      expect(env.ok, name).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("rejects hostile user-id values (incl. me) with zero network calls", async () => {
    const hostile = [
      "me",
      "ME",
      "delta",
      "$batch",
      "..",
      "../me",
      "a@b.com/../x",
      "a@b.com/calendar/events",
      "shared%40contoso.com",
      "a%2Fb@contoso.com",
      "a@b.com?$select=body",
      "a@b.com#x",
      "a@b.com\n",
      "a b@contoso.com",
      " ",
      `${GUID}/events`,
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

  it("maps 403 to capability_gap (single call, no retry, no token)", async () => {
    for (const [name, args] of SHARED_CASES) {
      calls = [];
      responses.push(() =>
        json(
          { error: { code: "ErrorAccessDenied", message: `Access is denied. Bearer ${TOKEN}` } },
          403,
        ),
      );
      const { env, text } = await call(name, { "user-id": UPN, ...args });
      expect(env.error?.code).toBe("capability_gap");
      expect(env.error?.details).toMatchObject({ status: 403, graphCode: "ErrorAccessDenied" });
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe("shared calendar tools via graph-batch (SM-7 contract)", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const W = `startDateTime=${START}&endDateTime=${END}`;

  it("accepts in-policy subrequests", async () => {
    const ok: Array<[string, Record<string, string>?]> = [
      [`/users/${UPN}/calendar/events?$select=id,subject&$top=100`],
      [`/users/${GUID}/calendar/events?$select=id`, { Prefer: 'outlook.timezone="UTC"' }],
      [`/users/${UPN}/calendarView?$select=id&${W}`],
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
      [`/users/${UPN}/calendar/events`], // $select required in batch
      [`/users/${UPN}/calendar/events?$select=id&$top=101`],
      [`/users/${UPN}/calendar/events?$select=body`],
      [
        `/users/${UPN}/calendar/events?$select=id&$expand=singleValueExtendedProperties($filter=id eq 'String {8ECCC264-6880-4EBE-992F-8888D2EEAA1D} Name TestProperty')`,
      ],
      [`/users/${UPN}/calendarView?$select=id`], // window required
      [
        `/users/${UPN}/calendarView?$select=id&startDateTime=${START}&endDateTime=2024-03-08T01:00:00%2B01:00`,
      ],
      [`/users/me/calendar/events?$select=id`],
      [`/users/me/calendarView?$select=id&${W}`],
      [`/users/${UPN}/calendarView/delta?$select=id&${W}`],
      [`/users/${UPN}/events/delta?$select=id`],
      [`/users/${UPN}/calendar/events/delta?$select=id`],
      [`/users/${UPN}/calendar/events?$select=id`, { Prefer: "odata.maxpagesize=5" }],
      [`/users/a@b.com%2F..%2Fme/calendar/events?$select=id`],
    ];
    for (const [url, headers] of bad) {
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(env.error?.code, url).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });
});

describe(VIEW_DELTA, () => {
  const P = "https://graph.microsoft.com/v1.0/me/calendarView/delta";
  const nextLink = (tok = SKIP) => `${P}?$skiptoken=${tok}`;
  const deltaLink = (tok = DELTA_TOK) => `${P}?$deltatoken=${tok}`;
  const page =
    (value: unknown[], links: Record<string, string> = {}) =>
    () =>
      json({ "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#x", value, ...links });

  describe("initial call", () => {
    it("fixed path, canonical UTC window, page size via Prefer, no $select/$top", async () => {
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      const { env } = await call(VIEW_DELTA, {
        startDateTime: "2024-03-01T09:00:00+01:00",
        endDateTime: "2024-03-08T00:00:00.1230000Z",
      });
      expect(env.ok).toBe(true);
      expect(calls).toHaveLength(1);
      expect(parsed().pathname).toBe("/v1.0/me/calendarView/delta()");
      expect([...parsed().searchParams.keys()].sort()).toEqual(["endDateTime", "startDateTime"]);
      expect(parsed().searchParams.get("startDateTime")).toBe("2024-03-01T08:00:00Z");
      expect(parsed().searchParams.get("endDateTime")).toBe("2024-03-08T00:00:00.123Z");
      expect(calls[0]?.headers["Prefer"]).toBe("odata.maxpagesize=100");
      expect(calls[0]?.method).toBe("GET");
    });

    it("sends timeZone as outlook.timezone after the page size; caps page size at 100", async () => {
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      await call(VIEW_DELTA, { ...WIN, timeZone: "Pacific Standard Time", top: 1000 });
      expect(calls[0]?.headers["Prefer"]).toBe(
        'odata.maxpagesize=100, outlook.timezone="Pacific Standard Time"',
      );
    });

    it("requires the window on the initial call, validated like SM-7, with zero calls", async () => {
      const bad: Array<Record<string, unknown>> = [
        {},
        { startDateTime: START },
        { endDateTime: END },
        { startDateTime: END, endDateTime: START },
        { startDateTime: START, endDateTime: START },
        { startDateTime: "2024-01-01T00:00:00", endDateTime: END },
        { startDateTime: "tomorrow", endDateTime: END },
        { startDateTime: "2023-01-01T00:00:00Z", endDateTime: "2024-03-01T00:00:00Z" },
        { ...WIN, timeZone: "Not/AZone" },
        { ...WIN, timeZone: `UTC"; x=1` },
        { ...WIN, top: 0 },
        { ...WIN, top: 1001 },
        { ...WIN, select: ["id"] },
        { ...WIN, $select: "body" },
        { ...WIN, $filter: "x" },
        { ...WIN, changeType: "created" },
        { ...WIN, url: P },
      ];
      for (const args of bad) {
        const { env } = await call(VIEW_DELTA, args);
        expect(env.ok, JSON.stringify(args)).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
      }
      expect(calls).toHaveLength(0);
    });
  });

  describe("rounds and continuation", () => {
    it("follows pages in one call and returns the deltaToken when the round completes", async () => {
      responses.push(page([{ id: "1", subject: "a" }], { "@odata.nextLink": nextLink() }));
      responses.push(page([{ id: "2" }], { "@odata.deltaLink": deltaLink() }));
      const { env } = await call(VIEW_DELTA, { ...WIN, top: 1000 });
      expect(env.ok).toBe(true);
      expect((env.data as unknown[]).length).toBe(2);
      expect(env.meta).toMatchObject({
        roundComplete: true,
        hasMore: false,
        deltaToken: DELTA_TOK,
        pages: 2,
        count: 2,
      });
      expect(calls).toHaveLength(2);
      expect(parsed(1).pathname).toBe("/v1.0/me/calendarView/delta()");
      expect([...parsed(1).searchParams.keys()]).toEqual(["$skiptoken"]);
      expect(parsed(1).searchParams.get("$skiptoken")).toBe(SKIP);
      expect(calls[1]?.headers["Prefer"]).toBe(calls[0]?.headers["Prefer"]);
    });

    it("stops at top on a page boundary and returns a skipToken", async () => {
      responses.push(
        page([{ id: "1" }, { id: "2" }, { id: "3" }], { "@odata.nextLink": nextLink() }),
      );
      const { env } = await call(VIEW_DELTA, { ...WIN, top: 1 });
      expect((env.data as unknown[]).length).toBe(3);
      expect(env.meta).toMatchObject({ hasMore: true, roundComplete: false, skipToken: SKIP });
      expect(calls).toHaveLength(1);
    });

    it("resumes with skipToken / deltaToken: only that parameter, timeZone still sent", async () => {
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      await call(VIEW_DELTA, { skipToken: SKIP, timeZone: "UTC" });
      expect(parsed().pathname).toBe("/v1.0/me/calendarView/delta()");
      expect([...parsed().searchParams.keys()]).toEqual(["$skiptoken"]);
      expect(calls[0]?.headers["Prefer"]).toBe('odata.maxpagesize=100, outlook.timezone="UTC"');
      calls = [];
      responses.push(
        page([], { "@odata.deltaLink": deltaLink("bmV4dC1yb3VuZC10b2tlbi0xMjM0NTY") }),
      );
      const { env } = await call(VIEW_DELTA, { deltaToken: DELTA_TOK });
      expect([...parsed().searchParams.keys()]).toEqual(["$deltatoken"]);
      expect(parsed().searchParams.get("$deltatoken")).toBe(DELTA_TOK);
      expect(env.meta?.["deltaToken"]).toBe("bmV4dC1yb3VuZC10b2tlbi0xMjM0NTY");
    });

    it("rejects the window (and both tokens, and bad tokens) on a token call, zero calls", async () => {
      const bad: Array<Record<string, unknown>> = [
        { deltaToken: DELTA_TOK, ...WIN },
        { skipToken: SKIP, startDateTime: START },
        { skipToken: SKIP, endDateTime: END },
        { deltaToken: DELTA_TOK, skipToken: SKIP },
        { deltaToken: P },
        { deltaToken: `${P}?$deltatoken=${DELTA_TOK}` },
        { skipToken: "abc" },
        { skipToken: "a".repeat(4097) },
        { skipToken: `${SKIP}&startDateTime=2000-01-01T00:00:00Z` },
      ];
      for (const args of bad) {
        const { env, text } = await call(VIEW_DELTA, args);
        expect(env.ok, JSON.stringify(args).slice(0, 80)).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
        expect(text).not.toContain(SKIP);
      }
      expect(calls).toHaveLength(0);
    });
  });

  describe("forged or malicious Graph links", () => {
    const FORGED_NEXT = [
      `https://evil.example/v1.0/me/calendarView/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com.evil.example/v1.0/me/calendarView/delta?$skiptoken=${SKIP}`,
      `http://graph.microsoft.com/v1.0/me/calendarView/delta?$skiptoken=${SKIP}`,
      // other endpoint / principal / version
      `https://graph.microsoft.com/v1.0/me/events/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/v1.0/users/victim%40contoso.com/calendarView/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/beta/me/calendarView/delta?$skiptoken=${SKIP}`,
      `${P}/extra?$skiptoken=${SKIP}`,
      // extra / different query parameters, including a re-sent or widened window (exact
      // single-pair rule: human decision, matching Microsoft's documented link shape)
      `${P}?$skiptoken=${SKIP}&startDateTime=2000-01-01T00:00:00Z`,
      `${P}?startDateTime=2000-01-01T00:00:00Z&$skiptoken=${SKIP}`,
      `${P}?$skiptoken=${SKIP}&endDateTime=2099-01-01T00:00:00Z`,
      `${P}?$skiptoken=${SKIP}&$skiptoken=${SKIP}`,
      `${P}?$skiptoken=${SKIP}&toString=x`,
      `${P}?$skiptoken=${SKIP}&constructor=x`,
      `${P}?$skiptoken=${SKIP}&__proto__=x`,
      `${P}?$skiptoken=${SKIP}&hasOwnProperty=x`,
      `${P}?$skiptoken=${SKIP}&$select=body`,
      `${P}?$skiptoken=${SKIP}&$expand=attachments`,
      `${P}?$deltatoken=${DELTA_TOK}`, // wrong name for a nextLink
      P,
      `${P}?$skiptoken=`,
      `${P}?$skiptoken=${SKIP}#frag`,
      `${P}?$skiptoken=abc`,
      `${P}?$skiptoken=abcdefgh%26startDateTime%3D2000-01-01`,
    ];

    it("rejects every forged nextLink with one request made and no raw link or token returned", async () => {
      for (const bad of FORGED_NEXT) {
        calls = [];
        responses.length = 0;
        responses.push(page([{ id: "1" }], { "@odata.nextLink": bad }));
        const { env, text } = await call(VIEW_DELTA, { ...WIN, top: 1000 });
        expect(env.ok, bad.slice(0, 120)).toBe(false);
        expect(env.error?.code, bad.slice(0, 120)).toBe("invalid_input");
        expect(calls, bad.slice(0, 120)).toHaveLength(1);
        expect(text).not.toContain("evil.example");
        expect(text).not.toContain(TOKEN);
      }
    });

    it("rejects forged deltaLinks (origin, user, extra query, wrong token name)", async () => {
      const bad = [
        `https://evil.example/v1.0/me/calendarView/delta?$deltatoken=${DELTA_TOK}`,
        `https://graph.microsoft.com/v1.0/users/victim%40contoso.com/calendarView/delta?$deltatoken=${DELTA_TOK}`,
        `${P}?$deltatoken=${DELTA_TOK}&startDateTime=2000-01-01T00:00:00Z`,
        `${P}?$deltatoken=${DELTA_TOK}&toString=x`,
        `${P}?$deltatoken=${DELTA_TOK}&$select=body`,
        `${P}?$skiptoken=${DELTA_TOK}`,
        `${P}?$deltatoken=latest`,
      ];
      for (const link of bad) {
        calls = [];
        responses.length = 0;
        responses.push(page([{ id: "1" }], { "@odata.deltaLink": link }));
        const { env, text } = await call(VIEW_DELTA, WIN);
        expect(env.ok, link.slice(0, 120)).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
        expect(text).not.toContain("evil.example");
        expect(text).not.toContain("deltatoken");
      }
    });

    it("graph_error for both links, neither link, or a repeating skip token", async () => {
      responses.push(page([], { "@odata.nextLink": nextLink(), "@odata.deltaLink": deltaLink() }));
      expect((await call(VIEW_DELTA, WIN)).env.error?.code).toBe("graph_error");
      responses.push(page([]));
      expect((await call(VIEW_DELTA, WIN)).env.error?.code).toBe("graph_error");
      responses.push(page([{ id: "1" }], { "@odata.nextLink": nextLink() }));
      responses.push(page([{ id: "2" }], { "@odata.nextLink": nextLink() }));
      expect((await call(VIEW_DELTA, { ...WIN, top: 1000 })).env.error?.code).toBe("graph_error");
      responses.push(page([{ id: "1" }], { "@odata.nextLink": nextLink() }));
      expect((await call(VIEW_DELTA, { skipToken: SKIP, top: 1 })).env.error?.code).toBe(
        "graph_error",
      );
    });
  });

  describe("results and errors", () => {
    it("keeps @removed markers, never returns body/attachments/extensions or links", async () => {
      responses.push(
        page(
          [
            {
              "@odata.type": "#microsoft.graph.event",
              "@odata.etag": 'W/"x"',
              id: "e1",
              subject: "Standup",
              start: { dateTime: "2024-03-02T10:00:00.0000000", timeZone: "UTC" },
              body: { contentType: "html", content: "<p>TOP SECRET BODY</p>" },
              bodyPreview: "preview",
              attendees: [{ emailAddress: { address: "a@b.com" } }],
              attachments: [{ id: "att" }],
              extensions: [{ id: "ext" }],
              webLink: "https://outlook.office365.com/x",
              "@odata.id": "https://graph.microsoft.com/v1.0/x",
            },
            { id: "e2", "@removed": { reason: "deleted", other: "x" } },
          ],
          { "@odata.deltaLink": deltaLink() },
        ),
      );
      const { env, text } = await call(VIEW_DELTA, WIN);
      expect(env.data).toEqual([
        {
          "@odata.type": "#microsoft.graph.event",
          "@odata.etag": 'W/"x"',
          id: "e1",
          subject: "Standup",
          start: { dateTime: "2024-03-02T10:00:00.0000000", timeZone: "UTC" },
        },
        { id: "e2", "@removed": { reason: "deleted" } },
      ]);
      expect(text).not.toContain("TOP SECRET");
      expect(text).not.toContain("graph.microsoft.com");
      expect(text).not.toContain("$skiptoken");
      expect(text).not.toContain(TOKEN);
    });

    it("fails the whole page closed on any malformed change, never returning its token", async () => {
      const fresh = "ZnJlc2gtZGVsdGEtdG9rZW4tbmV2ZXItcmV0dXJuZWQ_0123456789";
      for (const bad of [
        "junk",
        null,
        { subject: "no id" },
        { id: "" },
        { id: "e3", "@removed": { reason: "<script>" } },
        { id: "e4", "@removed": "oops" },
      ]) {
        for (const links of [
          { "@odata.deltaLink": deltaLink(fresh) },
          { "@odata.nextLink": nextLink(fresh) },
        ]) {
          calls = [];
          responses.length = 0;
          responses.push(page([{ id: "ok" }, bad], links));
          const { env, text } = await call(VIEW_DELTA, { ...WIN, top: 1000 });
          expect(env.error?.code, JSON.stringify(bad)).toBe("graph_error");
          expect(text).not.toContain(fresh);
          expect(calls).toHaveLength(1);
        }
      }
    });

    it("fails closed on a page without a value array instead of returning its token", async () => {
      const fresh = "ZnJlc2gtZGVsdGEtdG9rZW4tbmV2ZXItcmV0dXJuZWQ_0123456789";
      for (const body of [
        { "@odata.deltaLink": deltaLink(fresh) },
        { value: "not-an-array", "@odata.deltaLink": deltaLink(fresh) },
        { value: null, "@odata.nextLink": nextLink(fresh) },
      ]) {
        calls = [];
        responses.push(() => json(body));
        const { env, text } = await call(VIEW_DELTA, WIN);
        expect(env.error?.code, JSON.stringify(body)).toBe("graph_error");
        expect(text).not.toContain(fresh);
        expect(calls).toHaveLength(1);
      }
    });

    it("maps 410 Gone (also mid-round) and 4xx syncState codes to restart_required", async () => {
      responses.push(() => json({ error: { code: "syncStateNotFound", message: "gone" } }, 410));
      const a = await call(VIEW_DELTA, { deltaToken: DELTA_TOK });
      expect(a.env.error?.code).toBe("restart_required");
      expect(calls).toHaveLength(1);
      responses.push(page([{ id: "1" }], { "@odata.nextLink": nextLink() }));
      responses.push(() => json({ error: { code: "syncStateNotFound" } }, 410));
      expect((await call(VIEW_DELTA, { ...WIN, top: 1000 })).env.error?.code).toBe(
        "restart_required",
      );
      responses.push(() => json({ error: { code: "resyncRequired" } }, 400));
      expect((await call(VIEW_DELTA, { skipToken: SKIP })).env.error?.code).toBe(
        "restart_required",
      );
      responses.push(() => json({ error: { code: "BadRequest" } }, 400));
      expect((await call(VIEW_DELTA, WIN)).env.error?.code).toBe("graph_error");
    });

    it("maps 403 to capability_gap and 401 to reauth_required (no token echoed)", async () => {
      responses.push(() => json({ error: { code: "ErrorAccessDenied", message: TOKEN } }, 403));
      const gap = await call(VIEW_DELTA, WIN);
      expect(gap.env.error?.code).toBe("capability_gap");
      expect(gap.text).not.toContain(TOKEN);
      responses.push(() => json({ error: { code: "x" } }, 401));
      expect((await call(VIEW_DELTA, WIN)).env.error?.code).toBe("reauth_required");
    });
  });

  describe("surface", () => {
    it("is not batch-reachable, and /me/events/delta is not reachable either", async () => {
      expect(handlers.get(VIEW_DELTA)?.batchPolicy).toBeUndefined();
      for (const url of [
        `/me/calendarView/delta()?startDateTime=${START}&endDateTime=${END}`,
        `/me/calendarView/delta?startDateTime=${START}&endDateTime=${END}`,
        `/me/calendarView/delta?$deltatoken=${DELTA_TOK}`,
        "/me/events/delta",
        "/me/events/delta()",
        "/me/events/delta?$select=id",
        "/me/calendar/events/delta",
        "/users/a@b.com/events/delta",
      ]) {
        const { env } = await call("graph-batch", { requests: [{ id: "1", method: "GET", url }] });
        expect(env.ok, url).toBe(false);
        expect(env.error?.code, url).toBe("invalid_input");
      }
      expect(calls).toHaveLength(0);
    });

    it("registers before graph-batch and exposes only the intended inputs", () => {
      const names = [...handlers.keys()];
      for (const n of [SHARED_EVENTS, SHARED_VIEW]) {
        expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
        expect(names.indexOf(n)).toBeLessThan(names.indexOf("graph-batch"));
      }
      expect(names.indexOf(VIEW_DELTA)).toBeLessThan(names.indexOf("graph-batch"));
      expect(Object.keys(handlers.get(VIEW_DELTA)?.inputSchema ?? {}).sort()).toEqual(
        ["deltaToken", "endDateTime", "skipToken", "startDateTime", "timeZone", "top"].sort(),
      );
    });
  });
});
