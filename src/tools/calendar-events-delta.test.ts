import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { registerCalendarEventsDeltaTool } from "./calendar-events-delta.js";
import { createHandlers } from "./index.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const SKIP = "R0usmcCM996atia_sQz-4aBc9DeFgHiJkLmNoPqRsTuVwXyZ0123456789";
const DELTA_TOK = "ZmFrZS1kZWx0YS10b2tlbi1mb3ItdGVzdHM_0123456789abcdefABCDEF";
const START = "2024-03-01T00:00:00Z";

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

const EVENTS_DELTA = "list-calendar-events-delta";
const VIEW_DELTA = "list-calendar-view-delta";
const P = "https://graph.microsoft.com/beta/me/events/delta";
const nextLink = (tok = SKIP) => `${P}?$skiptoken=${tok}`;
const deltaLink = (tok = DELTA_TOK) => `${P}?$deltatoken=${tok}`;
const FRESH = "ZnJlc2gtZGVsdGEtdG9rZW4tbmV2ZXItcmV0dXJuZWQ_0123456789";
const page =
  (value: unknown[], links: Record<string, string> = {}) =>
  () =>
    json({ "@odata.context": "https://graph.microsoft.com/beta/$metadata#x", value, ...links });
const ev = (id: string) => ({
  id,
  type: "singleInstance",
  start: { dateTime: "2024-03-02T10:00:00.0000000", timeZone: "UTC" },
  end: { dateTime: "2024-03-02T11:00:00.0000000", timeZone: "UTC" },
});

describe("catalog", () => {
  it("is the beta-only /me/events/delta() entry (D2) and the calendarView delta stays v1.0", () => {
    const entry = catalog.find((e) => e.toolName === EVENTS_DELTA);
    expect(entry?.pathPattern).toBe("/me/events/delta()");
    expect(entry?.method).toBe("get");
    expect(entry?.["apiVersion"]).toBe("beta");
    // D6: delegated least privilege per the beta event: delta page.
    expect(entry?.["scopes"]).toEqual(["Calendars.ReadBasic"]);
    // D5: catalog metadata matches the beta contract.
    expect(entry?.["supportsTimezone"]).toBe(true);
    const tip = String(entry?.["llmTip"]);
    expect(tip).toContain("across all calendars in the signed-in user's mailbox");
    expect(tip).toContain("$select, $expand, $filter, $orderby and $search are not supported");
    expect(tip).toContain("meta.skipToken");
    expect(tip).toContain("meta.deltaToken");
    expect(tip.toLowerCase()).not.toContain("use $select");
    expect(handlers.has(EVENTS_DELTA)).toBe(true);
  });

  it("advertises the beta-only, limited-projection description (D5)", async () => {
    const tool = (await mcp.listTools()).tools.find((t) => t.name === EVENTS_DELTA);
    expect(tool?.description).toMatch(/BETA-ONLY/);
    expect(tool?.description).toMatch(/all calendars in the signed-in user's mailbox/);
    expect(tool?.description).toMatch(/only id, type, start and end/);
    expect(tool?.description).toMatch(/get-calendar-event/);
    expect(catalog.find((e) => e.toolName === VIEW_DELTA)?.["apiVersion"]).toBeUndefined();
  });
});

describe(EVENTS_DELTA, () => {
  describe("initial call", () => {
    it("without startDateTime: beta path, no query at all, page size via Prefer", async () => {
      responses.push(page([ev("1")], { "@odata.deltaLink": deltaLink() }));
      const { env } = await call(EVENTS_DELTA, {});
      expect(env.ok).toBe(true);
      expect(calls).toHaveLength(1);
      expect(parsed().origin).toBe("https://graph.microsoft.com");
      expect(parsed().pathname).toBe("/beta/me/events/delta()");
      expect([...parsed().searchParams.keys()]).toEqual([]);
      expect(calls[0]?.headers["Prefer"]).toBe("odata.maxpagesize=100");
      expect(calls[0]?.method).toBe("GET");
      expect(env.meta).toMatchObject({ roundComplete: true, deltaToken: DELTA_TOK });
    });

    it("with startDateTime: canonical UTC (offset converted), only that query parameter", async () => {
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      await call(EVENTS_DELTA, { startDateTime: "2024-03-01T09:00:00.1230000+01:00" });
      expect(parsed().pathname).toBe("/beta/me/events/delta()");
      expect([...parsed().searchParams.keys()]).toEqual(["startDateTime"]);
      expect(parsed().searchParams.get("startDateTime")).toBe("2024-03-01T08:00:00.123Z");
    });

    it("sends timeZone as outlook.timezone after the page size; caps page size at 100", async () => {
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      await call(EVENTS_DELTA, { timeZone: "Pacific Standard Time", top: 1000 });
      expect(calls[0]?.headers["Prefer"]).toBe(
        'odata.maxpagesize=100, outlook.timezone="Pacific Standard Time"',
      );
    });

    it("schema negatives: endDateTime, OData options, unknown keys, bad values; zero calls", async () => {
      const bad: Array<Record<string, unknown>> = [
        { endDateTime: "2024-03-08T00:00:00Z" },
        { startDateTime: START, endDateTime: "2024-03-08T00:00:00Z" },
        { $select: "body" },
        { $expand: "attachments" },
        { $filter: "x" },
        { $orderby: "start/dateTime" },
        { $search: "x" },
        { select: ["id"] },
        { changeType: "created" },
        { url: P },
        { "user-id": "a@b.com" },
        { startDateTime: "2024-03-01T00:00:00" },
        { startDateTime: "tomorrow" },
        { startDateTime: "2024-03-01T00:00:00+15:00" },
        { startDateTime: 5 },
        { timeZone: "Not/AZone" },
        { timeZone: `UTC"; x=1` },
        { top: 0 },
        { top: 1001 },
      ];
      for (const args of bad) {
        const { env } = await call(EVENTS_DELTA, args);
        expect(env.ok, JSON.stringify(args)).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
      }
      expect(calls).toHaveLength(0);
    });
  });

  describe("rounds and continuation", () => {
    it("follows nextLink in one call via skiptoken-only beta requests, then returns the deltaToken", async () => {
      responses.push(page([ev("1")], { "@odata.nextLink": nextLink() }));
      responses.push(page([ev("2")], { "@odata.deltaLink": deltaLink() }));
      const { env } = await call(EVENTS_DELTA, { startDateTime: START, top: 1000 });
      expect(env.ok).toBe(true);
      expect((env.data as unknown[]).length).toBe(2);
      expect(env.meta).toMatchObject({ roundComplete: true, deltaToken: DELTA_TOK, pages: 2 });
      expect(parsed(1).pathname).toBe("/beta/me/events/delta()");
      expect([...parsed(1).searchParams.keys()]).toEqual(["$skiptoken"]);
      expect(parsed(1).searchParams.get("$skiptoken")).toBe(SKIP);
      expect(calls[1]?.headers["Prefer"]).toBe(calls[0]?.headers["Prefer"]);
    });

    it("stops at top on a page boundary with a skipToken, which resumes on the beta path", async () => {
      responses.push(page([ev("1"), ev("2")], { "@odata.nextLink": nextLink() }));
      const first = await call(EVENTS_DELTA, { top: 1 });
      expect(first.env.meta).toMatchObject({
        hasMore: true,
        roundComplete: false,
        skipToken: SKIP,
      });
      expect(calls).toHaveLength(1);
      calls = [];
      responses.push(page([], { "@odata.deltaLink": deltaLink() }));
      await call(EVENTS_DELTA, { skipToken: SKIP, timeZone: "UTC" });
      expect(parsed().pathname).toBe("/beta/me/events/delta()");
      expect([...parsed().searchParams.keys()]).toEqual(["$skiptoken"]);
      expect(calls[0]?.headers["Prefer"]).toBe('odata.maxpagesize=100, outlook.timezone="UTC"');
    });

    it("deltaToken resume sends only $deltatoken on the beta path", async () => {
      responses.push(
        page([], { "@odata.deltaLink": deltaLink("bmV4dC1yb3VuZC10b2tlbi0xMjM0NTY") }),
      );
      const { env } = await call(EVENTS_DELTA, { deltaToken: DELTA_TOK });
      expect(parsed().pathname).toBe("/beta/me/events/delta()");
      expect([...parsed().searchParams.keys()]).toEqual(["$deltatoken"]);
      expect(parsed().searchParams.get("$deltatoken")).toBe(DELTA_TOK);
      expect(env.meta?.["deltaToken"]).toBe("bmV4dC1yb3VuZC10b2tlbi0xMjM0NTY");
    });

    it("rejects startDateTime with a token, both tokens, and URL/bad tokens; zero calls", async () => {
      const bad: Array<Record<string, unknown>> = [
        { deltaToken: DELTA_TOK, startDateTime: START },
        { skipToken: SKIP, startDateTime: START },
        { deltaToken: DELTA_TOK, skipToken: SKIP },
        { deltaToken: P },
        { deltaToken: `${P}?$deltatoken=${DELTA_TOK}` },
        { skipToken: "abc" },
        { skipToken: `${SKIP}&startDateTime=2000-01-01T00:00:00Z` },
      ];
      for (const args of bad) {
        const { env, text } = await call(EVENTS_DELTA, args);
        expect(env.ok, JSON.stringify(args).slice(0, 80)).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
        expect(text).not.toContain(SKIP);
      }
      expect(calls).toHaveLength(0);
    });
  });

  describe("forged or version-confused Graph links", () => {
    const FORGED_NEXT = [
      // version confusion: a v1.0 link for the beta endpoint, and the neighbouring endpoints
      `https://graph.microsoft.com/v1.0/me/events/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/v1.0/me/events/delta()?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/beta/me/calendarView/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/v1.0/me/calendarView/delta?$skiptoken=${SKIP}`,
      // other beta paths / principals
      `https://graph.microsoft.com/beta/me/calendar/events/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/beta/me/calendars/x/events/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/beta/users/victim%40contoso.com/events/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/beta/me/events?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com/beta/me/messages/delta?$skiptoken=${SKIP}`,
      `${P}/extra?$skiptoken=${SKIP}`,
      // origin
      `https://evil.example/beta/me/events/delta?$skiptoken=${SKIP}`,
      `https://graph.microsoft.com.evil.example/beta/me/events/delta?$skiptoken=${SKIP}`,
      `http://graph.microsoft.com/beta/me/events/delta?$skiptoken=${SKIP}`,
      // exact single-pair rule (an echoed or widened startDateTime included)
      `${P}?$skiptoken=${SKIP}&startDateTime=2000-01-01T00:00:00Z`,
      `${P}?startDateTime=2000-01-01T00:00:00Z&$skiptoken=${SKIP}`,
      `${P}?$skiptoken=${SKIP}&endDateTime=2099-01-01T00:00:00Z`,
      `${P}?$skiptoken=${SKIP}&$skiptoken=${SKIP}`,
      `${P}?$skiptoken=${SKIP}&__proto__=x`,
      `${P}?$skiptoken=${SKIP}&$select=body`,
      `${P}?$deltatoken=${DELTA_TOK}`,
      P,
      `${P}?$skiptoken=`,
      `${P}?$skiptoken=${SKIP}#frag`,
      `${P}?$skiptoken=abc`,
      `${P}?$skiptoken=abcdefgh%26startDateTime%3D2000-01-01`,
    ];

    it("rejects every forged nextLink with one request made and no raw link returned", async () => {
      for (const bad of FORGED_NEXT) {
        calls = [];
        responses.length = 0;
        responses.push(page([ev("1")], { "@odata.nextLink": bad }));
        const { env, text } = await call(EVENTS_DELTA, { top: 1000 });
        expect(env.ok, bad.slice(0, 120)).toBe(false);
        expect(env.error?.code, bad.slice(0, 120)).toBe("invalid_input");
        expect(calls, bad.slice(0, 120)).toHaveLength(1);
        expect(text).not.toContain("evil.example");
        expect(text).not.toContain(TOKEN);
      }
    });

    it("rejects forged deltaLinks (version, origin, user, extra query, wrong token name)", async () => {
      const bad = [
        `https://graph.microsoft.com/v1.0/me/events/delta?$deltatoken=${DELTA_TOK}`,
        `https://graph.microsoft.com/beta/me/calendarView/delta?$deltatoken=${DELTA_TOK}`,
        `https://evil.example/beta/me/events/delta?$deltatoken=${DELTA_TOK}`,
        `https://graph.microsoft.com/beta/users/victim%40contoso.com/events/delta?$deltatoken=${DELTA_TOK}`,
        `${P}?$deltatoken=${DELTA_TOK}&startDateTime=2000-01-01T00:00:00Z`,
        `${P}?$deltatoken=${DELTA_TOK}&$select=body`,
        `${P}?$skiptoken=${DELTA_TOK}`,
        `${P}?$deltatoken=latest`,
      ];
      for (const link of bad) {
        calls = [];
        responses.length = 0;
        responses.push(page([ev("1")], { "@odata.deltaLink": link }));
        const { env, text } = await call(EVENTS_DELTA, {});
        expect(env.ok, link.slice(0, 120)).toBe(false);
        expect(env.error?.code).toBe("invalid_input");
        expect(text).not.toContain("evil.example");
        expect(text).not.toContain("deltatoken");
      }
    });

    it("graph_error for both links, neither link, or a repeating skip token", async () => {
      responses.push(page([], { "@odata.nextLink": nextLink(), "@odata.deltaLink": deltaLink() }));
      expect((await call(EVENTS_DELTA, {})).env.error?.code).toBe("graph_error");
      responses.push(page([]));
      expect((await call(EVENTS_DELTA, {})).env.error?.code).toBe("graph_error");
      responses.push(page([ev("1")], { "@odata.nextLink": nextLink() }));
      responses.push(page([ev("2")], { "@odata.nextLink": nextLink() }));
      expect((await call(EVENTS_DELTA, { top: 1000 })).env.error?.code).toBe("graph_error");
    });
  });

  describe("registration guard", () => {
    it("refuses to register unless the catalog entry is beta (D2)", () => {
      const withVersion = (apiVersion: unknown): EndpointCatalog =>
        catalog.map((e) =>
          e.toolName === EVENTS_DELTA ? { ...e, apiVersion } : e,
        ) as EndpointCatalog;
      const noClient = {} as GraphClient;
      for (const version of [undefined, "v1.0", "BETA"]) {
        expect(() =>
          registerCalendarEventsDeltaTool(new Map(), noClient, withVersion(version)),
        ).toThrowError(/^list-calendar-events-delta is a beta-only endpoint\.$/);
      }
      const ok = new Map();
      registerCalendarEventsDeltaTool(ok, noClient, withVersion("beta"));
      expect(ok.has(EVENTS_DELTA)).toBe(true);
    });
  });

  describe("results and errors", () => {
    it("projects to id/type/start/end, keeps @removed markers, drops everything else", async () => {
      responses.push(
        page(
          [
            {
              "@odata.type": "#microsoft.graph.event",
              "@odata.etag": 'W/"x"',
              ...ev("e1"),
              subject: "Standup",
              body: { contentType: "html", content: "<p>TOP SECRET BODY</p>" },
              bodyPreview: "preview",
              attendees: [{ emailAddress: { address: "a@b.com" } }],
              webLink: "https://outlook.office365.com/x",
              "@odata.id": "https://graph.microsoft.com/beta/x",
            },
            { id: "e2", "@removed": { reason: "deleted", other: "x" } },
          ],
          { "@odata.deltaLink": deltaLink() },
        ),
      );
      const { env, text } = await call(EVENTS_DELTA, {});
      expect(env.data).toEqual([ev("e1"), { id: "e2", "@removed": { reason: "deleted" } }]);
      for (const leak of [
        "TOP SECRET",
        "Standup",
        "a@b.com",
        "outlook.office365.com",
        "#microsoft.graph.event",
        "$skiptoken",
      ]) {
        expect(text).not.toContain(leak);
      }
      expect(text).not.toContain(TOKEN);
    });

    it("fails the whole page closed on any malformed change, never returning its token", async () => {
      for (const bad of [
        "junk",
        null,
        { start: {} },
        { id: "" },
        { id: "e3", "@removed": { reason: "<script>" } },
        { id: "e4", "@removed": "oops" },
      ]) {
        for (const links of [
          { "@odata.deltaLink": deltaLink(FRESH) },
          { "@odata.nextLink": nextLink(FRESH) },
        ]) {
          calls = [];
          responses.length = 0;
          responses.push(page([{ id: "ok" }, bad], links));
          const { env, text } = await call(EVENTS_DELTA, { top: 1000 });
          expect(env.error?.code, JSON.stringify(bad)).toBe("graph_error");
          expect(text).not.toContain(FRESH);
          expect(calls).toHaveLength(1);
        }
      }
    });

    it("fails closed on a page without a value array instead of returning its token", async () => {
      for (const body of [
        { "@odata.deltaLink": deltaLink(FRESH) },
        { value: "not-an-array", "@odata.deltaLink": deltaLink(FRESH) },
        { value: null, "@odata.nextLink": nextLink(FRESH) },
      ]) {
        calls = [];
        responses.push(() => json(body));
        const { env, text } = await call(EVENTS_DELTA, {});
        expect(env.error?.code, JSON.stringify(body)).toBe("graph_error");
        expect(text).not.toContain(FRESH);
      }
    });

    it("maps 410 Gone (also mid-round) and 4xx syncState codes to restart_required", async () => {
      responses.push(() => json({ error: { code: "syncStateNotFound", message: "gone" } }, 410));
      expect((await call(EVENTS_DELTA, { deltaToken: DELTA_TOK })).env.error?.code).toBe(
        "restart_required",
      );
      responses.push(page([ev("1")], { "@odata.nextLink": nextLink() }));
      responses.push(() => json({ error: { code: "syncStateNotFound" } }, 410));
      expect((await call(EVENTS_DELTA, { top: 1000 })).env.error?.code).toBe("restart_required");
      responses.push(() => json({ error: { code: "resyncRequired" } }, 400));
      expect((await call(EVENTS_DELTA, { skipToken: SKIP })).env.error?.code).toBe(
        "restart_required",
      );
      responses.push(() => json({ error: { code: "BadRequest" } }, 400));
      expect((await call(EVENTS_DELTA, {})).env.error?.code).toBe("graph_error");
    });

    it("maps 403 to capability_gap and 401 to reauth_required (no token echoed)", async () => {
      responses.push(() => json({ error: { code: "ErrorAccessDenied", message: TOKEN } }, 403));
      const gap = await call(EVENTS_DELTA, {});
      expect(gap.env.error?.code).toBe("capability_gap");
      expect(gap.text).not.toContain(TOKEN);
      responses.push(() => json({ error: { code: "x" } }, 401));
      expect((await call(EVENTS_DELTA, {})).env.error?.code).toBe("reauth_required");
    });
  });

  describe("surface", () => {
    it("is not batch-reachable under any version/path spelling", async () => {
      expect(handlers.get(EVENTS_DELTA)?.batchPolicy).toBeUndefined();
      for (const url of [
        "/me/events/delta",
        "/me/events/delta()",
        `/me/events/delta()?startDateTime=${START}`,
        `/me/events/delta?$deltatoken=${DELTA_TOK}`,
        "/beta/me/events/delta",
        "/beta/me/events/delta()",
        "https://graph.microsoft.com/beta/me/events/delta",
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
      expect(names.indexOf(EVENTS_DELTA)).toBeLessThan(names.indexOf("graph-batch"));
      expect(Object.keys(handlers.get(EVENTS_DELTA)?.inputSchema ?? {}).sort()).toEqual(
        ["deltaToken", "skipToken", "startDateTime", "timeZone", "top"].sort(),
      );
    });
  });
});
