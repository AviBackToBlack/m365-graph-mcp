import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { MAX_WINDOW_DAYS } from "./calendar.js";
import { createHandlers } from "./index.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const VIEW = "get-group-calendar-view";
const EVENTS = "list-group-events";
const GID = "123e4567-e89b-12d3-a456-426614174000";
const OTHER_GID = "223e4567-e89b-12d3-a456-426614174000";
const START = "2024-01-01T00:00:00Z";
const END = "2024-01-08T00:00:00Z";
const WIN = { startDateTime: START, endDateTime: END };

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

const json = (body: unknown, status = 200): Response =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
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
const prefer = (i = 0): string | undefined => (calls[i] as Call).headers["Prefer"];

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

const TOOLS = [
  { name: VIEW, nav: "calendarView", params: { "group-id": GID, ...WIN } },
  { name: EVENTS, nav: "events", params: { "group-id": GID } },
] as const;

describe("catalog and schema", () => {
  it("pins the two v1.0 GET entries and the Group.Read.All scope", () => {
    for (const t of TOOLS) {
      const e = catalog.find((x) => x.toolName === t.name);
      expect(e?.pathPattern).toBe(`/groups/{group-id}/${t.nav}`);
      expect(e?.method).toBe("get");
      expect(e?.["apiVersion"]).toBeUndefined();
      expect(e?.["workScopes"]).toEqual(["Group.Read.All"]);
      expect(e?.["supportsTimezone"]).toBe(true);
    }
  });

  it("describes the typed contracts, not raw Graph syntax (D15)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    const tip = (n: string) => String(catalog.find((x) => x.toolName === n)?.["llmTip"]);
    for (const n of [VIEW, EVENTS]) {
      expect(desc(n)).toMatch(/group-id is the group's GUID/);
      expect(desc(n)).toMatch(/Typed parameters only/);
      expect(desc(n)).toMatch(/Prefer: outlook\.timezone/);
      expect(desc(n)).toMatch(/body and attachment objects are never returned/);
      expect(desc(n)).toMatch(/never as an empty successful list/);
      expect(tip(n)).toMatch(/not raw OData/);
    }
    expect(desc(VIEW)).toMatch(/startDateTime and endDateTime are required/);
    expect(desc(VIEW)).toMatch(
      /at most 366 days \(a PROJECT SAFETY LIMIT, not a Microsoft Graph limit\)/,
    );
    expect(desc(VIEW)).toMatch(/skip is not offered/);
    expect(desc(VIEW)).toMatch(/createdDateTime and lastModifiedDateTime are not selectable/);
    expect(desc(EVENTS)).toMatch(/without expanding recurring occurrences/);
    expect(desc(EVENTS)).toMatch(/use get-group-calendar-view/);
    expect(desc(EVENTS)).toMatch(/No calendar-view window parameters/);
  });

  it("exposes only the typed options (no extendedPropertyId, no skip on the view)", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = (n: string) =>
      Object.keys(
        (tools.find((t) => t.name === n)?.inputSchema.properties ?? {}) as Record<string, unknown>,
      ).sort();
    expect(props(VIEW)).toEqual(
      [
        "endDateTime",
        "filter",
        "group-id",
        "orderby",
        "select",
        "startDateTime",
        "timeZone",
        "top",
      ].sort(),
    );
    expect(props(EVENTS)).toEqual(
      ["filter", "group-id", "orderby", "select", "skip", "timeZone", "top"].sort(),
    );
  });
});

describe.each(TOOLS)("$name", (t) => {
  it("sends a GET to the group collection with the default compact select and no body", async () => {
    responses.push(() => json({ value: [{ id: "e1", subject: "s" }] }));
    const { env } = await call(t.name, { ...t.params });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(`/v1.0/groups/${GID}/${t.nav}`);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().searchParams.get("$top")).toBe("25");
    const select = parsed().searchParams.get("$select") ?? "";
    expect(select.split(",")).toContain("subject");
    expect(select.split(",")).not.toContain("body");
    expect(calls[0]?.body).toBeUndefined();
    expect(prefer()).toBeUndefined();
    expect(env.data).toEqual([{ id: "e1", subject: "s" }]);
  });

  it("projects to the effective select: body, attachments and annotations are dropped", async () => {
    responses.push(() =>
      json({
        "@odata.context": "ctx",
        value: [
          {
            "@odata.etag": "W/x",
            id: "e1",
            subject: "s",
            body: { content: "SECRET-BODY" },
            attachments: [1],
          },
        ],
      }),
    );
    const { env, text } = await call(t.name, { ...t.params, select: ["id", "subject"] });
    expect(parsed().searchParams.get("$select")).toBe("id,subject");
    // The shared projector keeps only @odata.type/@odata.etag annotations (as for user events).
    expect(env.data).toEqual([{ "@odata.etag": "W/x", id: "e1", subject: "s" }]);
    expect(text).not.toContain("SECRET-BODY");
    expect(text).not.toContain("attachments");
    expect(text).not.toContain("ctx");
  });

  it("sends the canonical Prefer outlook.timezone and keeps it on followed pages", async () => {
    const link = `https://graph.microsoft.com/v1.0/groups/${GID}/${t.nav}?${
      t.nav === "calendarView" ? `startDateTime=${START}&endDateTime=${END}&` : ""
    }%24skiptoken=abc`;
    responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": link }));
    responses.push(() => json({ value: [{ id: "2" }] }));
    const { env } = await call(t.name, { ...t.params, top: 150, timeZone: "europe/berlin" });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(2);
    expect(prefer(0)).toBe('outlook.timezone="Europe/Berlin"');
    expect(prefer(1)).toBe('outlook.timezone="Europe/Berlin"');
    expect(parsed(1).pathname).toBe(`/v1.0/groups/${GID}/${t.nav}`);
  });

  it("rejects bad timeZone values with zero network calls", async () => {
    for (const timeZone of [
      "",
      "Mars/Olympus_Mons",
      'UTC", outlook.allow-unsafe-html',
      "UTC\r\nX-Evil: 1",
      1,
    ]) {
      const { env, text } = await call(t.name, { ...t.params, timeZone });
      expect(env.error?.code).toBe("invalid_input");
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });

  it("rejects select outside the allowlist, raw OData and unknown options with zero network calls", async () => {
    const bad: Array<Record<string, unknown>> = [
      { select: ["body"] },
      { select: ["attachments"] },
      { select: ["calendar"] },
      { select: [] },
      { expand: ["attachments"] },
      { extendedPropertyId: "String {00020329-0000-0000-C000-000000000046} Name x" },
      { search: "x" },
      { count: true },
      { headers: { Prefer: "outlook.allow-unsafe-html" } },
      { Prefer: "outlook.allow-unsafe-html" },
      { top: 0 },
      { top: 501 },
      { orderby: [{ property: "body" }] },
      { unknownKey: 1 },
    ];
    for (const extra of bad) {
      calls = [];
      const { env, text } = await call(t.name, { ...t.params, ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(env.error?.code, JSON.stringify(extra)).toBe("invalid_input");
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("rejects non-GUID group ids and path injection with zero network calls", async () => {
    for (const id of [
      "delta",
      "$ref",
      "$count",
      "..",
      "alice@contoso.com", // a UPN is a user-id shape, not a group id
      `${GID}/members`,
      `${GID}/events`,
      `${GID}%2Fmembers`,
      `${GID}%2F..%2Fusers`,
      `${GID}?$top=1`,
      `${GID}#x`,
      "microsoft.graph.getByIds()",
      `${GID}/${t.nav}/delta`,
      "",
    ]) {
      calls = [];
      const { env } = await call(t.name, { ...t.params, "group-id": id });
      expect(env.ok, id).toBe(false);
      expect(env.error?.code, id).toBe("invalid_input");
      expect(calls, id).toHaveLength(0);
    }
    const missing = await call(t.name, { ...t.params, "group-id": undefined });
    expect(missing.env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("refuses a nextLink for another group, collection, host or API version (fail closed)", async () => {
    const other = t.nav === "events" ? "calendarView" : "events";
    for (const bad of [
      `https://evil.example/v1.0/groups/${GID}/${t.nav}?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${OTHER_GID}/${t.nav}?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/${other}?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/calendar/${t.nav}?$skiptoken=1`,
      `https://graph.microsoft.com/v1.0/groups/${GID}/${t.nav}/delta`,
      `https://graph.microsoft.com/v1.0/groups/${GID}`,
      `https://graph.microsoft.com/v1.0/me/${t.nav}?$skiptoken=1`,
      `https://graph.microsoft.com/beta/groups/${GID}/${t.nav}?$skiptoken=1`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": bad }));
      const { env } = await call(t.name, { ...t.params, top: 150 });
      expect(env.error?.code, bad).toBe("invalid_input");
      expect(calls, bad).toHaveLength(1);
    }
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "x" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(t.name, { ...t.params })).env.error?.code).toBe("graph_error");
    }
  });

  it("maps 401, 403 and 404/400 (e.g. a group without a calendar) to the shared conventions", async () => {
    for (const [status, code] of [
      [401, "reauth_required"],
      [403, "capability_gap"],
      [404, "graph_error"],
      [400, "graph_error"],
    ] as const) {
      responses.push(() => json({ error: { code: "ErrorItemNotFound", message: "m" } }, status));
      const { env, text } = await call(t.name, { ...t.params });
      expect(env.ok).toBe(false);
      expect(env.error?.code, String(status)).toBe(code);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe(VIEW, () => {
  it("sends the UTC window and the optional filter/orderby", async () => {
    await call(VIEW, {
      "group-id": GID,
      startDateTime: "2024-01-01T09:00:00+01:00",
      endDateTime: END,
      filter: "isAllDay eq false",
      orderby: [{ property: "start/dateTime", direction: "desc" }],
    });
    const u = parsed();
    expect(u.searchParams.get("startDateTime")).toBe("2024-01-01T08:00:00Z");
    expect(u.searchParams.get("endDateTime")).toBe(END);
    expect(u.searchParams.get("$filter")).toBe("isAllDay eq false");
    expect(u.searchParams.get("$orderby")).toBe("start/dateTime desc");
  });

  it("does not let timeZone change the window", async () => {
    await call(VIEW, { "group-id": GID, ...WIN, timeZone: "Pacific Standard Time" });
    expect(parsed().searchParams.get("startDateTime")).toBe(START);
    expect(prefer()).toBe('outlook.timezone="Pacific Standard Time"');
  });

  it("does not offer createdDateTime/lastModifiedDateTime in select (Graph rejects them on calendarView)", async () => {
    for (const p of ["createdDateTime", "lastModifiedDateTime"]) {
      const { env } = await call(VIEW, { "group-id": GID, ...WIN, select: [p] });
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  const MS = 86_400_000;
  const plus = (days: number, s = 0): string =>
    new Date(Date.parse(START) + days * MS + s * 1000).toISOString();
  it.each([
    ["missing both", {}],
    ["missing start", { endDateTime: END }],
    ["missing end", { startDateTime: START }],
    ["no zone (local)", { startDateTime: "2024-01-01T00:00:00", endDateTime: END }],
    ["date only", { startDateTime: "2024-01-01", endDateTime: END }],
    ["reversed", { startDateTime: END, endDateTime: START }],
    ["zero length", { startDateTime: START, endDateTime: START }],
    ["too long by a second", { startDateTime: START, endDateTime: plus(MAX_WINDOW_DAYS, 1) }],
    ["a decade", { startDateTime: START, endDateTime: "2034-01-01T00:00:00Z" }],
    ["injected query", { startDateTime: `${START}&$expand=attachments`, endDateTime: END }],
  ])("rejects an invalid window (%s) with zero network calls", async (_l, args) => {
    const { env } = await call(VIEW, { "group-id": GID, ...args });
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("accepts the maximum window exactly", async () => {
    const { env } = await call(VIEW, {
      "group-id": GID,
      startDateTime: START,
      endDateTime: plus(MAX_WINDOW_DAYS),
    });
    expect(env.ok).toBe(true);
  });

  it("does not offer skip", async () => {
    const { env } = await call(VIEW, { "group-id": GID, ...WIN, skip: 1 });
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });
});

describe(EVENTS, () => {
  it("offers createdDateTime/lastModifiedDateTime, skip and filter/orderby", async () => {
    await call(EVENTS, {
      "group-id": GID,
      select: ["id", "createdDateTime", "lastModifiedDateTime"],
      skip: 5,
      filter: "isAllDay eq false",
      orderby: [{ property: "start/dateTime" }],
    });
    const u = parsed();
    expect(u.searchParams.get("$select")).toBe("id,createdDateTime,lastModifiedDateTime");
    expect(u.searchParams.get("$skip")).toBe("5");
    expect(u.searchParams.get("$orderby")).toBe("start/dateTime");
  });

  it("rejects window parameters (not part of the events list)", async () => {
    const { env } = await call(EVENTS, { "group-id": GID, ...WIN });
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("keeps recurrence in the default select so series masters show their pattern", async () => {
    await call(EVENTS, { "group-id": GID });
    expect((parsed().searchParams.get("$select") ?? "").split(",")).toContain("recurrence");
  });
});

describe("graph-batch reachability", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const win = `startDateTime=${START}&endDateTime=${END}`;
  const sel = "$select=id,subject,start,end";
  const ok = (body: unknown) => () => json({ responses: [{ id: "1", status: 200, body }] });

  it("registers both tools with a batch policy before graph-batch", () => {
    const handlers = createHandlers(
      new GraphClient({ tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) }, catalog }),
      catalog,
    );
    const names = [...handlers.keys()];
    for (const n of [VIEW, EVENTS]) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
      expect(names.indexOf(n)).toBeLessThan(names.indexOf("graph-batch"));
    }
  });

  it("accepts in-policy subrequests and projects like the single path", async () => {
    const VALID: Array<[string, Record<string, string>?]> = [
      [`/groups/${GID}/calendarView?${win}&${sel}`],
      [
        `/groups/${GID}/calendarView?${win}&${sel}&$top=100&$filter=isAllDay%20eq%20false`,
        { Prefer: 'outlook.timezone="Europe/Berlin"' },
      ],
      [`/groups/${GID}/events?${sel}`],
      [`/groups/${GID}/events?${sel}&$top=100&$skip=3&$orderby=start/dateTime%20desc`],
      [`/groups/${GID}/events?${sel}`, { Prefer: 'outlook.timezone="UTC"' }],
    ];
    for (const [url, headers] of VALID) {
      calls = [];
      responses.push(ok({ value: [{ id: "a", subject: "s", body: { content: "SECRET-BODY" } }] }));
      const { env, text } = await batch(url, headers);
      expect(env.ok, url).toBe(true);
      expect(calls, url).toHaveLength(1);
      expect(calls[0]?.url).toContain("/$batch");
      expect(text).not.toContain("SECRET-BODY");
    }
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const view = `/groups/${GID}/calendarView`;
    const ev = `/groups/${GID}/events`;
    const bad: Array<[string, Record<string, string>?]> = [
      // $select is required in batch
      [`${view}?${win}`],
      [`${ev}`],
      [`${ev}?$top=10`],
      // missing/partial/invalid/oversized window
      [`${view}?${sel}`],
      [`${view}?startDateTime=${START}&${sel}`],
      [`${view}?startDateTime=2024-01-01T00:00:00&endDateTime=${END}&${sel}`],
      [`${view}?startDateTime=${END}&endDateTime=${START}&${sel}`],
      [`${view}?startDateTime=${START}&endDateTime=2034-01-01T00:00:00Z&${sel}`],
      [`${view}?${win}&${win}&${sel}`],
      // window params are not part of events; skip is not part of the view
      [`${ev}?${sel}&${win}`],
      [`${view}?${win}&${sel}&$skip=1`],
      // other options and limits
      [`${view}?${win}&${sel}&$top=101`],
      [`${view}?${win}&${sel}&$expand=attachments`],
      [`${view}?${win}&${sel}&$search=%22x%22`],
      [`${view}?${win}&$select=id,body`],
      [`${view}?${win}&$select=id,createdDateTime`],
      [`${ev}?$select=id,body`],
      [`${ev}?$select=id,attachments`],
      [`${ev}?${sel}&$orderby=body`],
      [`${ev}?${sel}&$skip=10001`],
      // headers
      [`${view}?${win}&${sel}`, { Prefer: "outlook.allow-unsafe-html" }],
      [`${view}?${win}&${sel}`, { Prefer: 'outlook.timezone="Mars/Base"' }],
      [`${view}?${win}&${sel}`, { Prefer: 'outlook.body-content-type="text"' }],
      [`${ev}?${sel}`, { ConsistencyLevel: "eventual" }],
      // group-id injection and path variants
      [`/groups/delta/calendarView?${win}&${sel}`],
      [`/groups/delta/events?${sel}`],
      [`/groups/$ref/events?${sel}`],
      [`/groups/../events?${sel}`],
      [`/groups/${GID}%2Fmembers/events?${sel}`],
      [`/groups/alice@contoso.com/events?${sel}`],
      [`/groups/${GID}/events/delta?${sel}`],
      [`/groups/${GID}/calendarView/delta?${win}&${sel}`],
      [`/groups/${GID}/calendar/events?${sel}`],
      [`/groups/${GID}/calendar/calendarView?${win}&${sel}`],
      [`/groups/${GID}/events/someEventId?${sel}`],
      [`/groups/${GID}/events/someEventId/instances?${win}&${sel}`],
      [`/groups/${GID}/threads/someThreadId/posts?${sel}`],
      [`/groups/${GID}/memberOf?${sel}`],
      [`/groups/${GID}/events?${sel}#frag`],
    ];
    for (const [url, headers] of bad) {
      const { env, text } = await batch(url, headers);
      expect(env.ok, `${url} ${JSON.stringify(headers)}`).toBe(false);
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });
});
