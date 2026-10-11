import { readFileSync } from "node:fs";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { RESERVED_INPUT_NAMES } from "../registry.js";
import { createServer } from "../server.js";
import { MAX_WINDOW_DAYS } from "./calendar.js";
import { createHandlers } from "./index.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const CAL_ID =
  "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEGAAA=";
const EVT_ID =
  "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAENAAAiIsqMbYjsT5e-T7KzowPTAAAa_WKzAAA=";
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

const EVENT_LISTS = [
  { name: "list-calendar-events", path: "/v1.0/me/events", params: {} },
  {
    name: "list-specific-calendar-events",
    path: `/v1.0/me/calendars/${CAL_ID}/events`,
    params: { "calendar-id": CAL_ID },
  },
] as const;
const WINDOWED = [
  { name: "get-calendar-view", path: "/v1.0/me/calendarView", params: { ...WIN } },
  {
    name: "get-specific-calendar-view",
    path: `/v1.0/me/calendars/${CAL_ID}/calendarView`,
    params: { "calendar-id": CAL_ID, ...WIN },
  },
  {
    name: "list-calendar-event-instances",
    path: `/v1.0/me/calendars/${CAL_ID}/events/${EVT_ID}/instances`,
    params: { "calendar-id": CAL_ID, "event-id": EVT_ID, ...WIN },
  },
] as const;
const EVENT_GETS = [
  { name: "get-calendar-event", path: `/v1.0/me/events/${EVT_ID}`, params: { "event-id": EVT_ID } },
  {
    name: "get-specific-calendar-event",
    path: `/v1.0/me/calendars/${CAL_ID}/events/${EVT_ID}`,
    params: { "calendar-id": CAL_ID, "event-id": EVT_ID },
  },
] as const;
const ALL_EVENT_TOOLS = [...EVENT_LISTS, ...WINDOWED, ...EVENT_GETS];
const ALL_NAMES = [...ALL_EVENT_TOOLS.map((t) => t.name), "list-calendars"];

describe("calendar tool schemas", () => {
  it("advertises typed inputs; windows are required only where Graph requires them", async () => {
    const tools = (await mcp.listTools()).tools;
    const by = (n: string) => tools.find((t) => t.name === n);
    for (const name of ALL_NAMES) {
      const props = Object.keys(by(name)?.inputSchema.properties ?? {});
      expect(props.length, name).toBeGreaterThan(0);
      for (const p of props) {
        expect(RESERVED_INPUT_NAMES).not.toContain(p.toLowerCase());
        expect(p.startsWith("$")).toBe(false);
        expect(["expand", "search", "expandExtendedProperties"]).not.toContain(p);
      }
    }
    for (const t of WINDOWED) {
      const req = by(t.name)?.inputSchema.required ?? [];
      expect(req, t.name).toEqual(expect.arrayContaining(["startDateTime", "endDateTime"]));
    }
    for (const t of [...EVENT_LISTS, ...EVENT_GETS]) {
      expect(Object.keys(by(t.name)?.inputSchema.properties ?? {})).not.toContain("startDateTime");
    }
    expect(by("get-calendar-event")?.inputSchema.required).toEqual(["event-id"]);
    expect(by("get-specific-calendar-event")?.inputSchema.required).toEqual(
      expect.arrayContaining(["calendar-id", "event-id"]),
    );
  });

  it("offers a body only on get-*-event and timeZone only on event tools", async () => {
    const tools = (await mcp.listTools()).tools;
    const schemaOf = (n: string) => JSON.stringify(tools.find((t) => t.name === n)?.inputSchema);
    for (const t of [...EVENT_LISTS, ...WINDOWED]) {
      expect(schemaOf(t.name), t.name).not.toContain('"body"');
      expect(schemaOf(t.name), t.name).not.toContain("bodyType");
      expect(schemaOf(t.name), t.name).toContain("timeZone");
    }
    for (const t of EVENT_GETS) {
      expect(schemaOf(t.name)).toContain('"body"');
      expect(schemaOf(t.name)).toContain("bodyType");
      expect(schemaOf(t.name)).toContain("timeZone");
    }
    expect(schemaOf("list-calendars")).not.toContain("timeZone");
    expect(schemaOf("list-calendars")).not.toContain("startDateTime");
  });

  it("keeps the untrusted-content notice on every calendar tool", async () => {
    const tools = (await mcp.listTools()).tools;
    for (const name of ALL_NAMES) {
      expect(tools.find((t) => t.name === name)?.description, name).toContain("untrusted data");
    }
  });
});

describe("single-request tools", () => {
  it.each([...EVENT_LISTS, ...WINDOWED])(
    "$name: fixed path, compact default query without body, no token in result",
    async (t) => {
      responses.push(() => json({ value: [{ id: "1", subject: "hi", body: { content: "x" } }] }));
      const { r, env, text } = await call(t.name, { ...t.params });
      expect(r.isError).toBe(false);
      expect(env.ok).toBe(true);
      expect(calls).toHaveLength(1);
      expect(calls[0]?.method).toBe("GET");
      expect(parsed().origin).toBe("https://graph.microsoft.com");
      expect(parsed().pathname).toBe(t.path);
      expect(parsed().searchParams.get("$top")).toBe("25");
      const select = parsed().searchParams.get("$select") ?? "";
      expect(select.split(",")).toEqual(expect.arrayContaining(["id", "subject", "start", "end"]));
      expect(select.split(",")).not.toContain("body");
      expect(select.split(",")).not.toContain("attendees");
      expect(prefer()).toBeUndefined();
      expect(text).not.toContain(TOKEN);
      // The projector drops the body even if Graph ignored $select.
      expect(JSON.stringify(env.data)).not.toContain("content");
    },
  );

  it("only the event listings select recurrence by default (series masters)", async () => {
    await call("list-calendar-events");
    expect(parsed().searchParams.get("$select")?.split(",")).toContain("recurrence");
    calls = [];
    await call("get-calendar-view", { ...WIN });
    expect(parsed().searchParams.get("$select")?.split(",")).not.toContain("recurrence");
  });

  it("sends the window as exact UTC wire values", async () => {
    await call("get-calendar-view", { ...WIN });
    expect(parsed().searchParams.get("startDateTime")).toBe(START);
    expect(parsed().searchParams.get("endDateTime")).toBe(END);
    expect([...parsed().searchParams.keys()].sort()).toEqual(
      ["$select", "$top", "endDateTime", "startDateTime"].sort(),
    );
  });

  it("builds filter, orderby, skip, select and top as data; filter need not repeat orderby", async () => {
    await call("list-calendar-events", {
      top: 10,
      select: ["id", "subject", "attendees"],
      filter: "contains(subject,'review') and isAllDay eq false",
      orderby: [{ property: "start/dateTime", direction: "desc" }],
      skip: 5,
    });
    const q = parsed().searchParams;
    expect(q.get("$top")).toBe("10");
    expect(q.get("$select")).toBe("id,subject,attendees");
    expect(q.get("$filter")).toBe("contains(subject,'review') and isAllDay eq false");
    expect(q.get("$orderby")).toBe("start/dateTime desc");
    expect(q.get("$skip")).toBe("5");
    expect(q.has("$search")).toBe(false);
  });

  it("windowed tools accept filter and orderby but no skip", async () => {
    await call("get-calendar-view", {
      ...WIN,
      orderby: [{ property: "start/dateTime" }],
      filter: "contains(subject,'x')",
    });
    expect(parsed().searchParams.get("$orderby")).toBe("start/dateTime");
    calls = [];
    const { env } = await call("get-calendar-view", { ...WIN, skip: 1 });
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("allows createdDateTime/lastModifiedDateTime on /events list and get (not calendarView)", async () => {
    for (const [name, args] of [
      ["list-calendar-events", { select: ["id", "createdDateTime", "lastModifiedDateTime"] }],
      ["get-calendar-event", { "event-id": EVT_ID, select: ["id", "lastModifiedDateTime"] }],
    ] as Array<[string, Record<string, unknown>]>) {
      calls = [];
      responses.push(() => json(name.startsWith("list") ? { value: [] } : { id: "e" }));
      const { env } = await call(name, args);
      expect(env.ok, name).toBe(true);
      expect(new URL((calls[0] as Call).url).searchParams.get("$select")).toContain("DateTime");
    }
  });

  it("rejects out-of-contract arguments with invalid_input and zero network calls", async () => {
    const bad: Array<[string, Record<string, unknown>]> = [
      ["list-calendar-events", { top: 501 }],
      ["list-calendar-events", { top: 0 }],
      ["list-calendar-events", { select: ["body"] }],
      // calendarView documents that these two are not $select-able; instances: conservative.
      ["get-calendar-view", { ...WIN, select: ["createdDateTime"] }],
      ["get-calendar-view", { ...WIN, select: ["lastModifiedDateTime"] }],
      ["list-calendar-events", { orderby: [{ property: "body" }] }],
      ["list-calendar-events", { orderby: [{ property: "start" }] }],
      ["list-calendar-events", { filter: "x".repeat(1001) }],
      ["list-calendar-events", { filter: "a\nb" }],
      ["list-calendar-events", { search: "x" }],
      ["list-calendar-events", { expand: "attachments" }],
      ["list-calendar-events", { expandExtendedProperties: true }],
      ["list-calendar-events", { startDateTime: START, endDateTime: END }],
      ["list-calendar-events", { url: "/me" }],
      ["list-calendar-events", { nextLink: "https://graph.microsoft.com/v1.0/me/events" }],
      ["list-calendar-events", { bodyType: "html" }],
      ["list-calendar-events", { skip: 10_001 }],
      ["get-calendar-view", { ...WIN, top: 501 }],
      ["get-calendar-view", { ...WIN, select: ["body"] }],
      ["list-calendars", { top: 1001 }],
      ["list-calendars", { skip: 0 }],
      ["list-calendars", { expand: "events" }],
      ["list-calendars", { extendedPropertyId: "String 0x4001" }],
      ["list-calendars", { timeZone: "UTC" }],
      ["list-calendars", { select: ["events"] }],
      ["list-calendars", { filter: "name eq 'x'" }],
      ["get-calendar-event", { "event-id": EVT_ID, top: 5 }],
      ["get-calendar-event", { "event-id": EVT_ID, bodyType: "markdown" }],
      ["get-calendar-event", { "event-id": EVT_ID, filter: "isAllDay eq true" }],
    ];
    for (const [name, args] of bad) {
      const { env, text } = await call(name, args);
      expect(env.ok, `${name} ${JSON.stringify(args)}`).toBe(false);
      expect(env.error?.code, `${name} ${JSON.stringify(args)}`).toBe("invalid_input");
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });

  it("never echoes caller-chosen argument names", async () => {
    const { text } = await call("list-calendar-events", { "pasted-secret-key": 1 });
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
    "instances",
    "events/AAA=",
    "id()",
    "a:b",
    "AAA\\BBB",
    "microsoft.graph.foo",
    "?$expand=x",
    "#frag",
  ];

  it("rejects hostile calendar-id and event-id values before any request", async () => {
    for (const bad of hostile) {
      for (const [name, args] of [
        ["list-specific-calendar-events", { "calendar-id": bad }],
        ["get-specific-calendar-event", { "calendar-id": bad, "event-id": EVT_ID }],
        ["get-specific-calendar-event", { "calendar-id": CAL_ID, "event-id": bad }],
        ["get-calendar-event", { "event-id": bad }],
        ["get-specific-calendar-view", { "calendar-id": bad, ...WIN }],
        ["list-calendar-event-instances", { "calendar-id": bad, "event-id": EVT_ID, ...WIN }],
        ["list-calendar-event-instances", { "calendar-id": CAL_ID, "event-id": bad, ...WIN }],
      ] as const) {
        const { env, text } = await call(name, { ...args });
        expect(env.error?.code, `${name} ${JSON.stringify(bad)}`).toBe("invalid_input");
        expect(text).not.toContain(TOKEN);
      }
    }
    expect(calls).toHaveLength(0);
  });

  it("accepts Exchange ids ('=', '-', '_') and keeps '=' literal; '+' is encoded", async () => {
    await call("get-specific-calendar-event", { "calendar-id": CAL_ID, "event-id": EVT_ID });
    expect(parsed().pathname).toBe(`/v1.0/me/calendars/${CAL_ID}/events/${EVT_ID}`);
    expect(calls[0]?.url).not.toContain("%3D");
    calls = [];
    await call("get-calendar-event", { "event-id": "AAMk+AG_x-y=" });
    expect(parsed().pathname).toBe("/v1.0/me/events/AAMk%2BAG_x-y=");
  });

  it("requires every path parameter", async () => {
    for (const [name, args] of [
      ["get-calendar-event", {}],
      ["get-specific-calendar-event", { "calendar-id": CAL_ID }],
      ["list-specific-calendar-events", {}],
      ["list-calendar-event-instances", { "calendar-id": CAL_ID, ...WIN }],
    ] as const) {
      const r = await mcp.callTool({ name, arguments: { ...args } });
      expect(r.isError, name).toBe(true);
    }
    expect(calls).toHaveLength(0);
  });
});

describe("start/end window validation", () => {
  const view = (args: Record<string, unknown>) => call("get-calendar-view", args);
  const instances = (args: Record<string, unknown>) =>
    call("list-calendar-event-instances", {
      "calendar-id": CAL_ID,
      "event-id": EVT_ID,
      ...args,
    });
  const MS = 86_400_000;
  const plus = (days: number, extraSeconds = 0): string =>
    new Date(Date.parse(START) + days * MS + extraSeconds * 1000).toISOString();

  const badWindows: Array<[string, Record<string, unknown>]> = [
    ["missing both", {}],
    ["missing start", { endDateTime: END }],
    ["missing end", { startDateTime: START }],
    ["empty start", { startDateTime: "", endDateTime: END }],
    ["not a date", { startDateTime: "tomorrow", endDateTime: END }],
    ["date only", { startDateTime: "2024-01-01", endDateTime: END }],
    ["no zone (local)", { startDateTime: "2024-01-01T00:00:00", endDateTime: END }],
    ["offset beyond +14:00", { startDateTime: "2024-01-01T00:00:00+14:01", endDateTime: END }],
    ["offset +15:00", { startDateTime: "2024-01-01T00:00:00+15:00", endDateTime: END }],
    ["offset minutes 60", { startDateTime: "2024-01-01T00:00:00+01:60", endDateTime: END }],
    ["offset beyond -14:00", { startDateTime: START, endDateTime: "2024-01-08T00:00:00-14:30" }],
    ["offset without colon", { startDateTime: "2024-01-01T00:00:00+0100", endDateTime: END }],
    ["offset hours only", { startDateTime: "2024-01-01T00:00:00+01", endDateTime: END }],
    ["offset with seconds", { startDateTime: "2024-01-01T00:00:00+01:00:00", endDateTime: END }],
    ["one-digit offset hour", { startDateTime: "2024-01-01T00:00:00+1:00", endDateTime: END }],
    ["unknown-offset -00:00", { startDateTime: "2024-01-01T00:00:00-00:00", endDateTime: END }],
    ["space before offset", { startDateTime: "2024-01-01T00:00:00 +01:00", endDateTime: END }],
    ["space instead of T", { startDateTime: "2024-01-01 00:00:00Z", endDateTime: END }],
    ["Z and offset", { startDateTime: "2024-01-01T00:00:00Z+01:00", endDateTime: END }],
    ["local fraction only", { startDateTime: "2024-01-01T00:00:00.5", endDateTime: END }],
    [
      "year before 0100 after conversion",
      { startDateTime: "0100-01-01T00:00:00+14:00", endDateTime: "0100-02-01T00:00:00Z" },
    ],
    [
      "year after 9999 after conversion",
      { startDateTime: "9999-12-01T00:00:00Z", endDateTime: "9999-12-31T23:59:59-14:00" },
    ],
    [
      "reversed in UTC although later on the wall clock",
      { startDateTime: "2024-01-01T00:00:00+01:00", endDateTime: "2024-01-01T00:30:00+02:00" },
    ],
    ["lowercase z", { startDateTime: "2024-01-01T00:00:00z", endDateTime: END }],
    [
      "impossible date",
      { startDateTime: "2024-02-30T00:00:00Z", endDateTime: "2024-03-05T00:00:00Z" },
    ],
    ["month 13", { startDateTime: "2024-13-01T00:00:00Z", endDateTime: "2025-01-05T00:00:00Z" }],
    ["hour 24", { startDateTime: "2024-01-01T24:00:00Z", endDateTime: END }],
    [
      "non-leap Feb 29",
      { startDateTime: "2023-02-29T00:00:00Z", endDateTime: "2023-03-05T00:00:00Z" },
    ],
    ["year 24", { startDateTime: "0024-01-01T00:00:00Z", endDateTime: END }],
    ["injected query", { startDateTime: `${START}&$expand=attachments`, endDateTime: END }],
    ["injected header", { startDateTime: `${START}\r\nX: y`, endDateTime: END }],
    ["reversed", { startDateTime: END, endDateTime: START }],
    ["zero length", { startDateTime: START, endDateTime: START }],
    ["too long by a second", { startDateTime: START, endDateTime: plus(MAX_WINDOW_DAYS, 1) }],
    [
      "too long by a tenth of a second",
      { startDateTime: START, endDateTime: plus(MAX_WINDOW_DAYS).replace(".000Z", ".100Z") },
    ],
    [
      "reversed within one second",
      { startDateTime: "2024-01-01T00:00:00.2Z", endDateTime: "2024-01-01T00:00:00.1Z" },
    ],
    ["a decade", { startDateTime: START, endDateTime: "2034-01-01T00:00:00Z" }],
    ["number", { startDateTime: 1704067200, endDateTime: END }],
  ];

  it.each(badWindows)("rejects %s with zero network calls", async (_label, args) => {
    for (const run of [view, instances]) {
      const { env, text } = await run(args);
      expect(env.error?.code).toBe("invalid_input");
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });

  it("accepts the maximum window exactly, fractional seconds and a one-second window", async () => {
    for (const [w, wire] of [
      [
        { startDateTime: START, endDateTime: plus(MAX_WINDOW_DAYS) },
        [START, "2025-01-01T00:00:00Z"],
      ],
      [
        { startDateTime: "2024-01-01T00:00:00.0000000Z", endDateTime: "2024-01-01T00:00:01Z" },
        // Canonical form: an all-zero fraction is dropped.
        ["2024-01-01T00:00:00Z", "2024-01-01T00:00:01Z"],
      ],
      [
        { startDateTime: "2024-01-01T00:00:00.1Z", endDateTime: "2024-01-01T00:00:00.2Z" },
        ["2024-01-01T00:00:00.1Z", "2024-01-01T00:00:00.2Z"],
      ],
    ] as const) {
      for (const run of [view, instances]) {
        calls = [];
        const { env } = await run(w);
        expect(env.ok, JSON.stringify(w)).toBe(true);
        expect(parsed().searchParams.get("startDateTime")).toBe(wire[0]);
        expect(parsed().searchParams.get("endDateTime")).toBe(wire[1]);
      }
    }
  });

  it("states the window bound in the tool description of the inputs", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = tools.find((t) => t.name === "get-calendar-view")?.inputSchema
      .properties as Record<string, { description?: string }>;
    expect(props["startDateTime"]?.description).toContain(String(MAX_WINDOW_DAYS));
    expect(props["startDateTime"]?.description).toContain("project safety limit");
  });
});

describe("timeZone", () => {
  it.each([
    ["Europe/Berlin", "Europe/Berlin"],
    ["America/New_York", "America/New_York"],
    ["Asia/Kolkata", "Asia/Kolkata"],
    ["Europe/Kyiv", "Europe/Kyiv"],
    ["UTC", "UTC"],
    ["utc", "UTC"],
    ["europe/london", "Europe/London"],
    ["Pacific Standard Time", "Pacific Standard Time"],
    ["pacific standard time", "Pacific Standard Time"],
    ["yukon standard time", "Yukon Standard Time"],
    ["kamchatka standard time", "Kamchatka Standard Time"],
    ["Central Standard Time (Mexico)", "Central Standard Time (Mexico)"],
    ["Turkey Standard Time", "Turkey Standard Time"],
    ["Etc/GMT", "Etc/GMT"],
    ["asia/yangon", "Asia/Yangon"],
    ["Asia/Toshkent", "Asia/Toshkent"],
    ["Asia/Astana", "Asia/Astana"],
    ["asia/tashkent", "Asia/Tashkent"],
    ["Asia/Rangoon", "Asia/Rangoon"],
    ["Asia/Almaty", "Asia/Almaty"],
    ["UTC+12", "UTC+12"],
  ])("sends %s as the canonical Prefer %s on a list and a view", async (input, canonical) => {
    for (const [name, args] of [
      ["list-calendar-events", {}],
      ["get-calendar-view", { ...WIN }],
    ] as const) {
      calls = [];
      const { env } = await call(name, { ...args, timeZone: input });
      expect(env.ok, `${name} ${input}`).toBe(true);
      expect(prefer()).toBe(`outlook.timezone="${canonical}"`);
    }
  });

  it("does not change the window: start/end stay UTC wire values", async () => {
    await call("get-calendar-view", { ...WIN, timeZone: "Pacific Standard Time" });
    expect(parsed().searchParams.get("startDateTime")).toBe(START);
    expect(parsed().searchParams.get("endDateTime")).toBe(END);
  });

  it("sends no Prefer when timeZone is not chosen (Graph default UTC)", async () => {
    await call("list-calendar-events");
    expect(prefer()).toBeUndefined();
  });

  it.each(WINDOWED)("$name accepts timeZone", async (t) => {
    const { env } = await call(t.name, { ...t.params, timeZone: "Asia/Tokyo" });
    expect(env.ok).toBe(true);
    expect(prefer()).toBe('outlook.timezone="Asia/Tokyo"');
  });

  it("rejects unknown, malformed and injected zones with zero network calls", async () => {
    const bad = [
      "",
      " ",
      "Mars/Olympus_Mons",
      // Valid IANA names that the runtime knows but Microsoft does not document for calendar.
      "Europe/Zurich",
      "Europe/Madrid",
      "America/Toronto",
      "Australia/Melbourne",
      "Europe/Berlin ",
      'UTC"',
      'UTC", outlook.allow-unsafe-html',
      'Europe/Berlin", outlook.body-content-type="html',
      "UTC, outlook.allow-unsafe-html",
      "UTC\r\nX-Evil: 1",
      "UTC\u0000",
      "Europe/Berlin;x=y",
      "../../etc/passwd",
      "Türkiye Standard Time",
      "a".repeat(65),
    ];
    for (const timeZone of bad) {
      for (const [name, args] of [
        ["list-calendar-events", {}],
        ["get-calendar-view", { ...WIN }],
        ["get-calendar-event", { "event-id": EVT_ID }],
      ] as const) {
        const { env, text } = await call(name, { ...args, timeZone });
        expect(env.error?.code, `${name} ${JSON.stringify(timeZone)}`).toBe("invalid_input");
        expect(text).not.toContain(TOKEN);
      }
    }
    expect(calls).toHaveLength(0);
  });

  it("rejects non-string zones", async () => {
    for (const timeZone of [1, true, null, ["UTC"], { a: 1 }]) {
      const { env } = await call("list-calendar-events", { timeZone });
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("keeps the same Prefer on every followed page", async () => {
    const link = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${START}&endDateTime=${END}&%24skiptoken=abc123`;
    responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": link }));
    responses.push(() => json({ value: [{ id: "2" }] }));
    const { env } = await call("get-calendar-view", {
      ...WIN,
      top: 150,
      timeZone: "Europe/Berlin",
    });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(2);
    expect(prefer(0)).toBe('outlook.timezone="Europe/Berlin"');
    expect(prefer(1)).toBe('outlook.timezone="Europe/Berlin"');
    expect(parsed(1).pathname).toBe("/v1.0/me/calendarView");
    expect(parsed(1).searchParams.get("startDateTime")).toBe(START);
  });
});

describe("get-*-event", () => {
  it.each(EVENT_GETS)("$name: body, default text, honours html, default select", async (t) => {
    responses.push(() =>
      json({ id: "e", subject: "s", body: { contentType: "text", content: "hello" } }),
    );
    const { env } = await call(t.name, { ...t.params });
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe(t.path);
    expect(prefer()).toBe('outlook.body-content-type="text"');
    const select = parsed().searchParams.get("$select")?.split(",") ?? [];
    expect(select).toEqual(expect.arrayContaining(["body", "attendees", "start", "end"]));
    expect(JSON.stringify(env.data)).toContain("hello");
    calls = [];
    await call(t.name, { ...t.params, bodyType: "html" });
    expect(prefer()).toBe('outlook.body-content-type="html"');
  });

  it("combines timeZone and bodyType in one Prefer, timezone first", async () => {
    await call("get-calendar-event", {
      "event-id": EVT_ID,
      timeZone: "Pacific Standard Time",
      bodyType: "html",
    });
    expect(calls[0]?.headers["Prefer"]).toBe(
      'outlook.timezone="Pacific Standard Time", outlook.body-content-type="html"',
    );
    calls = [];
    await call("get-calendar-event", { "event-id": EVT_ID, timeZone: "UTC" });
    expect(prefer()).toBe('outlook.timezone="UTC", outlook.body-content-type="text"');
    // Exactly one Prefer header is sent, never an unsafe-html preference.
    expect(Object.keys(calls[0]?.headers ?? {}).filter((h) => /^prefer$/i.test(h))).toHaveLength(1);
    expect(prefer()).not.toContain("unsafe");
  });

  it("maps a Graph 404 to an envelope error without leaking the token", async () => {
    responses.push(() => json({ error: { code: "ErrorItemNotFound", message: "gone" } }, 404));
    const { env, text } = await call("get-calendar-event", { "event-id": EVT_ID });
    expect(env.ok).toBe(false);
    expect(text).not.toContain(TOKEN);
  });
});

describe("list-calendars", () => {
  it("lists calendars at the fixed path with a bounded default", async () => {
    responses.push(() =>
      json({ value: [{ id: "c1", name: "Calendar", isDefaultCalendar: true, x: 1 }] }),
    );
    const { env, text } = await call("list-calendars");
    expect(env.ok).toBe(true);
    expect(parsed().pathname).toBe("/v1.0/me/calendars");
    expect(parsed().searchParams.get("$top")).toBe("25");
    expect(parsed().searchParams.get("$select")).toContain("isDefaultCalendar");
    expect(prefer()).toBeUndefined();
    expect(text).not.toContain('"x"');
    expect(text).not.toContain(TOKEN);
  });
});

describe("list projection", () => {
  it("drops body, attachments and unknown properties on the first and later pages", async () => {
    const junk = { body: { content: "SECRET-BODY" }, attachments: [1], "@odata.id": "u" };
    const link = `https://graph.microsoft.com/v1.0/me/events?%24skiptoken=abc123`;
    responses.push(() =>
      json({
        value: [{ id: "1", subject: "a", start: { dateTime: "x" }, ...junk }],
        "@odata.nextLink": link,
      }),
    );
    responses.push(() => json({ value: [{ id: "2", subject: "b", ...junk }] }));
    const { env, text } = await call("list-calendar-events", { top: 150 });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(2);
    expect(text).not.toContain("SECRET-BODY");
    expect(text).not.toContain("attachments");
    expect(text).not.toContain("odata.id");
    expect((env.data as unknown[]).length).toBe(2);
    expect(env.meta?.["hasMore"]).toBe(false);
  });

  it("reports hasMore and never returns a Graph link", async () => {
    responses.push(() =>
      json({
        value: [{ id: "1" }, { id: "2" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/events?%24skiptoken=abc123",
      }),
    );
    const { env, text } = await call("list-calendar-events", { top: 2 });
    expect(env.meta?.["hasMore"]).toBe(true);
    expect(text).not.toContain("skiptoken");
    expect(calls).toHaveLength(1);
  });

  it("refuses a nextLink that leaves the endpoint or the calendar", async () => {
    const other = "AAMkOTHERCALENDAR123=";
    for (const link of [
      "https://evil.example/v1.0/me/calendarView?%24skiptoken=abc123",
      "https://graph.microsoft.com/v1.0/me/messages?%24skiptoken=abc123",
      `https://graph.microsoft.com/v1.0/me/calendars/${other}/calendarView?%24skiptoken=abc123`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [{ id: "1" }], "@odata.nextLink": link }));
      const { env, text } = await call("get-specific-calendar-view", {
        "calendar-id": CAL_ID,
        ...WIN,
        top: 50,
      });
      expect(env.ok, link).toBe(false);
      expect(calls).toHaveLength(1);
      expect(text).not.toContain(TOKEN);
    }
  });
});

describe("graph-batch reachability", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const win = `startDateTime=${START}&endDateTime=${END}`;
  const sel = "$select=id,subject,start,end";
  const okBatch = () =>
    responses.push(() => json({ responses: [{ id: "1", status: 200, body: { value: [] } }] }));

  const VALID: Array<[string, Record<string, string>?]> = [
    [`/me/events?${sel}`],
    [`/me/events?${sel}&$top=100&$orderby=start/dateTime%20desc&$skip=3`],
    [`/me/events?${sel}&$filter=contains(subject,'x')&$orderby=start/dateTime`],
    [`/me/events?${sel}`, { Prefer: 'outlook.timezone="Europe/Berlin"' }],
    [`/me/events/${EVT_ID}`, { Prefer: 'outlook.body-content-type="text"' }],
    [
      `/me/events/${EVT_ID}?$select=id,body`,
      { Prefer: 'outlook.timezone="UTC", outlook.body-content-type="html"' },
    ],
    [`/me/calendars/${CAL_ID}/events/${EVT_ID}`, { Prefer: 'outlook.body-content-type="text"' }],
    [`/me/calendars/${CAL_ID}/events?${sel}`],
    [`/me/calendarView?${win}&${sel}`],
    [
      `/me/calendarView?${win}&${sel}&$top=100`,
      { Prefer: 'outlook.timezone="Pacific Standard Time"' },
    ],
    [`/me/calendars/${CAL_ID}/calendarView?${win}&${sel}&$filter=isAllDay%20eq%20false`],
    [`/me/calendars/${CAL_ID}/events/${EVT_ID}/instances?${win}&${sel}`],
    [`/me/calendars`],
    [`/me/calendars?$top=100&$select=id,name`],
  ];

  it("accepts in-policy subrequests for every calendar tool", async () => {
    for (const [url, headers] of VALID) {
      calls = [];
      okBatch();
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(true);
      expect(calls).toHaveLength(1);
      expect(calls[0]?.url).toContain("/$batch");
      const sent = JSON.parse(calls[0]?.body ?? "{}") as {
        requests: Array<{ method: string; headers?: Record<string, string> }>;
      };
      expect(sent.requests[0]?.method).toBe("GET");
    }
  });

  it("rejects subrequests that violate the contracts, with zero network calls", async () => {
    const wEnd = `endDateTime=${END}`;
    const wStart = `startDateTime=${START}`;
    const bad: Array<[string, Record<string, string>?]> = [
      // $select required where Graph would otherwise return bodies
      [`/me/events`],
      [`/me/events?$top=10`],
      [`/me/calendarView?${win}`],
      [`/me/calendars/${CAL_ID}/events/${EVT_ID}/instances?${win}`],
      // missing / partial / invalid window
      [`/me/calendarView?${sel}`],
      [`/me/calendarView?${wStart}&${sel}`],
      [`/me/calendarView?${wEnd}&${sel}`],
      [`/me/calendars/${CAL_ID}/calendarView?${sel}`],
      [`/me/calendars/${CAL_ID}/events/${EVT_ID}/instances?${sel}`],
      [`/me/calendarView?startDateTime=&${wEnd}&${sel}`],
      [`/me/calendarView?startDateTime=2024-01-01&${wEnd}&${sel}`],
      [`/me/calendarView?startDateTime=2024-01-01T00:00:00&${wEnd}&${sel}`],
      [`/me/calendarView?startDateTime=2024-01-01T00:00:00%2B01:00&${wEnd}&${sel}`],
      [
        `/me/calendarView?startDateTime=2024-02-30T00:00:00Z&endDateTime=2024-03-05T00:00:00Z&${sel}`,
      ],
      [`/me/calendarView?startDateTime=${END}&endDateTime=${START}&${sel}`],
      [`/me/calendarView?startDateTime=${START}&endDateTime=${START}&${sel}`],
      [`/me/calendarView?startDateTime=${START}&endDateTime=2034-01-01T00:00:00Z&${sel}`],
      [`/me/calendarView?${win}&${win}&${sel}`],
      [`/me/calendarView?${win}&${sel}&startdatetime=${START}`],
      [`/me/calendarView?${win}&${sel}&$top=101`],
      [`/me/calendarView?${win}&${sel}&$skip=1`],
      [`/me/calendarView?${win}&${sel}&$expand=attachments`],
      [`/me/calendarView?${win}&${sel}&$search=%22x%22`],
      [`/me/calendarView?${win}&$select=id,body`],
      [`/me/calendarView?${win}&$select=id,createdDateTime`],
      [`/me/calendarView?${win}&${sel}&$orderby=body`],
      // window parameters are not part of the other event tools
      [`/me/events?${sel}&${win}`],
      [`/me/calendars/${CAL_ID}/events?${sel}&${wStart}`],
      [`/me/events/${EVT_ID}?${win}`, { Prefer: 'outlook.body-content-type="text"' }],
      [`/me/calendars?${win}`],
      // other limits
      [`/me/events?${sel}&$top=101`],
      [`/me/events?${sel}&$skip=10001`],
      [`/me/calendars?$top=101`],
      [`/me/calendars?$select=events`],
    ];
    for (const [url, headers] of bad) {
      const { env, text } = await batch(url, headers);
      expect(env.error?.code, `${url} ${JSON.stringify(headers)}`).toBe("invalid_input");
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });

  it("accepts only canonical Prefer headers (single zone, optional body type, fixed order)", async () => {
    const view = `/me/calendarView?${win}&${sel}`;
    const get = `/me/events/${EVT_ID}`;
    const bad: Array<[string, Record<string, string>]> = [
      [view, { Prefer: "outlook.allow-unsafe-html" }],
      [view, { Prefer: 'outlook.timezone="UTC", outlook.allow-unsafe-html' }],
      [view, { Prefer: 'outlook.allow-unsafe-html, outlook.timezone="UTC"' }],
      [view, { Prefer: 'outlook.timezone="Mars/Base"' }],
      [view, { Prefer: 'outlook.timezone="europe/berlin"' }], // not the canonical spelling
      [view, { Prefer: "outlook.timezone=UTC" }],
      [view, { Prefer: 'outlook.timezone="UTC",outlook.body-content-type="text"' }],
      [view, { Prefer: 'outlook.body-content-type="text"' }], // views do not return bodies
      [view, { Prefer: 'outlook.timezone="UTC"; foo=bar' }],
      [view, { Prefer: 'outlook.timezone="UTC", outlook.timezone="UTC"' }],
      [view, { Prefer: 'outlook.timezone="UTC", ' }],
      [view, { Prefer: "" }],
      [view, { Prefer: 'OUTLOOK.TIMEZONE="UTC"' }],
      [view, { Prefer: "odata.maxpagesize=500" }],
      [view, { ConsistencyLevel: "eventual" }],
      [`/me/events?${sel}`, { Prefer: 'outlook.timezone="UTC\\", x' }],
      [`/me/calendars`, { Prefer: 'outlook.timezone="UTC"' }],
      [`/me/calendars`, { Prefer: "outlook.allow-unsafe-html" }],
      [get, {}], // body-content-type is required for get
      [get, { Prefer: 'outlook.timezone="UTC"' }],
      [get, { Prefer: "outlook.allow-unsafe-html" }],
      [get, { Prefer: 'outlook.body-content-type="TEXT"' }],
      [get, { Prefer: 'outlook.body-content-type="text", outlook.timezone="UTC"' }], // wrong order
      [get, { Prefer: 'outlook.timezone="UTC", outlook.body-content-type="markdown"' }],
      [
        get,
        {
          Prefer:
            'outlook.timezone="UTC", outlook.body-content-type="text", outlook.allow-unsafe-html',
        },
      ],
      [get, { Prefer: 'outlook.body-content-type="text", outlook.body-content-type="html"' }],
      [get, { ConsistencyLevel: "eventual", Prefer: 'outlook.body-content-type="text"' }],
    ];
    for (const [url, headers] of bad) {
      const { env, text } = await batch(url, headers);
      expect(env.error?.code, `${url} ${JSON.stringify(headers)}`).toBe("invalid_input");
      expect(text).not.toContain(TOKEN);
    }
    expect(calls).toHaveLength(0);
  });

  it("projects event collections: no bodies, clipped to $top, links dropped", async () => {
    responses.push(() =>
      json({
        responses: [
          {
            id: "1",
            status: 200,
            body: {
              "@odata.context": "ctx",
              "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendarView?%24skiptoken=x1",
              value: [
                { id: "a", subject: "s", body: { content: "SECRET-BODY" }, attachments: [1] },
                { id: "b", subject: "t" },
                { id: "c", subject: "u" },
              ],
            },
          },
        ],
      }),
    );
    const { env, text } = await batch(`/me/calendarView?${win}&${sel}&$top=2`);
    expect(env.ok).toBe(true);
    expect(text).not.toContain("SECRET-BODY");
    expect(text).not.toContain("attachments");
    expect(text).not.toContain("skiptoken");
    expect(text).not.toContain("odata.context");
    expect(text).toContain("hasMore");
    expect(text).not.toContain('"u"');
  });

  it("does not echo rejected values", async () => {
    const { text } = await batch(`/me/calendarView?${win}&${sel}&$top=PASTED-SECRET`);
    expect(text).not.toContain("PASTED-SECRET");
    const r = await batch(`/me/calendarView?${win}&${sel}`, {
      Prefer: 'outlook.timezone="PASTED-ZONE"',
    });
    expect(r.text).not.toContain("PASTED-ZONE");
    expect(calls).toHaveLength(0);
  });
});

describe("registration", () => {
  it("registers exactly the eight calendar tools with a batch policy, before graph-batch", () => {
    const handlers = createHandlers(
      new GraphClient({ tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) }, catalog }),
      catalog,
    );
    const names = [...handlers.keys()];
    // SM-8 tools (shared calendar, calendarView delta) are covered in calendar-shared-delta.test.ts.
    const sm8 = new Set([
      "list-shared-calendar-events",
      "get-shared-calendar-view",
      "list-calendar-view-delta",
      "list-calendar-events-delta",
      // SM-10 part B (group calendar) lives in group-calendar.test.ts.
      "get-group-calendar-view",
    ]);
    expect(
      names.filter((n) => /calendar/.test(n) && n !== "graph-batch" && !sm8.has(n)).sort(),
    ).toEqual([...ALL_NAMES].sort());
    for (const n of ALL_NAMES) {
      expect(handlers.get(n)?.batchPolicy?.check, n).toBeTypeOf("function");
      expect(names.indexOf(n)).toBeLessThan(names.indexOf("graph-batch"));
    }
  });

  it("uses exactly the catalog GET paths and Calendars.Read", () => {
    for (const t of ALL_EVENT_TOOLS) {
      const e = catalog.find((x) => x.toolName === t.name);
      expect(e?.method, t.name).toBe("get");
      expect(e?.scopes, t.name).toEqual(["Calendars.Read"]);
    }
  });
});

describe("window offsets (tool boundary) and canonical UTC output", () => {
  const view = (args: Record<string, unknown>) => call("get-calendar-view", args);
  const instances = (args: Record<string, unknown>) =>
    call("list-calendar-event-instances", { "calendar-id": CAL_ID, "event-id": EVT_ID, ...args });

  it.each([
    ["+01:00 start", "2024-01-01T09:00:00+01:00", END, "2024-01-01T08:00:00Z", END],
    [
      "-08:00 both",
      "2024-01-01T00:00:00-08:00",
      "2024-01-08T00:00:00-08:00",
      "2024-01-01T08:00:00Z",
      "2024-01-08T08:00:00Z",
    ],
    ["+05:30", "2024-01-01T00:00:00+05:30", END, "2023-12-31T18:30:00Z", END],
    ["+14:00 limit", "2024-01-01T00:00:00+14:00", END, "2023-12-31T10:00:00Z", END],
    ["-14:00 limit", START, "2024-01-08T00:00:00-14:00", START, "2024-01-08T14:00:00Z"],
    ["+00:00 is UTC", "2024-01-01T00:00:00+00:00", END, START, END],
    [
      "fraction kept exactly",
      "2024-01-01T00:00:00.1234567+01:00",
      END,
      "2023-12-31T23:00:00.1234567Z",
      END,
    ],
    ["zero fraction dropped", "2024-01-01T00:00:00.0000000Z", END, START, END],
    ["trailing zeros trimmed", "2024-01-01T00:00:00.5000000Z", END, "2024-01-01T00:00:00.5Z", END],
    [
      "offset crosses the date",
      "2024-03-01T00:30:00+01:00",
      "2024-03-08T00:00:00Z",
      "2024-02-29T23:30:00Z",
      "2024-03-08T00:00:00Z",
    ],
  ])("%s", async (_label, start, end, wireStart, wireEnd) => {
    for (const run of [view, instances]) {
      calls = [];
      const { env } = await run({ startDateTime: start, endDateTime: end });
      expect(env.ok).toBe(true);
      expect(parsed().searchParams.get("startDateTime")).toBe(wireStart);
      expect(parsed().searchParams.get("endDateTime")).toBe(wireEnd);
      // Nothing but canonical UTC reaches the URL: no offset, no '+', no '%2B'.
      expect(calls[0]?.url).not.toMatch(/%2B|\+/);
    }
  });

  it("measures the 366-day bound on the absolute instants, across offsets", async () => {
    // +14:00 .. -10:00 are 366 days apart in UTC (2023-12-31T10:00Z .. 2024-12-31T10:00Z).
    const ok = await view({
      startDateTime: "2024-01-01T00:00:00+14:00",
      endDateTime: "2024-12-31T00:00:00-10:00",
    });
    expect(ok.env.ok).toBe(true);
    expect(parsed().searchParams.get("endDateTime")).toBe("2024-12-31T10:00:00Z");
    calls = [];
    const over = await view({
      startDateTime: "2024-01-01T00:00:00+14:00",
      endDateTime: "2024-12-31T00:00:01-10:00",
    });
    expect(over.env.error?.code).toBe("invalid_input");
    const wide = await view({
      startDateTime: "2024-01-01T00:00:00+14:00",
      endDateTime: "2025-01-01T00:00:00-14:00",
    });
    expect(wide.env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("states the accepted forms in the tool input descriptions", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = tools.find((t) => t.name === "get-calendar-view")?.inputSchema
      .properties as Record<string, { description?: string }>;
    expect(props["endDateTime"]?.description).toContain("offset");
    expect(props["endDateTime"]?.description).toContain("sent to Graph as the equivalent UTC");
  });
});

describe("batch window form: canonical UTC only", () => {
  const sel = "$select=id,subject,start,end";
  const batch = (path: string) =>
    call("graph-batch", { requests: [{ id: "1", method: "GET", url: path }] });
  const okBatch = () =>
    responses.push(() => json({ responses: [{ id: "1", status: 200, body: { value: [] } }] }));

  it("sends exactly the canonical UTC value the single path emits", async () => {
    okBatch();
    const w = "startDateTime=2024-01-01T00:00:00.5Z&endDateTime=2024-01-08T00:00:00Z";
    const { env } = await batch(`/me/calendarView?${w}&${sel}`);
    expect(env.ok).toBe(true);
    const sent = JSON.parse(calls[0]?.body ?? "{}") as { requests: Array<{ url: string }> };
    const sentUrl = decodeURIComponent(sent.requests[0]?.url ?? "");
    expect(sentUrl).toContain("startDateTime=2024-01-01T00:00:00.5Z");
    expect(sentUrl).toContain("endDateTime=2024-01-08T00:00:00Z");
  });

  it("rejects offsets (raw or %2B-encoded) and non-canonical spellings, with zero calls", async () => {
    const end = `endDateTime=${END}`;
    const bad = [
      "2024-01-01T00:00:00%2B01:00",
      "2024-01-01T00:00:00%2b00:00",
      "2024-01-01T00:00:00+01:00", // raw '+' is ambiguous: refused by the URL layer
      "2024-01-01T00:00:00-08:00",
      "2024-01-01T00:00:00.0000000Z",
      "2024-01-01T00:00:00.50Z",
      "2024-01-01T00:00:00.000Z",
      "2024-01-01T00:00:00",
    ];
    for (const value of bad) {
      for (const query of [
        `startDateTime=${value}&${end}`,
        `startDateTime=${START}&endDateTime=${value}`,
      ]) {
        const { env, text } = await batch(`/me/calendarView?${query}&${sel}`);
        expect(env.error?.code, query).toBe("invalid_input");
        expect(text).not.toContain(TOKEN);
      }
    }
    expect(calls).toHaveLength(0);
  });

  it("still enforces order and the 366-day bound on canonical values", async () => {
    for (const w of [
      `startDateTime=${END}&endDateTime=${START}`,
      `startDateTime=${START}&endDateTime=2025-01-01T00:00:01Z`,
    ]) {
      const { env } = await batch(`/me/calendarView?${w}&${sel}`);
      expect(env.error?.code).toBe("invalid_input");
    }
    okBatch();
    const exact = await batch(
      `/me/calendarView?startDateTime=${START}&endDateTime=2025-01-01T00:00:00Z&${sel}`,
    );
    expect(exact.env.ok).toBe(true);
  });
});

describe("extended properties ($expand)", () => {
  const G = "{8ECCC264-6880-4EBE-992F-8888D2EEAA1D}";
  const VALID_IDS = [
    `String ${G} Name TestProperty`,
    `Integer ${G} Id 0x8012`,
    "String 0x4001",
    "CLSID {00062008-0000-0000-C000-000000000046} Name ConnectorSenderGuid",
    "String {f17e0000-0000-4000-8000-000000000007} Name Color",
    `Boolean ${G} Name My Prop.v2:x-y_z`,
    `SystemTime ${G} Id 0x80ab`,
    "Binary 0x0FFF",
  ];
  const expandOf = (id: string): string => `singleValueExtendedProperties($filter=id eq '${id}')`;
  const EXT_TOOLS = [...EVENT_LISTS, ...WINDOWED, ...EVENT_GETS];

  it("is offered on exactly the tools the catalog flags with supportsExpandExtendedProperties", async () => {
    const raw = JSON.parse(
      readFileSync(new URL("../../endpoints.json", import.meta.url), "utf8"),
    ) as Array<{ toolName: string; supportsExpandExtendedProperties?: boolean }>;
    const flagged = raw
      .filter((e) => e.supportsExpandExtendedProperties === true)
      .map((e) => e.toolName)
      .filter((n) => ALL_NAMES.includes(n))
      .sort();
    expect(flagged).toHaveLength(7);
    const tools = (await mcp.listTools()).tools;
    const offered = ALL_NAMES.filter((n) =>
      Object.keys(tools.find((t) => t.name === n)?.inputSchema.properties ?? {}).includes(
        "extendedPropertyId",
      ),
    ).sort();
    expect(offered).toEqual(flagged);
    expect(offered).not.toContain("list-calendars");
    for (const t of tools.filter(
      (x) =>
        x.name !== "get-my-profile" &&
        x.name !== "list-chats" &&
        x.name !== "list-pinned-chat-messages" &&
        x.name !== "list-channel-tabs" &&
        x.name !== "list-sharepoint-site-list-items" &&
        x.name !== "get-sharepoint-site-list-item",
    )) {
      // Only the typed profile relationship `expand` (SM-9 part A) and the typed
      // `lastMessagePreview` expand of list-chats (SM-11 part A) the pinned-message `message` expand (part B) and the channel-tab `teamsApp` expand (SM-12 part C) and the list-item `fields` expand (SM-13 part C) exist; no other tool has it.
      expect(Object.keys(t.inputSchema.properties ?? {})).not.toContain("expand");
    }
  });

  it("documents the Calendars.Read requirement in the input description", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = tools.find((t) => t.name === "get-calendar-view")?.inputSchema
      .properties as Record<string, { description?: string }>;
    expect(props["extendedPropertyId"]?.description).toContain("Calendars.Read");
  });

  it.each(EXT_TOOLS)(
    "$name sends exactly the filtered expansion for each documented id form",
    async (t) => {
      for (const id of VALID_IDS) {
        calls = [];
        const { env } = await call(t.name, { ...t.params, extendedPropertyId: id });
        expect(env.ok, `${t.name} ${id}`).toBe(true);
        const q = parsed().searchParams;
        expect(q.get("$expand")).toBe(expandOf(id));
        expect(q.getAll("$expand")).toHaveLength(1);
        expect(calls[0]?.url).toContain(
          "$expand=singleValueExtendedProperties(%24filter%3Did%20eq%20'",
        );
      }
    },
  );

  it("sends no $expand unless an id is given", async () => {
    await call("get-calendar-view", { ...WIN });
    expect(parsed().searchParams.has("$expand")).toBe(false);
  });

  it("rejects hostile and malformed ids with zero network calls", async () => {
    const bad: unknown[] = [
      "",
      " ",
      `String ${G} Name x' or id eq 'y`,
      `String ${G} Name a'b`,
      "String 0x4001' or id eq 'x",
      `String ${G} Name a)`,
      `String ${G} Name a(b`,
      `String ${G} Name $filter`,
      `String ${G} Name a/b`,
      `String ${G} Name a\\b`,
      `String ${G} Name a&$select=body`,
      `String ${G} Name a"b`,
      `String ${G} Name a,b`,
      `String ${G} Name a;b`,
      `String ${G} Name x\r\nX-Evil: 1`,
      `String ${G} Name x\u0000`,
      `String ${G} Name ${"a".repeat(200)}`,
      `String ${G} Name ${"a".repeat(129)}`,
      `String ${G} Name `,
      `String ${G} Name  lead`,
      `String ${G} Name trail `,
      `String ${G} Name`,
      `Frobnicate ${G} Name x`, // unknown type
      "Object 0x4001",
      "Error 0x4001",
      "Null 0x4001",
      "StringArray 0x4001", // multi-value is not offered
      `IntegerArray ${G} Name x`,
      "string 0x4001", // type is exact case
      `String ${G} name x`, // keyword is exact case
      `String ${G} ID 0x8012`,
      `String ${G.slice(1, -1)} Name x`, // GUID must be braced
      "String {8ECCC264-6880-4EBE-992F-8888D2EEAA1} Name x", // short GUID
      "String {8ECCC264-6880-4EBE-992F-8888D2EEAA1G} Name x",
      `String ${G} Id 8012`,
      `String ${G} Id 0x`,
      `String ${G} Id 0x123456789`,
      `String ${G} Id 0xZZ`,
      "String 0x",
      "String 4001",
      "String 0x4001 ",
      " String 0x4001",
      "String  0x4001",
      `String ${G}`,
      "String 0x4001, String 0x4002",
      "singleValueExtendedProperties",
      "attachments",
      "String 0x4001 ".repeat(30),
      5,
      true,
      ["String 0x4001"],
      { id: "String 0x4001" },
      null,
    ];
    for (const id of bad) {
      for (const t of [EVENT_LISTS[0], WINDOWED[0], EVENT_GETS[0]]) {
        const { env, text } = await call(t.name, { ...t.params, extendedPropertyId: id });
        expect(env.error?.code, `${t.name} ${JSON.stringify(id)}`).toBe("invalid_input");
        expect(text).not.toContain(TOKEN);
      }
    }
    expect(calls).toHaveLength(0);
  });

  it("never offers a free-form expand, $expand or expandExtendedProperties input", async () => {
    for (const args of [
      { expand: "attachments" },
      { $expand: "attachments" },
      { expandExtendedProperties: true },
      { extendedPropertyId: "String 0x4001", expand: "attachments" },
    ]) {
      const { env } = await call("list-calendar-events", args);
      expect(env.error?.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("keeps only id/value of expanded properties (lists on every page, and gets)", async () => {
    const link = "https://graph.microsoft.com/v1.0/me/events?%24skiptoken=abc123";
    const props = [
      { id: "String 0x4001", value: "v1", "@odata.type": "x", extra: 1 },
      { id: "String 0x4002" },
      { value: "orphan" },
      "junk",
      null,
    ];
    responses.push(() =>
      json({
        value: [{ id: "1", subject: "a", singleValueExtendedProperties: props, body: "SECRET" }],
        "@odata.nextLink": link,
      }),
    );
    responses.push(() =>
      json({ value: [{ id: "2", singleValueExtendedProperties: "not-an-array" }] }),
    );
    const list = await call("list-calendar-events", {
      top: 150,
      extendedPropertyId: "String 0x4001",
    });
    expect(list.env.ok).toBe(true);
    expect(list.env.data).toEqual([
      {
        id: "1",
        subject: "a",
        singleValueExtendedProperties: [{ id: "String 0x4001", value: "v1" }],
      },
      { id: "2" },
    ]);
    expect(list.text).not.toContain("SECRET");
    expect(calls).toHaveLength(2);

    responses.push(() => json({ id: "e", subject: "s", singleValueExtendedProperties: props }));
    const get = await call("get-calendar-event", {
      "event-id": EVT_ID,
      extendedPropertyId: "String 0x4001",
    });
    expect(get.env.data).toEqual({
      id: "e",
      subject: "s",
      singleValueExtendedProperties: [{ id: "String 0x4001", value: "v1" }],
    });
  });

  it("drops singleValueExtendedProperties from tools that do not offer the feature", async () => {
    responses.push(() =>
      json({ value: [{ id: "1", singleValueExtendedProperties: [{ id: "x", value: "y" }] }] }),
    );
    const { env } = await call("list-calendars");
    expect(JSON.stringify(env.data)).not.toContain("singleValueExtendedProperties");
  });

  describe("in graph-batch", () => {
    const sel = "$select=id,subject,start,end";
    const win = `startDateTime=${START}&endDateTime=${END}`;
    const exp = (id: string) => `$expand=${encodeURIComponent(expandOf(id))}`;
    const batch = (url: string, headers?: Record<string, string>) =>
      call("graph-batch", {
        requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
      });
    const okBatch = (body: unknown = { value: [] }) =>
      responses.push(() => json({ responses: [{ id: "1", status: 200, body }] }));
    const BODY_TEXT = { Prefer: 'outlook.body-content-type="text"' };

    it("accepts only the canonical expansion on every flagged tool", async () => {
      for (const id of VALID_IDS) {
        const cases: Array<[string, Record<string, string> | undefined]> = [
          [`/me/events?${sel}&${exp(id)}`, undefined],
          [`/me/calendars/${CAL_ID}/events?${sel}&${exp(id)}`, undefined],
          [`/me/calendarView?${win}&${sel}&${exp(id)}`, undefined],
          [`/me/calendars/${CAL_ID}/calendarView?${win}&${sel}&${exp(id)}`, undefined],
          [
            `/me/calendars/${CAL_ID}/events/${EVT_ID}/instances?${win}&${sel}&${exp(id)}`,
            undefined,
          ],
          [`/me/events/${EVT_ID}?${exp(id)}`, BODY_TEXT],
          [`/me/calendars/${CAL_ID}/events/${EVT_ID}?${exp(id)}`, BODY_TEXT],
        ];
        for (const [url, headers] of cases) {
          calls = [];
          okBatch();
          const { env } = await batch(url, headers);
          expect(env.ok, url).toBe(true);
          expect(calls).toHaveLength(1);
        }
      }
    });

    it("rejects every other $expand, with zero network calls", async () => {
      const id = "String 0x4001";
      const e = encodeURIComponent;
      const bad = [
        `$expand=attachments`,
        `$expand=singleValueExtendedProperties`,
        `$expand=${e("singleValueExtendedProperties($filter=id eq 'x')")}`,
        `$expand=${e(`${expandOf(id)},attachments`)}`,
        `$expand=${e(`attachments,${expandOf(id)}`)}`,
        `$expand=${e(`singleValueExtendedProperties($filter=id eq '${id}' or id eq 'String 0x4002')`)}`,
        `$expand=${e(`singleValueExtendedProperties($filter=id eq '${id}' and value eq 'x')`)}`,
        `$expand=${e(`singleValueExtendedProperties($select=id;$filter=id eq '${id}')`)}`,
        `$expand=${e(`multiValueExtendedProperties($filter=id eq 'StringArray 0x4001')`)}`,
        `$expand=${e(`singleValueExtendedProperties($filter=id eq "${id}")`)}`,
        `$expand=${e(`singleValueExtendedProperties($filter=id  eq '${id}')`)}`,
        `$expand=${e(`singleValueExtendedProperties($filter=Id eq '${id}')`)}`,
        `$expand=${e(`singleValueExtendedProperties($filter=id eq '${id}''x')`)}`,
        `$expand=${e(`singleValueExtendedProperties($filter=id eq ''))`)}`,
        `$expand=${e(`${expandOf(id)} `)}`,
        `$expand=${e(`singleValueExtendedProperties/$ref`)}`,
        `$expand=*`,
        `$expand=`,
        `$Expand=${e(expandOf(id))}`,
        `expand=${e(expandOf(id))}`,
        `${exp(id)}&${exp(id)}`,
      ];
      for (const q of bad) {
        const { env, text } = await batch(`/me/events?${sel}&${q}`);
        expect(env.error?.code, q).toBe("invalid_input");
        expect(text).not.toContain(TOKEN);
      }
      // A hostile id inside the otherwise canonical shape.
      for (const hostile of [
        "String 0x4001' or id eq 'x",
        "Foo 0x4001",
        `String ${G} Name a/b`,
        "StringArray 0x4001",
      ]) {
        const { env } = await batch(`/me/events?${sel}&${exp(hostile)}`);
        expect(env.error?.code, hostile).toBe("invalid_input");
      }
      // list-calendars does not offer it, and get keeps its other rules.
      const { env } = await batch(`/me/calendars?${exp(id)}`);
      expect(env.error?.code).toBe("invalid_input");
      const noPrefer = await batch(`/me/events/${EVT_ID}?${exp(id)}`);
      expect(noPrefer.env.error?.code).toBe("invalid_input");
      expect(calls).toHaveLength(0);
    });

    it("projects expanded properties to id/value in batch collections and entities", async () => {
      const props = [{ id: "String 0x4001", value: "v", extra: { a: 1 } }, "junk", { id: 5 }];
      okBatch({
        value: [
          { id: "1", body: { content: "SECRET-BODY" }, singleValueExtendedProperties: props },
        ],
      });
      const list = await batch(`/me/events?${sel}&${exp("String 0x4001")}`);
      expect(list.text).not.toContain("SECRET-BODY");
      expect(list.text).not.toContain("extra");
      expect(list.text).toContain('\\"value\\":\\"v\\"');
      calls = [];
      okBatch({ id: "e", singleValueExtendedProperties: props, "@odata.context": "ctx" });
      const get = await batch(`/me/events/${EVT_ID}?${exp("String 0x4001")}`, BODY_TEXT);
      expect(get.env.ok).toBe(true);
      expect(get.text).not.toContain("extra");
      expect(get.text).toContain('\\"value\\":\\"v\\"');
      expect(get.text).not.toContain("junk");
    });
  });
});

describe("list-calendars beyond one page", () => {
  const page = (n: number, from: number, next?: string) =>
    json({
      value: Array.from({ length: n }, (_, i) => ({ id: `c${from + i}`, name: `n${from + i}` })),
      ...(next === undefined ? {} : { "@odata.nextLink": next }),
    });
  const nextLink = (n: number) =>
    `https://graph.microsoft.com/v1.0/me/calendars?%24top=100&%24skip=${n * 100}`;

  it("accepts top up to 1000 and sends the project page size 100", async () => {
    await call("list-calendars", { top: 1000 });
    expect(parsed().searchParams.get("$top")).toBe("100");
    calls = [];
    await call("list-calendars", { top: 60 });
    expect(parsed().searchParams.get("$top")).toBe("60");
    calls = [];
    await call("list-calendars", { top: 150 });
    expect(parsed().searchParams.get("$top")).toBe("100");
    expect(parsed().searchParams.has("$skip")).toBe(false);
  });

  it("follows validated nextLinks beyond 100 items", async () => {
    responses.push(() => page(100, 0, nextLink(1)));
    responses.push(() => page(100, 100, nextLink(2)));
    responses.push(() => page(37, 200));
    const { env } = await call("list-calendars", { top: 1000 });
    expect(env.ok).toBe(true);
    expect((env.data as unknown[]).length).toBe(237);
    expect(env.meta?.["hasMore"]).toBe(false);
    expect(calls).toHaveLength(3);
    expect(parsed(1).pathname).toBe("/v1.0/me/calendars");
    expect(parsed(2).searchParams.get("$skip")).toBe("200");
  });

  it("stops at the requested top with hasMore, never exceeding it", async () => {
    responses.push(() => page(100, 0, nextLink(1)));
    responses.push(() => page(100, 100, nextLink(2)));
    const { env } = await call("list-calendars", { top: 150 });
    expect((env.data as unknown[]).length).toBe(150);
    expect(env.meta?.["hasMore"]).toBe(true);
    expect(calls).toHaveLength(2);
  });

  it("stops at the client ceiling (10 pages / 1000 items) with hasMore", async () => {
    for (let i = 0; i < 12; i += 1) responses.push(() => page(100, i * 100, nextLink(i + 1)));
    const { env } = await call("list-calendars", { top: 1000 });
    expect((env.data as unknown[]).length).toBe(1000);
    expect(env.meta?.["hasMore"]).toBe(true);
    expect(calls.length).toBeLessThanOrEqual(10);
    expect(env.meta?.["pages"]).toBeLessThanOrEqual(10);
  });

  it("refuses a nextLink that leaves the endpoint", async () => {
    responses.push(() => page(1, 0, "https://graph.microsoft.com/v1.0/me/messages?%24skip=100"));
    const { env, text } = await call("list-calendars", { top: 500 });
    expect(env.ok).toBe(false);
    expect(text).not.toContain(TOKEN);
    expect(calls).toHaveLength(1);
  });

  it("offers no skip or expand input, and batch stays bound to the page size", async () => {
    const tools = (await mcp.listTools()).tools;
    const keys = Object.keys(
      tools.find((t) => t.name === "list-calendars")?.inputSchema.properties ?? {},
    );
    expect(keys.sort()).toEqual(["select", "top"]);
    for (const url of [
      "/me/calendars?$top=101",
      "/me/calendars?$skip=100",
      "/me/calendars?$top=1000",
    ]) {
      const { env } = await call("graph-batch", { requests: [{ id: "1", method: "GET", url }] });
      expect(env.error?.code, url).toBe("invalid_input");
    }
    responses.push(() =>
      json({ responses: [{ id: "1", status: 200, body: { value: [{ id: "c" }] } }] }),
    );
    const ok = await call("graph-batch", {
      requests: [{ id: "1", method: "GET", url: "/me/calendars?$top=100" }],
    });
    expect(ok.env.ok).toBe(true);
    expect(calls).toHaveLength(1);
  });
});
