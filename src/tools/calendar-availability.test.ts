import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike, type GraphRequest } from "../graph/client.js";
import { GraphClientError } from "../graph/errors.js";
import { createServer } from "../server.js";
import { MAX_ATTENDEES, MAX_SCHEDULES, MAX_TIME_SLOTS } from "./calendar-availability.js";
import { createHandlers } from "./index.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const G = "https://graph.microsoft.com";
const START = "2024-03-15T09:00:00Z";
const END = "2024-03-15T18:00:00Z";
const SCHED = { schedules: ["adelev@contoso.com"], startDateTime: START, endDateTime: END };

interface Call {
  url: string;
  method: string;
  headers: Record<string, string>;
  body: string | undefined;
}

let catalog: EndpointCatalog;
let mcp: Client;
let graph: GraphClient;
let calls: Call[];
let tokenCalls: number;
let responses: Array<() => Response>;

const json = (body: unknown, status = 200): Response =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

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
  graph = new GraphClient({
    tokenProvider: {
      getAccessToken: () => {
        tokenCalls += 1;
        return Promise.resolve(TOKEN);
      },
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
  tokenCalls = 0;
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
const sent = (i = 0): Record<string, unknown> =>
  JSON.parse((calls[i] as Call).body as string) as Record<string, unknown>;

async function expectInvalid(name: string, args: Record<string, unknown>): Promise<void> {
  const { r, env } = await call(name, args);
  expect(r.isError, JSON.stringify(args).slice(0, 200)).toBe(true);
  expect(env.error?.code).toBe("invalid_input");
  expect(calls).toHaveLength(0);
  expect(tokenCalls).toBe(0);
}

describe("registration", () => {
  it("implements both tools, neither exposes a proxy-like or body parameter", async () => {
    const tools = (await mcp.listTools()).tools;
    for (const name of ["get-schedule", "find-meeting-times"]) {
      const tool = tools.find((t) => t.name === name);
      expect(tool).toBeDefined();
      const props = JSON.stringify(tool?.inputSchema);
      expect(props).not.toMatch(/"(url|uri|method|path|endpoint|body)"\s*:/i);
      expect((tool?.inputSchema as { additionalProperties?: unknown }).additionalProperties).toBe(
        false,
      );
    }
  });
});

describe("get-schedule", () => {
  it("POSTs the exact path with a server-built body and no Prefer by default", async () => {
    responses.push(() =>
      json({
        value: [
          {
            scheduleId: "adelev@contoso.com",
            availabilityView: "0022",
            scheduleItems: [
              {
                isPrivate: false,
                status: "busy",
                subject: "Lunch",
                location: "Bar",
                start: { dateTime: "2024-03-15T12:00:00.0000000", timeZone: "UTC" },
                end: { dateTime: "2024-03-15T13:00:00.0000000", timeZone: "UTC" },
                sneaky: "dropped",
              },
            ],
            workingHours: {
              daysOfWeek: ["monday"],
              startTime: "08:00:00.0000000",
              endTime: "17:00:00.0000000",
              timeZone: { name: "Pacific Standard Time", bias: 480 },
            },
            extra: "dropped",
          },
          {
            scheduleId: "meganb@contoso.com",
            error: { message: "Not found", responseCode: "ErrorMailRecipientNotFound" },
          },
        ],
      }),
    );
    const { r, env, text } = await call("get-schedule", {
      ...SCHED,
      schedules: ["adelev@contoso.com", "meganb@contoso.com"],
      availabilityViewInterval: 60,
    });
    expect(r.isError).toBe(false);
    expect(calls).toHaveLength(1);
    const c = calls[0] as Call;
    expect(c.method).toBe("POST");
    expect(c.url).toBe(`${G}/v1.0/me/calendar/getSchedule`);
    expect(c.headers["Prefer"]).toBeUndefined();
    expect(c.headers["Content-Type"]).toBe("application/json");
    expect(sent()).toEqual({
      schedules: ["adelev@contoso.com", "meganb@contoso.com"],
      startTime: { dateTime: "2024-03-15T09:00:00", timeZone: "UTC" },
      endTime: { dateTime: "2024-03-15T18:00:00", timeZone: "UTC" },
      availabilityViewInterval: 60,
    });
    expect(env.data).toEqual([
      {
        scheduleId: "adelev@contoso.com",
        availabilityView: "0022",
        scheduleItems: [
          {
            status: "busy",
            isPrivate: false,
            subject: "Lunch",
            location: "Bar",
            start: { dateTime: "2024-03-15T12:00:00.0000000", timeZone: "UTC" },
            end: { dateTime: "2024-03-15T13:00:00.0000000", timeZone: "UTC" },
          },
        ],
        workingHours: {
          daysOfWeek: ["monday"],
          startTime: "08:00:00.0000000",
          endTime: "17:00:00.0000000",
          timeZone: { name: "Pacific Standard Time", bias: 480 },
        },
      },
      {
        scheduleId: "meganb@contoso.com",
        error: { message: "Not found", responseCode: "ErrorMailRecipientNotFound" },
      },
    ]);
    expect(env.meta?.["count"]).toBe(2);
    expect(text).not.toContain(TOKEN);
    expect(text).not.toContain("dropped");
  });

  it("converts offsets to UTC, omits the interval by default and sends the canonical Prefer", async () => {
    await call("get-schedule", {
      ...SCHED,
      startDateTime: "2024-03-15T10:00:00+01:00",
      endDateTime: "2024-03-15T19:00:00.50+01:00",
      timeZone: "pacific standard time",
    });
    expect(sent()).toEqual({
      schedules: ["adelev@contoso.com"],
      startTime: { dateTime: "2024-03-15T09:00:00", timeZone: "UTC" },
      endTime: { dateTime: "2024-03-15T18:00:00.5", timeZone: "UTC" },
    });
    expect((calls[0] as Call).headers["Prefer"]).toBe('outlook.timezone="Pacific Standard Time"');
  });

  it("surfaces a per-schedule error (200 + error) instead of dropping it", async () => {
    responses.push(() =>
      json({
        value: [
          {
            scheduleId: "ghost@contoso.com",
            scheduleItems: [],
            error: { message: "No mailbox", responseCode: "ErrorMailRecipientNotFound", x: 1 },
          },
        ],
      }),
    );
    const { env } = await call("get-schedule", { ...SCHED, schedules: ["ghost@contoso.com"] });
    expect(env.ok).toBe(true);
    expect(env.data).toEqual([
      {
        scheduleId: "ghost@contoso.com",
        scheduleItems: [],
        error: { message: "No mailbox", responseCode: "ErrorMailRecipientNotFound" },
      },
    ]);
  });

  it("maps Graph failures to envelope codes and rejects a non-array response", async () => {
    responses.push(() => json({ error: { code: "ErrorAccessDenied", message: "no" } }, 403));
    const denied = await call("get-schedule", SCHED);
    expect(denied.env.error?.code).toBe("capability_gap");
    responses.push(() => json({ value: "nope" }));
    const odd = await call("get-schedule", SCHED);
    expect(odd.env.error?.code).toBe("graph_error");
  });

  it("rejects hostile or malformed input before any token or network use", async () => {
    const many = (n: number) => Array.from({ length: n }, (_, i) => `u${i}@contoso.com`);
    const bad: Array<Record<string, unknown>> = [
      {},
      { ...SCHED, body: { schedules: ["x@y.com"] } },
      { ...SCHED, url: "https://evil.example" },
      { ...SCHED, method: "DELETE" },
      { ...SCHED, schedules: [] },
      { ...SCHED, schedules: many(MAX_SCHEDULES + 1) },
      { ...SCHED, schedules: ["not-an-email"] },
      { ...SCHED, schedules: ["a@b"] },
      { ...SCHED, schedules: ["a b@contoso.com"] },
      { ...SCHED, schedules: ['a"@contoso.com'] },
      { ...SCHED, schedules: ["a@contoso.com", "A@Contoso.com"] },
      { ...SCHED, schedules: [`${"a".repeat(65)}@contoso.com`] },
      { ...SCHED, schedules: [42] },
      { ...SCHED, startDateTime: "2024-03-15T09:00:00" },
      { ...SCHED, startDateTime: "yesterday" },
      { ...SCHED, endDateTime: "2024-02-30T09:00:00Z" },
      { ...SCHED, startDateTime: END, endDateTime: START },
      { ...SCHED, startDateTime: START, endDateTime: START },
      { ...SCHED, endDateTime: "2024-06-15T18:00:00Z" },
      { ...SCHED, availabilityViewInterval: 4 },
      { ...SCHED, availabilityViewInterval: 1441 },
      { ...SCHED, availabilityViewInterval: 30.5 },
      { ...SCHED, availabilityViewInterval: "30" },
      { ...SCHED, timeZone: "Mars/Olympus" },
      { ...SCHED, timeZone: 'UTC", outlook.body-content-type="html' },
      { ...SCHED, timeZone: "UTC\r\nX-Evil: 1" },
    ];
    for (const args of bad) await expectInvalid("get-schedule", args);
  });

  it("requires a window of LESS than 62 days (documented Graph limit)", async () => {
    await expectInvalid("get-schedule", { ...SCHED, endDateTime: "2024-05-16T09:00:00Z" });
    responses.push(() => json({ value: [{ scheduleId: "adelev@contoso.com" }] }));
    const { r } = await call("get-schedule", { ...SCHED, endDateTime: "2024-05-16T08:59:59Z" });
    expect(r.isError).toBe(false);
    expect(calls).toHaveLength(1);
  });

  it("returns a full-length availabilityView intact and fails closed on an over-long one", async () => {
    // 7 days at 5-minute slots = 2016 characters (longer than the generic text cap).
    const week = { ...SCHED, endDateTime: "2024-03-22T09:00:00Z", availabilityViewInterval: 5 };
    const view = "0123".repeat(504);
    responses.push(() =>
      json({ value: [{ scheduleId: "adelev@contoso.com", availabilityView: view }] }),
    );
    const ok = await call("get-schedule", week);
    expect(ok.env.data).toEqual([{ scheduleId: "adelev@contoso.com", availabilityView: view }]);
    responses.push(() =>
      json({ value: [{ scheduleId: "adelev@contoso.com", availabilityView: `${view}0` }] }),
    );
    expect((await call("get-schedule", week)).env.error?.code).toBe("graph_error");
    // A partial final slot counts, even when it is only 100 ns long.
    responses.push(() =>
      json({ value: [{ scheduleId: "adelev@contoso.com", availabilityView: "02" }] }),
    );
    const partial = await call("get-schedule", {
      ...SCHED,
      endDateTime: "2024-03-15T09:30:00.0000001Z",
    });
    expect(partial.env.ok).toBe(true);
    // Default interval (30 min): a 9-hour window allows 18 characters.
    responses.push(() =>
      json({ value: [{ scheduleId: "adelev@contoso.com", availabilityView: "0".repeat(19) }] }),
    );
    expect((await call("get-schedule", SCHED)).env.error?.code).toBe("graph_error");
  });

  it("fails closed on malformed or over-long schedule data instead of dropping it", async () => {
    const at = { dateTime: "2024-03-15T10:00:00.0000000", timeZone: "UTC" };
    const item = { status: "busy", start: at, end: at };
    const id = { scheduleId: "a@contoso.com" };
    const bodies: unknown[] = [
      { value: [null] },
      { value: [] },
      { value: [id, { scheduleId: "b@contoso.com" }] },
      { value: [{}] },
      { value: [{ scheduleId: "" }] },
      { value: [{ scheduleId: 42 }] },
      { value: [{ ...id, scheduleItems: [item, null] }] },
      { value: [{ ...id, scheduleItems: "busy" }] },
      { value: [{ ...id, scheduleItems: Array.from({ length: 1001 }, () => item) }] },
      { value: [{ ...id, scheduleItems: [{ status: "busy" }] }] },
      { value: [{ ...id, scheduleItems: [{ start: at, end: at }] }] },
      { value: [{ ...id, scheduleItems: [{ ...item, end: { dateTime: "x" } }] }] },
      { value: [{ ...id, scheduleItems: [{ ...item, start: { timeZone: "UTC" } }] }] },
      { value: [{ ...id, availabilityView: 2 }] },
      { value: [{ ...id, scheduleItems: [{ ...item, subject: 42 }] }] },
      { value: [{ ...id, scheduleItems: [{ ...item, subject: "s".repeat(2001) }] }] },
      { value: [{ ...id, scheduleItems: [{ ...item, isPrivate: "false" }] }] },
      { value: [{ ...id, workingHours: { startTime: 8 } }] },
      { value: [{ ...id, workingHours: { daysOfWeek: ["monday", 1] } }] },
      {},
    ];
    for (const body of bodies) {
      responses.push(() => json(body));
      const { r, env } = await call("get-schedule", SCHED);
      expect(r.isError, JSON.stringify(body).slice(0, 100)).toBe(true);
      expect(env.error?.code).toBe("graph_error");
    }
    responses.push(() =>
      json({ value: [{ ...id, scheduleItems: Array.from({ length: 1000 }, () => item) }] }),
    );
    expect((await call("get-schedule", SCHED)).env.ok).toBe(true);
  });

  it("keeps the offsets and transition rules of a custom working-hours time zone", async () => {
    const custom = {
      "@odata.type": "#microsoft.graph.customTimeZone",
      bias: 480,
      name: "Customized Time Zone",
      standardOffset: {
        time: "02:00:00.0000000",
        dayOccurrence: 1,
        dayOfWeek: "sunday",
        month: 11,
        year: 0,
        x: 1,
      },
      daylightOffset: {
        daylightBias: -60,
        time: "02:00:00.0000000",
        dayOccurrence: 2,
        dayOfWeek: "sunday",
        month: 3,
        year: 0,
      },
      sneaky: "dropped",
    };
    responses.push(() =>
      json({
        value: [
          {
            scheduleId: "a@contoso.com",
            workingHours: { startTime: "08:00:00.0000000", timeZone: custom },
          },
        ],
      }),
    );
    const { env, text } = await call("get-schedule", SCHED);
    expect(env.data).toEqual([
      {
        scheduleId: "a@contoso.com",
        workingHours: {
          startTime: "08:00:00.0000000",
          timeZone: {
            name: "Customized Time Zone",
            bias: 480,
            standardOffset: {
              time: "02:00:00.0000000",
              dayOccurrence: 1,
              dayOfWeek: "sunday",
              month: 11,
              year: 0,
            },
            daylightOffset: {
              time: "02:00:00.0000000",
              dayOccurrence: 2,
              dayOfWeek: "sunday",
              month: 3,
              year: 0,
              daylightBias: -60,
            },
          },
        },
      },
    ]);
    expect(text).not.toContain("dropped");
  });

  it("omits projected objects that have no fields", async () => {
    responses.push(() =>
      json({
        value: [
          {
            scheduleId: "a@contoso.com",
            error: { x: 1 },
            workingHours: { timeZone: {} },
            scheduleItems: [
              {
                status: "busy",
                start: { dateTime: "2024-03-15T10:00:00", timeZone: "UTC" },
                end: { dateTime: "2024-03-15T11:00:00", timeZone: "UTC", x: 1 },
                subject: null,
                location: null,
                isPrivate: null,
              },
            ],
          },
        ],
      }),
    );
    const { env } = await call("get-schedule", SCHED);
    expect(env.data).toEqual([
      {
        scheduleId: "a@contoso.com",
        scheduleItems: [
          {
            status: "busy",
            start: { dateTime: "2024-03-15T10:00:00", timeZone: "UTC" },
            end: { dateTime: "2024-03-15T11:00:00", timeZone: "UTC" },
          },
        ],
      },
    ]);
  });
});

describe("find-meeting-times", () => {
  it("POSTs the exact path with a server-built body, Prefer and a projected result", async () => {
    responses.push(() =>
      json({
        emptySuggestionsReason: "",
        meetingTimeSuggestions: [
          {
            confidence: 100,
            order: 1,
            organizerAvailability: "free",
            suggestionReason: "nearest",
            attendeeAvailability: [
              {
                availability: "free",
                attendee: { type: "required", emailAddress: { address: "alexw@contoso.com" } },
              },
            ],
            locations: [{ displayName: "Conf room", junk: 1 }],
            meetingTimeSlot: {
              start: { dateTime: "2024-04-18T16:00:00.0000000", timeZone: "Pacific Standard Time" },
              end: { dateTime: "2024-04-18T17:00:00.0000000", timeZone: "Pacific Standard Time" },
            },
            junk: "dropped",
          },
        ],
      }),
    );
    const { r, env, text } = await call("find-meeting-times", {
      attendees: [
        { address: "alexw@contoso.com" },
        { address: "room1@contoso.com", type: "resource" },
      ],
      timeSlots: [
        { startDateTime: "2024-04-16T09:00:00-07:00", endDateTime: "2024-04-18T17:00:00-07:00" },
      ],
      activityDomain: "work",
      meetingDuration: "PT1H30M",
      maxCandidates: 5,
      minimumAttendeePercentage: 80,
      isOrganizerOptional: false,
      returnSuggestionReasons: true,
      timeZone: "Pacific Standard Time",
    });
    expect(r.isError).toBe(false);
    const c = calls[0] as Call;
    expect(c.method).toBe("POST");
    expect(c.url).toBe(`${G}/v1.0/me/findMeetingTimes`);
    expect(c.headers["Prefer"]).toBe('outlook.timezone="Pacific Standard Time"');
    expect(sent()).toEqual({
      attendees: [
        { type: "required", emailAddress: { address: "alexw@contoso.com" } },
        { type: "resource", emailAddress: { address: "room1@contoso.com" } },
      ],
      timeConstraint: {
        activityDomain: "work",
        timeSlots: [
          {
            start: { dateTime: "2024-04-16T16:00:00", timeZone: "UTC" },
            end: { dateTime: "2024-04-19T00:00:00", timeZone: "UTC" },
          },
        ],
      },
      meetingDuration: "PT1H30M",
      maxCandidates: 5,
      minimumAttendeePercentage: 80,
      isOrganizerOptional: false,
      returnSuggestionReasons: true,
    });
    const data = env.data as { meetingTimeSuggestions: Array<Record<string, unknown>> };
    expect(data.meetingTimeSuggestions).toHaveLength(1);
    expect(data.meetingTimeSuggestions[0]).toEqual({
      order: 1,
      confidence: 100,
      organizerAvailability: "free",
      suggestionReason: "nearest",
      meetingTimeSlot: {
        start: { dateTime: "2024-04-18T16:00:00.0000000", timeZone: "Pacific Standard Time" },
        end: { dateTime: "2024-04-18T17:00:00.0000000", timeZone: "Pacific Standard Time" },
      },
      attendeeAvailability: [
        {
          availability: "free",
          attendee: { type: "required", emailAddress: { address: "alexw@contoso.com" } },
        },
      ],
      locations: [{ displayName: "Conf room" }],
    });
    expect(text).not.toContain("dropped");
    expect(text).not.toContain(TOKEN);
  });

  it("sends an empty body when no option is given (organizer-only defaults)", async () => {
    await call("find-meeting-times", {});
    expect(sent()).toEqual({});
    expect((calls[0] as Call).headers["Prefer"]).toBeUndefined();
  });

  it("returns empty suggestions with a reason as a normal success, not an error", async () => {
    responses.push(() =>
      json({ emptySuggestionsReason: "AttendeesUnavailable", meetingTimeSuggestions: [] }),
    );
    const { r, env } = await call("find-meeting-times", { meetingDuration: "PT30M" });
    expect(r.isError).toBe(false);
    expect(env.ok).toBe(true);
    expect(env.data).toEqual({
      emptySuggestionsReason: "AttendeesUnavailable",
      meetingTimeSuggestions: [],
    });
    expect(env.meta?.["count"]).toBe(0);
  });

  it("passes unknown reason values through", async () => {
    responses.push(() =>
      json({ emptySuggestionsReason: "SomeFutureReason", meetingTimeSuggestions: [] }),
    );
    const { env } = await call("find-meeting-times", {});
    expect(env.data).toEqual({
      emptySuggestionsReason: "SomeFutureReason",
      meetingTimeSuggestions: [],
    });
  });

  it("reports only the specific issue when a slot list already has a specific error", async () => {
    for (const timeSlots of [
      [],
      [{ startDateTime: "yesterday", endDateTime: END }],
      [
        { startDateTime: "2024-06-01T09:00:00Z", endDateTime: "2024-06-02T09:00:00Z" },
        { startDateTime: "2024-01-01T09:00:00Z", endDateTime: "2023-12-01T09:00:00Z" },
      ],
    ]) {
      const { env } = await call("find-meeting-times", { timeSlots });
      expect(env.error?.code).toBe("invalid_input");
      const fields = (env.error?.details as { fields: Record<string, string[]> }).fields;
      expect(fields["timeSlots"] ?? [], JSON.stringify(fields)).not.toContain("custom");
    }
    expect(calls).toHaveLength(0);
  });

  it("bounds each time slot and the aggregate search horizon to less than 62 days", async () => {
    await expectInvalid("find-meeting-times", {
      timeSlots: [{ startDateTime: START, endDateTime: "2024-05-16T09:00:00Z" }],
    });
    await expectInvalid("find-meeting-times", {
      timeSlots: [
        { startDateTime: START, endDateTime: "2024-03-16T09:00:00Z" },
        { startDateTime: "2024-05-15T09:00:00Z", endDateTime: "2024-05-16T09:00:00Z" },
      ],
    });
    responses.push(() => json({ meetingTimeSuggestions: [] }));
    const { r } = await call("find-meeting-times", {
      timeSlots: [
        { startDateTime: START, endDateTime: "2024-03-16T09:00:00Z" },
        { startDateTime: "2024-05-15T09:00:00Z", endDateTime: "2024-05-16T08:59:59Z" },
      ],
    });
    expect(r.isError).toBe(false);
  });

  it("fails closed when Graph returns more suggestions than maxCandidates", async () => {
    const at = { dateTime: "2024-03-15T10:00:00", timeZone: "UTC" };
    const slot = { meetingTimeSlot: { start: at, end: at } };
    const two = { meetingTimeSuggestions: [slot, slot] };
    responses.push(() => json(two));
    expect((await call("find-meeting-times", { maxCandidates: 1 })).env.error?.code).toBe(
      "graph_error",
    );
    responses.push(() => json(two));
    expect((await call("find-meeting-times", { maxCandidates: 2 })).env.ok).toBe(true);
  });

  it("fails closed on a malformed response instead of reporting no free time", async () => {
    const at = { dateTime: "2024-03-15T10:00:00", timeZone: "UTC" };
    const okSlot = { meetingTimeSlot: { start: at, end: at } };
    for (const body of [
      { emptySuggestionsReason: "SomeFutureReason" },
      { meetingTimeSuggestions: null },
      { meetingTimeSuggestions: {} },
      { meetingTimeSuggestions: [null] },
      { meetingTimeSuggestions: [{ order: 1 }, "x"] },
      { meetingTimeSuggestions: [{}] },
      { meetingTimeSuggestions: [{ ...okSlot, confidence: "high" }] },
      { emptySuggestionsReason: 7, meetingTimeSuggestions: [] },
      {
        meetingTimeSuggestions: [
          { meetingTimeSlot: { start: { dateTime: "x", timeZone: "UTC" } } },
        ],
      },
      {
        meetingTimeSuggestions: [{ ...okSlot, attendeeAvailability: [{}] }],
      },
      {
        meetingTimeSuggestions: [
          {
            ...okSlot,
            attendeeAvailability: [{ attendee: { emailAddress: { address: "a@contoso.com" } } }],
          },
        ],
      },
      {
        meetingTimeSuggestions: [
          {
            ...okSlot,
            attendeeAvailability: [{ availability: "free", attendee: { emailAddress: {} } }],
          },
        ],
      },
      { meetingTimeSuggestions: Array.from({ length: 101 }, () => ({})) },
      { meetingTimeSuggestions: [{ attendeeAvailability: [1] }] },
      { meetingTimeSuggestions: [{ locations: "Room 1" }] },
      [],
      "x",
    ]) {
      responses.push(() => json(body));
      const { r, env } = await call("find-meeting-times", {});
      expect(r.isError).toBe(true);
      expect(env.error?.code).toBe("graph_error");
    }
  });

  it("maps a Graph 403 to capability_gap", async () => {
    responses.push(() => json({ error: { code: "ErrorAccessDenied", message: "no" } }, 403));
    expect((await call("find-meeting-times", {})).env.error?.code).toBe("capability_gap");
  });

  it("rejects hostile or malformed input before any token or network use", async () => {
    const slot = { startDateTime: START, endDateTime: END };
    const att = { address: "a@contoso.com" };
    const bad: Array<Record<string, unknown>> = [
      { body: {} },
      { url: "/me/sendMail" },
      { attendees: [{ ...att, extra: 1 }] },
      { attendees: [{ emailAddress: { address: "a@contoso.com" } }] },
      { attendees: [{ address: "nope" }] },
      { attendees: [{ ...att, type: "organizer" }] },
      { attendees: [{ ...att, type: "optional" }] },
      {
        attendees: Array.from({ length: MAX_ATTENDEES + 1 }, (_, i) => ({
          address: `u${i}@contoso.com`,
        })),
      },
      { timeSlots: [] },
      { timeSlots: Array.from({ length: MAX_TIME_SLOTS + 1 }, () => slot) },
      { timeSlots: [{ ...slot, extra: 1 }] },
      { timeSlots: [{ startDateTime: "2024-03-15T09:00:00", endDateTime: END }] },
      { timeSlots: [{ startDateTime: END, endDateTime: START }] },
      { timeSlots: [{ startDateTime: START, endDateTime: "2025-03-15T18:00:00Z" }] },
      { activityDomain: "unknown" },
      { activityDomain: "everything" },
      { meetingDuration: "1h" },
      { meetingDuration: "PT" },
      { meetingDuration: "P1D" },
      { meetingDuration: "PT0M" },
      { meetingDuration: "PT25H" },
      { meetingDuration: "PT-5M" },
      { meetingDuration: "PT1H;x" },
      { maxCandidates: 0 },
      { maxCandidates: 1000 },
      { maxCandidates: 2.5 },
      { minimumAttendeePercentage: -1 },
      { minimumAttendeePercentage: 101 },
      { isOrganizerOptional: "true" },
      { returnSuggestionReasons: 1 },
      { locationConstraint: { isRequired: true } },
      { timeZone: "Nowhere Standard Time" },
      { timeZone: "UTC\r\nX-Evil: 1" },
    ];
    for (const args of bad) await expectInvalid("find-meeting-times", args);
  });
});

describe("POST allowlist and batch boundary (security)", () => {
  const client = (): GraphClient => graph;
  const fail = async (p: Promise<unknown>): Promise<GraphClientError> => {
    try {
      await p;
    } catch (e) {
      expect(e).toBeInstanceOf(GraphClientError);
      return e as GraphClientError;
    }
    throw new Error("expected failure");
  };

  it("client refuses POST to any path outside the catalog POST entries, with zero activity", async () => {
    const paths = [
      "/me/sendMail",
      "/me/events",
      "/me/messages",
      "/me/calendar/events",
      "/me/calendar/getSchedule/../x",
      "/me/calendar/getSchedule/",
      "/me/calendar/getSchedule?x=1",
      "/me/calendar/getSchedule#x",
      "/ME/calendar/getSchedule",
      "/me/calendar/GETSCHEDULE",
      "/me/calendar/getschedule",
      "/me/calendar%2FgetSchedule",
      "/me/calendar/get%53chedule",
      "/me/calendar/getSchedule%00",
      "/me/findmeetingtimes",
      "/me/findMeetingTimes/",
      "/me/findMeetingTimes%2e%2e/x",
      "/users/{user-id}/findMeetingTimes",
      "/users/x/calendar/getSchedule",
      "//me/findMeetingTimes",
      "https://graph.microsoft.com/v1.0/me/findMeetingTimes",
    ];
    for (const pathPattern of paths) {
      const err = await fail(client().request({ pathPattern, method: "post", body: {} }));
      expect(err.code, pathPattern).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
    expect(tokenCalls).toBe(0);
  });

  it("client refuses PATCH/PUT/DELETE against both availability paths", async () => {
    for (const pathPattern of ["/me/calendar/getSchedule", "/me/findMeetingTimes"]) {
      for (const method of ["patch", "put", "delete", "head", "options"]) {
        const req = { pathPattern, method: method as never, body: {} } as GraphRequest;
        expect((await fail(client().request(req))).code).toBe("invalid_input");
      }
      // GET to a POST-only entry is not catalogued either.
      expect((await fail(client().request({ pathPattern }))).code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
    expect(tokenCalls).toBe(0);
  });

  it("graph-batch rejects subrequests aimed at either availability endpoint (GET or POST)", async () => {
    const urls = [
      "/me/calendar/getSchedule",
      "/me/findMeetingTimes",
      "/v1.0/me/calendar/getSchedule",
      "/me/calendar/getSchedule?x=1",
      "/me/findMeetingTimes?$top=1",
    ];
    for (const url of urls) {
      for (const method of ["GET", "POST"]) {
        const { r, env } = await call("graph-batch", { requests: [{ id: "a", method, url }] });
        expect(r.isError, `${method} ${url}`).toBe(true);
        expect(env.error?.code).toBe("invalid_input");
      }
    }
    const withBody = await call("graph-batch", {
      requests: [{ id: "a", method: "POST", url: "/me/findMeetingTimes", body: {} }],
    });
    expect(withBody.env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
    expect(tokenCalls).toBe(0);
  });
});
