import { z } from "zod";
import type { EndpointCatalog } from "../catalog.js";
import { errorResult, successResult } from "../envelope.js";
import { GraphClientError } from "../graph/errors.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { canonicalInstant, instantTicks } from "./timestamps.js";
import { canonicalTimeZone } from "./timezones.js";

/**
 * Calendar availability tools (SM-8 part B): `get-schedule` and `find-meeting-times`.
 *
 * Both are retrieval-only POSTs (Graph requires POST for these actions). Safety properties:
 * - The input schemas are strict at every level; the Graph request body is CONSTRUCTED here from
 *   validated fields, never forwarded. There is no `body`, `url`, `method` or free-form JSON input.
 * - Neither tool declares a `batchPolicy`, so `graph-batch` can never reach them (batch is GET-only).
 * - The window/slot instants are accepted as ISO-8601 with `Z` or an explicit offset and always
 *   sent as UTC (`dateTime` without a zone suffix + `timeZone: "UTC"`). The optional `timeZone`
 *   only selects the zone of RETURNED times via `Prefer: outlook.timezone` (allowlisted, SM-7).
 * - Responses are untrusted: only an allowlisted set of typed fields is returned.
 *
 * Limits: Microsoft documents `availabilityViewInterval` 5..1440 (default 30),
 * `minimumAttendeePercentage` 0..100, and (free/busy concept page) getSchedule supports up to 20
 * entities with a lookup period of LESS THAN 62 days; a time slot with more than 1000 calendar
 * entries yields a per-schedule 5006 error instead of items. The attendee, time-slot, duration
 * and maxCandidates bounds are PROJECT safety limits; the docs state none. findMeetingTimes
 * slots reuse the getSchedule window rule for both each slot and the aggregate search horizon.
 *
 * Responses fail closed: a malformed entry or a list/view longer than the documented contract
 * allows is a `graph_error`, never silently dropped or truncated.
 *
 * Delegated permissions (current Microsoft docs; `endpoints.json` scope metadata is
 * non-normative): getSchedule needs Calendars.ReadBasic (endpoint-specific API reference; decision
 * D4, the catalog matches); findMeetingTimes needs Calendars.Read.Shared.
 * See PERMISSIONS.md.
 */

/** Graph limits for getSchedule (free/busy concept page). */
export const MAX_SCHEDULES = 20;
/** Exclusive: the lookup period must be less than this many days. */
export const MAX_AVAILABILITY_WINDOW_DAYS = 62;
/** Project safety limits (not Graph limits). */
export const MAX_ATTENDEES = 20;
export const MAX_TIME_SLOTS = 10;
export const MAX_CANDIDATES = 20;
export const MAX_DURATION_MINUTES = 24 * 60;
/** Output ceilings; exceeding one means Graph broke its documented contract (fail closed). */
const MAX_OUT_SCHEDULES = MAX_SCHEDULES;
const MAX_OUT_ITEMS = 1000;
const MAX_OUT_SUGGESTIONS = 100;
const MAX_OUT_ATTENDEES = 100;
const MAX_OUT_LOCATIONS = 100;
const MAX_TEXT = 2000;
const DEFAULT_VIEW_INTERVAL_MINUTES = 30;

const DAY_MS = 86_400_000;

/**
 * Conservative SMTP/UPN address check (RFC 5321 length limits). Covers guest UPNs
 * (`user_ext.com#EXT#@tenant.onmicrosoft.com`). The value only ever travels inside a JSON body
 * that is serialized by `JSON.stringify`, so this is validation, not an injection control.
 */
const EMAIL_RE =
  /^[A-Za-z0-9._%+'#=-]{1,64}@[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$/;

const email = (what: string) =>
  z
    .string()
    .max(254)
    .refine((v) => EMAIL_RE.test(v))
    .describe(what);

const instant = (what: string) =>
  z
    .string()
    .refine((v) => instantTicks(v) !== undefined)
    .describe(
      `${what} as an ISO-8601 timestamp with a Z or an explicit UTC offset, e.g. 2024-01-31T09:00:00Z or 2024-01-31T09:00:00+01:00 (local times without Z/offset are rejected). Sent to Graph as the equivalent UTC value.`,
    );

const timeZoneField = z
  .string()
  .min(1)
  .max(64)
  .refine((v) => canonicalTimeZone(v) !== undefined)
  .optional()
  .describe(
    'Time zone for the times in the RESULT: a Microsoft-documented IANA name (e.g. "Europe/Berlin"), a Windows name (e.g. "Pacific Standard Time") or "UTC". Default UTC. Does not change the requested window, which is always UTC.',
  );

/** True when `start < end` and the span is LESS THAN the documented 62-day limit. */
function windowOk(start: string, end: string): boolean {
  const s = instantTicks(start);
  const e = instantTicks(end);
  if (s === undefined || e === undefined || s >= e) return false;
  // 10_000 ticks per ms; compare in BigInt, no precision loss.
  return e - s < BigInt(MAX_AVAILABILITY_WINDOW_DAYS) * BigInt(DAY_MS) * 10_000n;
}

/** True when all candidate slots fit inside one bounded overall search horizon. */
function slotsHorizonOk(slots: Array<{ startDateTime: string; endDateTime: string }>): boolean {
  if (slots.length === 0) return false;
  let earliest: bigint | undefined;
  let latest: bigint | undefined;
  for (const slot of slots) {
    const start = instantTicks(slot.startDateTime);
    const end = instantTicks(slot.endDateTime);
    if (start === undefined || end === undefined) return false;
    if (earliest === undefined || start < earliest) earliest = start;
    if (latest === undefined || end > latest) latest = end;
  }
  return (
    (latest as bigint) - (earliest as bigint) <
    BigInt(MAX_AVAILABILITY_WINDOW_DAYS) * BigInt(DAY_MS) * 10_000n
  );
}

/**
 * Longest `availabilityView` a valid response can carry for this request: one character per
 * interval (a partial final interval counts as one). Only called after `windowOk` passed.
 */
function maxViewLength(start: string, end: string, interval: number): number {
  // Ceiling in 100 ns ticks, so a sub-millisecond partial final slot still counts.
  const span = (instantTicks(end) as bigint) - (instantTicks(start) as bigint);
  const slot = BigInt(interval) * 60_000n * 10_000n;
  return Number((span + slot - 1n) / slot);
}

/** Graph dateTimeTimeZone for a validated instant: canonical UTC without the `Z` suffix. */
function utcDateTime(value: string): { dateTime: string; timeZone: "UTC" } {
  const canonical = canonicalInstant(value) as string;
  return { dateTime: canonical.slice(0, -1), timeZone: "UTC" };
}

function preferHeaders(timeZone: unknown): Record<string, string> {
  if (typeof timeZone !== "string") return {};
  const name = canonicalTimeZone(timeZone);
  if (name === undefined) throw new GraphClientError("invalid_input", "Unsupported time zone.");
  return { Prefer: `outlook.timezone="${name}"` };
}

// ------------------------------------------------------------------ response projection

const isObj = (v: unknown): v is Record<string, unknown> =>
  typeof v === "object" && v !== null && !Array.isArray(v);
/**
 * Optional scalar projectors. Absent (`undefined` or Graph's `null`) stays absent; a present value
 * of the wrong type, or a string over MAX_TEXT, fails closed instead of being dropped or truncated.
 */
const absent = (v: unknown): v is undefined | null => v === undefined || v === null;
const text = (v: unknown): string | undefined => {
  if (absent(v)) return undefined;
  if (typeof v !== "string" || v.length > MAX_TEXT) malformed("text field");
  return v;
};
const num = (v: unknown): number | undefined => {
  if (absent(v)) return undefined;
  if (typeof v !== "number" || !Number.isFinite(v)) malformed("numeric field");
  return v;
};
const bool = (v: unknown): boolean | undefined => {
  if (absent(v)) return undefined;
  if (typeof v !== "boolean") malformed("boolean field");
  return v;
};

function put(out: Record<string, unknown>, key: string, value: unknown): void {
  if (value !== undefined) out[key] = value;
}

/** Puts `value` only when it has at least one projected field (no `{}` noise). */
function putObj(out: Record<string, unknown>, key: string, value: Record<string, unknown>): void {
  if (Object.values(value).some((v) => v !== undefined)) out[key] = value;
}

function malformed(what: string): never {
  throw new GraphClientError("graph_error", `Graph returned a malformed ${what}.`);
}

/**
 * Projects an optional response array. Absent stays absent; anything else must be an array of
 * at most `max` objects, otherwise the whole response fails closed (no silent drop/truncation).
 */
function objects<T>(
  v: unknown,
  max: number,
  what: string,
  project: (o: Record<string, unknown>) => T,
): T[] | undefined {
  if (v === undefined) return undefined;
  if (!Array.isArray(v) || v.length > max || !v.every(isObj)) malformed(what);
  return v.map(project);
}

/** A field that identifies or defines a record: a non-empty string, never truncated. */
function requiredText(v: unknown, what: string): string {
  if (typeof v !== "string" || v.length === 0 || v.length > MAX_TEXT) malformed(what);
  return v;
}

/** A complete dateTimeTimeZone (both parts present); anything else fails closed. */
function requiredDateTime(v: unknown, what: string): { dateTime: string; timeZone: string } {
  if (!isObj(v)) malformed(what);
  return {
    dateTime: requiredText(v["dateTime"], what),
    timeZone: requiredText(v["timeZone"], what),
  };
}

function daysOfWeek(v: unknown): string[] | undefined {
  if (v === undefined) return undefined;
  if (!Array.isArray(v) || v.length > 7 || !v.every((s) => typeof s === "string")) {
    malformed("working-hours day list");
  }
  return v.map((s: string) => s.slice(0, MAX_TEXT));
}

/** standardTimeZoneOffset / daylightTimeZoneOffset: the transition rule of a custom zone. */
function projectZoneOffset(v: unknown, daylight: boolean): Record<string, unknown> | undefined {
  if (!isObj(v)) return undefined;
  const out: Record<string, unknown> = {};
  put(out, "time", text(v["time"]));
  put(out, "dayOccurrence", num(v["dayOccurrence"]));
  put(out, "dayOfWeek", text(v["dayOfWeek"]));
  put(out, "month", num(v["month"]));
  put(out, "year", num(v["year"]));
  if (daylight) put(out, "daylightBias", num(v["daylightBias"]));
  return Object.keys(out).length === 0 ? undefined : out;
}

/**
 * Working-hours zone. Graph may return a `customTimeZone` whose `name` is only "Customized Time
 * Zone"; its bias and transition rules are kept (bounded, typed) so `startTime`/`endTime` stay
 * interpretable.
 */
function projectZone(tz: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  put(out, "name", text(tz["name"]));
  put(out, "bias", num(tz["bias"]));
  put(out, "standardOffset", projectZoneOffset(tz["standardOffset"], false));
  put(out, "daylightOffset", projectZoneOffset(tz["daylightOffset"], true));
  return out;
}

function projectWorkingHours(v: unknown): Record<string, unknown> | undefined {
  if (!isObj(v)) return undefined;
  const out: Record<string, unknown> = {};
  put(out, "daysOfWeek", daysOfWeek(v["daysOfWeek"]));
  put(out, "startTime", text(v["startTime"]));
  put(out, "endTime", text(v["endTime"]));
  const tz = v["timeZone"];
  if (isObj(tz)) putObj(out, "timeZone", projectZone(tz));
  return out;
}

function projectScheduleItem(v: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  out["status"] = requiredText(v["status"], "schedule item status");
  put(out, "isPrivate", bool(v["isPrivate"]));
  put(out, "subject", text(v["subject"]));
  put(out, "location", text(v["location"]));
  out["start"] = requiredDateTime(v["start"], "schedule item start");
  out["end"] = requiredDateTime(v["end"], "schedule item end");
  return out;
}

/**
 * One `scheduleInformation`. Per-schedule failures arrive as 200 + `error`; keep them visible.
 * `availabilityView` is data (one character per interval), so it is never truncated: a view
 * longer than the request window allows (`maxView`) fails closed. More than 1000 schedule items
 * is impossible per the docs (Graph answers with a 5006 error instead) and also fails closed.
 * Every entry must name its `scheduleId`, and every schedule item needs a status plus a
 * complete start and end.
 */
export function projectScheduleInformation(
  v: Record<string, unknown>,
  maxView: number,
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  out["scheduleId"] = requiredText(v["scheduleId"], "schedule id");
  const view = v["availabilityView"];
  if (view !== undefined) {
    if (typeof view !== "string" || view.length > maxView) malformed("availability view");
    out["availabilityView"] = view;
  }
  put(
    out,
    "scheduleItems",
    objects(v["scheduleItems"], MAX_OUT_ITEMS, "schedule item list", projectScheduleItem),
  );
  const hours = projectWorkingHours(v["workingHours"]);
  if (hours !== undefined) putObj(out, "workingHours", hours);
  const err = v["error"];
  if (isObj(err)) {
    putObj(out, "error", {
      message: text(err["message"]),
      responseCode: text(err["responseCode"]),
    });
  }
  return out;
}

/** An attendee's availability: requires the availability value and the attendee's address. */
function projectAttendeeAvailability(v: Record<string, unknown>): Record<string, unknown> {
  const attendee = v["attendee"];
  const ea = isObj(attendee) ? attendee["emailAddress"] : undefined;
  if (!isObj(attendee) || !isObj(ea)) malformed("attendee availability");
  const emailAddress: Record<string, unknown> = {
    address: requiredText(ea["address"], "attendee address"),
  };
  put(emailAddress, "name", text(ea["name"]));
  const a: Record<string, unknown> = {};
  put(a, "type", text(attendee["type"]));
  a["emailAddress"] = emailAddress;
  return { availability: requiredText(v["availability"], "attendee availability"), attendee: a };
}

function projectLocation(l: Record<string, unknown>): Record<string, unknown> {
  const o: Record<string, unknown> = {};
  put(o, "displayName", text(l["displayName"]));
  put(o, "locationEmailAddress", text(l["locationEmailAddress"]));
  return o;
}

/** A suggestion is only meaningful with a complete meeting slot; anything less fails closed. */
function projectSuggestion(v: Record<string, unknown>): Record<string, unknown> {
  const slot = v["meetingTimeSlot"];
  if (!isObj(slot)) malformed("meeting time slot");
  const out: Record<string, unknown> = {};
  put(out, "order", num(v["order"]));
  put(out, "confidence", num(v["confidence"]));
  put(out, "organizerAvailability", text(v["organizerAvailability"]));
  put(out, "suggestionReason", text(v["suggestionReason"]));
  out["meetingTimeSlot"] = {
    start: requiredDateTime(slot["start"], "meeting time slot start"),
    end: requiredDateTime(slot["end"], "meeting time slot end"),
  };
  put(
    out,
    "attendeeAvailability",
    objects(
      v["attendeeAvailability"],
      MAX_OUT_ATTENDEES,
      "attendee availability list",
      projectAttendeeAvailability,
    ),
  );
  put(
    out,
    "locations",
    objects(v["locations"], MAX_OUT_LOCATIONS, "location list", projectLocation),
  );
  return out;
}

/**
 * `meetingTimeSuggestionsResult`. An empty suggestion list plus a reason is a normal result, but
 * a response that is not an object with a `meetingTimeSuggestions` array of objects fails closed:
 * malformed data is never reported as "no free time". More suggestions than the request's
 * `maxCandidates` (or the local ceiling when omitted) also fails closed.
 */
export function projectMeetingTimes(
  v: unknown,
  maxSuggestions: number = MAX_OUT_SUGGESTIONS,
): Record<string, unknown> {
  if (!isObj(v)) malformed("meeting-times result");
  const list = objects(
    v["meetingTimeSuggestions"],
    maxSuggestions,
    "suggestion list",
    projectSuggestion,
  );
  if (list === undefined) malformed("meeting-times result");
  const out: Record<string, unknown> = {};
  put(out, "emptySuggestionsReason", text(v["emptySuggestionsReason"]));
  out["meetingTimeSuggestions"] = list;
  return out;
}

// ------------------------------------------------------------------ tools

const getSchedulePath = (catalog: EndpointCatalog): string => postPath(catalog, "get-schedule");
const findMeetingTimesPath = (catalog: EndpointCatalog): string =>
  postPath(catalog, "find-meeting-times");

export function postPath(catalog: EndpointCatalog, toolName: string): string {
  const entry = catalog.find((e) => e.toolName === toolName);
  if (entry === undefined || entry.method !== "post") {
    throw new Error(`Catalog has no POST entry for ${toolName}.`);
  }
  return entry.pathPattern;
}

function failure(toolName: string, err: unknown): ReturnType<typeof errorResult> {
  if (err instanceof GraphClientError) return err.toResult();
  return errorResult("graph_error", `Tool ${toolName} failed unexpectedly.`);
}

function windowIssue(ctx: z.RefinementCtx, path: Array<string | number>): void {
  ctx.addIssue({
    code: "custom",
    path,
    message: `start must be before end and the window less than ${MAX_AVAILABILITY_WINDOW_DAYS} days`,
  });
}

const GET_SCHEDULE_SHAPE = {
  schedules: z
    .array(email("SMTP address of a user, distribution list or resource."))
    .min(1)
    .max(MAX_SCHEDULES)
    .refine((list) => new Set(list.map((s) => s.toLowerCase())).size === list.length)
    .describe(
      `SMTP addresses to get free/busy for (1-${MAX_SCHEDULES}, unique; Graph supports up to ${MAX_SCHEDULES} entities). The period from startDateTime to endDateTime must be less than ${MAX_AVAILABILITY_WINDOW_DAYS} days (Graph limit).`,
    ),
  startDateTime: instant("Start of the period"),
  endDateTime: instant("End of the period"),
  availabilityViewInterval: z
    .number()
    .int()
    .min(5)
    .max(1440)
    .optional()
    .describe(
      "Minutes per slot in the returned availabilityView string (Graph: 5-1440, default 30).",
    ),
  timeZone: timeZoneField,
} satisfies z.ZodRawShape;

const FIND_MEETING_TIMES_SHAPE = {
  attendees: z
    .array(
      z.strictObject({
        address: email("SMTP address of the attendee or resource."),
        type: z
          .enum(["required", "resource"])
          .optional()
          .describe(
            '"required" (default) for a person or "resource" for a room/equipment. findMeetingTimes treats every person as required; use minimumAttendeePercentage to tolerate unavailable people.',
          ),
      }),
    )
    .max(MAX_ATTENDEES)
    .optional()
    .describe(
      `Attendees (max ${MAX_ATTENDEES}, a project safety limit). Omit or leave empty to look for free time of the organizer only.`,
    ),
  timeSlots: z
    .array(
      z
        .strictObject({
          startDateTime: instant("Start of the candidate period"),
          endDateTime: instant("End of the candidate period"),
        })
        .superRefine((slot, ctx) => {
          if (!windowOk(slot.startDateTime, slot.endDateTime)) windowIssue(ctx, []);
        }),
    )
    .min(1)
    .max(MAX_TIME_SLOTS)
    .superRefine((slots, ctx) => {
      // Zod also runs this on dirty input: leave an empty array, an unparseable slot, or an
      // inverted slot to its more specific min/item issue instead of adding a misleading horizon
      // error.
      const individuallyValid = slots.every((s) => {
        if (typeof s?.startDateTime !== "string" || typeof s?.endDateTime !== "string") {
          return false;
        }
        const start = instantTicks(s.startDateTime);
        const end = instantTicks(s.endDateTime);
        return start !== undefined && end !== undefined && start < end;
      });
      if (slots.length > 0 && individuallyValid && !slotsHorizonOk(slots)) {
        ctx.addIssue({
          code: "custom",
          message: `All candidate periods together must fit inside a search horizon less than ${MAX_AVAILABILITY_WINDOW_DAYS} days`,
        });
      }
    })
    .optional()
    .describe(
      `Candidate periods to search, each and the aggregate earliest-start-to-latest-end horizon less than ${MAX_AVAILABILITY_WINDOW_DAYS} days (project safety limit matching the getSchedule rule); up to ${MAX_TIME_SLOTS}. Strongly recommended: Graph picks its own search range otherwise.`,
    ),
  activityDomain: z
    .enum(["work", "personal", "unrestricted"])
    .optional()
    .describe(
      'Which hours qualify: "work" (default; the organizer\'s work hours), "personal" (work hours plus weekends) or "unrestricted" (any time).',
    ),
  meetingDuration: z
    .string()
    .refine((v) => {
      const m = /^PT(?:(\d{1,2})H)?(?:(\d{1,4})M)?$/.exec(v);
      if (m === null || (m[1] === undefined && m[2] === undefined)) return false;
      const minutes = Number(m[1] ?? 0) * 60 + Number(m[2] ?? 0);
      return minutes >= 1 && minutes <= MAX_DURATION_MINUTES;
    })
    .optional()
    .describe(
      `Meeting length as an ISO-8601 duration limited to hours/minutes, e.g. PT1H or PT1H30M or PT45M (Graph default PT30M; at most 24 hours, a project limit).`,
    ),
  maxCandidates: z
    .number()
    .int()
    .min(1)
    .max(MAX_CANDIDATES)
    .optional()
    .describe(`Maximum suggestions to return (1-${MAX_CANDIDATES}, a project safety limit).`),
  minimumAttendeePercentage: z
    .number()
    .min(0)
    .max(100)
    .optional()
    .describe("Minimum confidence (0-100) for a suggestion to be returned (Graph default 50)."),
  isOrganizerOptional: z
    .boolean()
    .optional()
    .describe("True if the organizer does not have to attend (default false)."),
  returnSuggestionReasons: z
    .boolean()
    .optional()
    .describe("True to include a suggestionReason for each suggestion (default false)."),
  timeZone: timeZoneField,
} satisfies z.ZodRawShape;

/** The two SM-8 availability tools. Not batch-reachable: no `batchPolicy`. */
export function registerCalendarAvailabilityTools(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const schedulePath = getSchedulePath(catalog);
  const meetingPath = findMeetingTimesPath(catalog);

  handlers.set("get-schedule", {
    // The window rule (start < end, max days) is a cross-field check made in the handler.
    inputSchema: GET_SCHEDULE_SHAPE,
    handler: async (args) => {
      try {
        const a = args as {
          schedules: string[];
          startDateTime: string;
          endDateTime: string;
          availabilityViewInterval?: number;
          timeZone?: string;
        };
        if (!windowOk(a.startDateTime, a.endDateTime)) {
          return errorResult(
            "invalid_input",
            `startDateTime must be before endDateTime and the window less than ${MAX_AVAILABILITY_WINDOW_DAYS} days.`,
          );
        }
        const body: Record<string, unknown> = {
          schedules: [...a.schedules],
          startTime: utcDateTime(a.startDateTime),
          endTime: utcDateTime(a.endDateTime),
        };
        if (a.availabilityViewInterval !== undefined) {
          body["availabilityViewInterval"] = a.availabilityViewInterval;
        }
        const res = await client.request({
          pathPattern: schedulePath,
          method: "post",
          body,
          headers: preferHeaders(a.timeZone),
        });
        const maxView = maxViewLength(
          a.startDateTime,
          a.endDateTime,
          a.availabilityViewInterval ?? DEFAULT_VIEW_INTERVAL_MINUTES,
        );
        // Graph returns exactly one scheduleInformation per requested schedule (an unresolved
        // one carries an `error`), so a missing or extra entry fails closed.
        const items = objects(
          (res.data as { value?: unknown } | null | undefined)?.value ?? null,
          Math.min(a.schedules.length, MAX_OUT_SCHEDULES),
          "schedule list",
          (o) => projectScheduleInformation(o, maxView),
        );
        if (items === undefined || items.length !== a.schedules.length) {
          malformed("schedule list");
        }
        const meta: Record<string, unknown> = {
          count: items.length,
          clientRequestId: res.clientRequestId,
        };
        if (res.requestId !== undefined) meta["requestId"] = res.requestId;
        return successResult(items, meta);
      } catch (err) {
        return failure("get-schedule", err);
      }
    },
  });

  handlers.set("find-meeting-times", {
    inputSchema: FIND_MEETING_TIMES_SHAPE,
    handler: async (args) => {
      try {
        const a = args as {
          attendees?: Array<{ address: string; type?: "required" | "resource" }>;
          timeSlots?: Array<{ startDateTime: string; endDateTime: string }>;
          activityDomain?: string;
          meetingDuration?: string;
          maxCandidates?: number;
          minimumAttendeePercentage?: number;
          isOrganizerOptional?: boolean;
          returnSuggestionReasons?: boolean;
          timeZone?: string;
        };
        for (const slot of a.timeSlots ?? []) {
          if (!windowOk(slot.startDateTime, slot.endDateTime)) {
            return errorResult(
              "invalid_input",
              `Each time slot must start before it ends and span less than ${MAX_AVAILABILITY_WINDOW_DAYS} days.`,
            );
          }
        }
        if (a.timeSlots !== undefined && !slotsHorizonOk(a.timeSlots)) {
          return errorResult(
            "invalid_input",
            `All time slots together must fit inside a search horizon less than ${MAX_AVAILABILITY_WINDOW_DAYS} days.`,
          );
        }
        const body: Record<string, unknown> = {};
        if (a.attendees !== undefined) {
          body["attendees"] = a.attendees.map((x) => ({
            type: x.type ?? "required",
            emailAddress: { address: x.address },
          }));
        }
        if (a.timeSlots !== undefined || a.activityDomain !== undefined) {
          const tc: Record<string, unknown> = {};
          if (a.activityDomain !== undefined) tc["activityDomain"] = a.activityDomain;
          if (a.timeSlots !== undefined) {
            tc["timeSlots"] = a.timeSlots.map((s) => ({
              start: utcDateTime(s.startDateTime),
              end: utcDateTime(s.endDateTime),
            }));
          }
          body["timeConstraint"] = tc;
        }
        if (a.meetingDuration !== undefined) body["meetingDuration"] = a.meetingDuration;
        if (a.maxCandidates !== undefined) body["maxCandidates"] = a.maxCandidates;
        if (a.minimumAttendeePercentage !== undefined) {
          body["minimumAttendeePercentage"] = a.minimumAttendeePercentage;
        }
        if (a.isOrganizerOptional !== undefined)
          body["isOrganizerOptional"] = a.isOrganizerOptional;
        if (a.returnSuggestionReasons !== undefined) {
          body["returnSuggestionReasons"] = a.returnSuggestionReasons;
        }
        const res = await client.request({
          pathPattern: meetingPath,
          method: "post",
          body,
          headers: preferHeaders(a.timeZone),
        });
        const result = projectMeetingTimes(res.data, a.maxCandidates ?? MAX_OUT_SUGGESTIONS);
        const suggestions = result["meetingTimeSuggestions"] as unknown[];
        const meta: Record<string, unknown> = {
          count: suggestions.length,
          clientRequestId: res.clientRequestId,
        };
        if (res.requestId !== undefined) meta["requestId"] = res.requestId;
        return successResult(result, meta);
      } catch (err) {
        return failure("find-meeting-times", err);
      }
    },
  });
}
