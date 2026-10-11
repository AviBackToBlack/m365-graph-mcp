import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";

/**
 * Calendar core tools (SM-7). Like the mail tools, every query option, default and limit
 * lives in a per-tool contract that is ALSO the tool's `batchPolicy`.
 *
 * - Lists/views/instances return compact events (no `body`; `bodyPreview` is selectable), the
 *   projector drops everything outside the `$select` allowlist on every page and in batch.
 * - `get-*-event` is the only place a body is returned (`bodyType`, default text).
 * - calendarView and instances REQUIRE `startDateTime`/`endDateTime` (start < end, at most
 *   `MAX_WINDOW_DAYS`). The tool accepts `Z` or `+/-HH:MM` offsets and sends canonical UTC `Z`;
 *   batch accepts only that canonical UTC form (see `query-contract.ts`).
 * - `timeZone` sets the zone of returned start/end via `Prefer: outlook.timezone`; it is never
 *   applied to the window.
 * - The seven event tools (`supportsExpandExtendedProperties`) take an optional typed
 *   `extendedPropertyId`, expanded as the one documented filtered `$expand` and projected as
 *   id/value pairs. list-calendars does not offer it. `$search` is not offered.
 *
 * Delegated permissions (current Microsoft docs; `endpoints.json` scope metadata is
 * non-normative): bodyless event lists, calendarView, instances and list-calendars need only
 * Calendars.ReadBasic; the two get-*-event tools return the body by default (Calendars.Read);
 * any `extendedPropertyId` request needs Calendars.Read. See PERMISSIONS.md.
 */

/**
 * Project SAFETY limit (not a Graph limit) on a calendarView / instances window. Graph
 * documents none, but an unbounded range expands every recurrence in it; a year (leap-year
 * safe) covers "this year" questions while keeping the expansion and result size predictable.
 */
export const MAX_WINDOW_DAYS = 366;

/**
 * Properties of the Graph v1.0 `event` resource that list-style tools may return.
 * `body` is deliberately absent (use get-*-event). `createdDateTime`/`lastModifiedDateTime`
 * are not here because Graph documents that calendarView rejects them in `$select`; they are
 * added only for the plain event list/get contracts (`EVENT_TIMESTAMPS`).
 */
export const EVENT_PROPERTIES = [
  "id",
  "subject",
  "bodyPreview",
  "start",
  "end",
  "originalStart",
  "originalStartTimeZone",
  "originalEndTimeZone",
  "isAllDay",
  "isCancelled",
  "isDraft",
  "isOnlineMeeting",
  "onlineMeetingProvider",
  "onlineMeeting",
  "location",
  "locations",
  "organizer",
  "attendees",
  "responseStatus",
  "showAs",
  "sensitivity",
  "importance",
  "type",
  "seriesMasterId",
  "recurrence",
  "categories",
  "hasAttachments",
  "iCalUId",
  "isOrganizer",
  "isReminderOn",
  "reminderMinutesBeforeStart",
  "responseRequested",
  "allowNewTimeProposals",
  "hideAttendees",
  "webLink",
  "changeKey",
  "transactionId",
] as const;

export const EVENT_COMPACT = [
  "id",
  "subject",
  "start",
  "end",
  "location",
  "organizer",
  "isAllDay",
  "isCancelled",
  "showAs",
  "responseStatus",
  "type",
  "seriesMasterId",
  "isOnlineMeeting",
  "onlineMeetingProvider",
  "onlineMeeting",
] as const;

/** Series masters only exist in event listings, where the recurrence pattern is the point. */
const EVENT_LIST_DEFAULT = [...EVENT_COMPACT, "recurrence"] as const;

const EVENT_GET_DEFAULT = [
  ...EVENT_COMPACT,
  "recurrence",
  "attendees",
  "bodyPreview",
  "body",
  "locations",
  "sensitivity",
  "importance",
  "categories",
  "hasAttachments",
  "isOrganizer",
  "isReminderOn",
  "reminderMinutesBeforeStart",
  "originalStartTimeZone",
  "originalEndTimeZone",
  "iCalUId",
  "webLink",
] as const;

const EVENT_ORDERBY = ["start/dateTime", "end/dateTime", "subject"] as const;

const CALENDAR_PROPERTIES = [
  "id",
  "name",
  "color",
  "hexColor",
  "isDefaultCalendar",
  "changeKey",
  "canShare",
  "canViewPrivateItems",
  "canEdit",
  "owner",
  "allowedOnlineMeetingProviders",
  "defaultOnlineMeetingProvider",
  "isTallyingResponses",
  "isRemovable",
] as const;

// $top: calendarView documents 1..1000, events/instances document no maximum, so one
// conservative PROJECT page size (100, a safety limit, not a Graph limit) applies to all of
// them; `top` up to 500 (also a project limit) is reached by following pages in the
// single-request path, batch is bound to the page size.
const EVENT_TOP = { default: 25, max: 500, pageMax: 100 } as const;
// Calendar documents no `$filter` grammar rule tying it to `$orderby` (unlike messages).
const EVENT_ORDER = { allowed: EVENT_ORDERBY, maxClauses: 2, requireInFilter: false } as const;

/** Selectable on /events list/get, but not on calendarView (documented); instances: conservative. */
const EVENT_TIMESTAMPS = ["createdDateTime", "lastModifiedDateTime"] as const;

export const eventListContract = (extendedProperty = true): QueryContract =>
  defineQueryContract({
    top: EVENT_TOP,
    select: {
      allowed: [...EVENT_PROPERTIES, ...EVENT_TIMESTAMPS],
      default: EVENT_LIST_DEFAULT,
      requiredInBatch: true,
    },
    filter: { maxLength: 1000 },
    orderby: EVENT_ORDER,
    skip: { max: 10_000 },
    timeZone: true,
    extendedProperty,
  });

/** calendarView and instances: a bounded, required UTC window instead of `skip`. */
export const eventWindowContract = (extendedProperty = true): QueryContract =>
  defineQueryContract({
    top: EVENT_TOP,
    select: { allowed: EVENT_PROPERTIES, default: EVENT_COMPACT, requiredInBatch: true },
    filter: { maxLength: 1000 },
    orderby: EVENT_ORDER,
    timeZone: true,
    window: { maxDays: MAX_WINDOW_DAYS },
    extendedProperty,
  });

const eventGetContract = (): QueryContract =>
  defineQueryContract({
    select: {
      allowed: [...EVENT_PROPERTIES, ...EVENT_TIMESTAMPS, "body"],
      default: EVENT_GET_DEFAULT,
      requiredInBatch: false,
    },
    bodyType: { default: "text" },
    timeZone: true,
    extendedProperty: true,
  });

// Calendars document no `$top` maximum or page size, so 100 per request is a project safety
// limit. `top` up to 1000 follows validated nextLinks (the client's ceiling is 10 pages), so
// calendars beyond the first page are reachable up to that 1000-item ceiling, after which
// the result reports hasMore; batch cannot paginate and stays at 100.
// No `$skip` is offered or synthesized; continuation comes only from Graph's own nextLink.
const calendarListContract = (): QueryContract =>
  defineQueryContract({
    top: { default: 25, max: 1000, pageMax: 100 },
    select: {
      allowed: CALENDAR_PROPERTIES,
      default: CALENDAR_PROPERTIES,
      requiredInBatch: false,
    },
  });

/** Adds the eight SM-7 tools and the two SM-8 shared-calendar tools. Each is batch-reachable only through its own contract. */
export function registerCalendarTools(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const patternOf = (toolName: string): string => {
    const entry = catalog.find((e) => e.toolName === toolName);
    if (entry === undefined || entry.method !== "get") {
      throw new Error(`Catalog has no GET entry for ${toolName}.`);
    }
    return entry.pathPattern;
  };
  const list = (toolName: string, contract: QueryContract) =>
    handlers.set(toolName, listTool(client, toolName, patternOf(toolName), contract));
  const get = (toolName: string, contract: QueryContract) =>
    handlers.set(toolName, getTool(client, toolName, patternOf(toolName), contract));

  list("list-calendar-events", eventListContract());
  get("get-calendar-event", eventGetContract());
  list("list-specific-calendar-events", eventListContract());
  get("get-specific-calendar-event", eventGetContract());
  list("get-calendar-view", eventWindowContract());
  list("get-specific-calendar-view", eventWindowContract());
  list("list-calendar-event-instances", eventWindowContract());
  list("list-calendars", calendarListContract());

  // SM-8 shared-calendar tools: the SM-7 list/window contracts unchanged, bound to another
  // user's default calendar (`{user-id}` is validated as a GUID or UPN by the client). The
  // catalog does not flag `supportsExpandExtendedProperties` on these entries, so no
  // `extendedPropertyId` is offered. Delegated access to the calendar is required.
  list("list-shared-calendar-events", eventListContract(false));
  list("get-shared-calendar-view", eventWindowContract(false));
}
