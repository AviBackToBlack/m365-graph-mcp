import type { z } from "zod";
import type { EndpointCatalog } from "../catalog.js";
import { GraphClientError } from "../graph/errors.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { EVENT_COMPACT, MAX_WINDOW_DAYS } from "./calendar.js";
import { createDeltaTool } from "./delta.js";
import {
  assertWindow,
  timeZonePreferFor,
  timeZoneSchema,
  windowStampSchema,
} from "./query-contract.js";
import { canonicalInstant } from "./timestamps.js";

/**
 * `list-calendar-view-delta` (SM-8): incremental event sync for a time window of the
 * signed-in user's default calendar (`GET /me/calendarView/delta`, v1.0).
 *
 * The opaque-token design lives in `delta.ts`. Calendar specifics, per the current Graph docs:
 * - `startDateTime`/`endDateTime` are REQUIRED on the initial call and are encoded into the
 *   state token afterwards, so they are initial-only (a token call that carries them is
 *   rejected). They follow the SM-7 window rules (offsets accepted, canonical UTC sent, start <
 *   end, at most `MAX_WINDOW_DAYS`).
 * - `$select`, `$expand`, `$filter`, `$orderby` and `$search` are NOT supported by this delta
 *   function, and Graph returns the full event (including `body`); none of them are offered,
 *   and every page is projected to a body-free allowlist.
 * - `$top` is not supported; the page size is `Prefer: odata.maxpagesize`.
 * - `timeZone` is `Prefer: outlook.timezone`. Headers are not part of the token, so it is
 *   accepted on every call and sent on every request.
 * - Graph may report events OUTSIDE the window (and deletions) under `@removed`; those are
 *   returned as markers, not filtered.
 *
 * The matching `/me/events/delta()` entry (`list-calendar-events-delta`) is documented by
 * Microsoft in the beta API only; it lives in `calendar-events-delta.ts` (D2).
 */

const DELTA_TOOL = "list-calendar-view-delta";
// Full events are heavy; this matches the SM-7 event page size.
const PAGE_SIZE_MAX = 100;

const initialShape = {
  startDateTime: windowStampSchema("Start of the time range", MAX_WINDOW_DAYS, false)
    .optional()
    .describe(
      `Initial call only (required there; encoded in the token afterwards). Start of the time range as an ISO-8601 timestamp with a Z or an explicit UTC offset, e.g. 2024-01-31T00:00:00Z. Sent to Graph as the equivalent UTC value. Must be before endDateTime; the window may span at most ${MAX_WINDOW_DAYS} days (a project safety limit).`,
    ),
  endDateTime: windowStampSchema("End of the time range", MAX_WINDOW_DAYS, false)
    .optional()
    .describe(
      "Initial call only (required there; encoded in the token afterwards). End of the time range, same format as startDateTime.",
    ),
} satisfies z.ZodRawShape;

const alwaysShape = { timeZone: timeZoneSchema() } satisfies z.ZodRawShape;

export function registerCalendarDeltaTool(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const entry = catalog.find((e) => e.toolName === DELTA_TOOL);
  if (entry === undefined || entry.method !== "get") {
    throw new Error(`Catalog has no GET entry for ${DELTA_TOOL}.`);
  }
  handlers.set(
    DELTA_TOOL,
    createDeltaTool(client, {
      toolName: DELTA_TOOL,
      entry,
      pathParamNames: [],
      initialShape,
      alwaysShape,
      initialOnlyMessage:
        "startDateTime and endDateTime apply to the initial call only; they are encoded in the token.",
      initialOnlyList: "startDateTime/endDateTime",
      buildInitialQuery: (typed) => {
        const start = typed["startDateTime"];
        const end = typed["endDateTime"];
        if (typeof start !== "string" || typeof end !== "string") {
          throw new GraphClientError(
            "invalid_input",
            "startDateTime and endDateTime are required on the initial call.",
          );
        }
        assertWindow(start, end, MAX_WINDOW_DAYS);
        return {
          startDateTime: canonicalInstant(start) as string,
          endDateTime: canonicalInstant(end) as string,
        };
      },
      preferences: (typed) =>
        typeof typed["timeZone"] === "string" ? [timeZonePreferFor(typed["timeZone"])] : [],
      // No $select exists for this function, so Graph returns whole events: `body`,
      // attachments and extension data are never in the allowlist.
      keepKeys: new Set(["id", ...EVENT_COMPACT, "changeKey", "@odata.type", "@odata.etag"]),
      pageSizeMax: PAGE_SIZE_MAX,
    }),
  );
}
