import { z } from "zod";
import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { createDeltaTool } from "./delta.js";
import { timeZonePreferFor, timeZoneSchema } from "./query-contract.js";
import { canonicalInstant, instantTicks } from "./timestamps.js";

/**
 * `list-calendar-events-delta` (SM-8 part C): incremental sync of the signed-in
 * user's events (`GET /me/events/delta`), documented by Microsoft in the BETA API only. The
 * catalog entry carries `apiVersion: "beta"`, which is the only thing that selects `/beta/`;
 * `delta.ts` validates every returned link against that version, so a `/v1.0/` link is refused.
 *
 * The opaque-token design lives in `delta.ts`. Contract per the current beta `event: delta` page:
 * - `startDateTime` is OPTIONAL and initial-only (encoded in the token afterwards): events
 *   starting on or after it. `endDateTime` is NOT supported and is not offered.
 * - `$select`, `$expand`, `$filter`, `$orderby` and `$search` are not supported; none is offered.
 * - Graph returns only `id`, `type`, `start` and `end` per event (fetch details with the
 *   per-event tools), so the projection allowlist is exactly those plus `@removed`.
 * - `Prefer: odata.maxpagesize` is the page size; `Prefer: outlook.timezone` is documented for
 *   this function and applies to the returned `start`/`end` (never to `startDateTime`).
 * - Unlike calendarView, recurring series are returned as their series master.
 */

const DELTA_TOOL = "list-calendar-events-delta";
// Graph documents no maximum for this function; stay at the calendar page size.
const PAGE_SIZE_MAX = 100;

const initialShape = {
  startDateTime: z
    .string()
    .refine((v) => instantTicks(v) !== undefined)
    .optional()
    .describe(
      "Initial call only (optional; encoded in the token afterwards). Only events starting on or after this instant are tracked. ISO-8601 timestamp with a Z or an explicit UTC offset, e.g. 2024-01-31T00:00:00Z (local times without Z/offset are rejected). Sent to Graph as the equivalent UTC value. Omit to track all events. There is no end bound.",
    ),
} satisfies z.ZodRawShape;

const alwaysShape = { timeZone: timeZoneSchema() } satisfies z.ZodRawShape;

export function registerCalendarEventsDeltaTool(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const entry = catalog.find((e) => e.toolName === DELTA_TOOL);
  if (entry === undefined || entry.method !== "get") {
    throw new Error(`Catalog has no GET entry for ${DELTA_TOOL}.`);
  }
  if (entry["apiVersion"] !== "beta") {
    throw new Error(`${DELTA_TOOL} is a beta-only endpoint.`);
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
        "startDateTime applies to the initial call only; it is encoded in the token.",
      initialOnlyList: "startDateTime",
      buildInitialQuery: (typed) => {
        const start = typed["startDateTime"];
        return typeof start === "string"
          ? { startDateTime: canonicalInstant(start) as string }
          : {};
      },
      preferences: (typed) =>
        typeof typed["timeZone"] === "string" ? [timeZonePreferFor(typed["timeZone"])] : [],
      // Documented limited projection: exactly id, type, start, end (+ `@removed`, kept by delta.ts).
      keepKeys: new Set(["id", "type", "start", "end"]),
      pageSizeMax: PAGE_SIZE_MAX,
    }),
  );
}
