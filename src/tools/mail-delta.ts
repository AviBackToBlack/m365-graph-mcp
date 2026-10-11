import { z } from "zod";
import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { createDeltaTool } from "./delta.js";
import { LIST_DEFAULT_SELECT, MESSAGE_METADATA } from "./mail.js";
import { isRealUtcTimestamp } from "./timestamps.js";

/**
 * `list-mail-folder-messages-delta` (SM-6): incremental message sync for one mail folder.
 *
 * The opaque-token design, bounds, `@removed` handling and restart mapping live in the
 * generic `delta.ts`. This file only declares what is specific to mail: the initial-only
 * options (`select`, `changeType`, `receivedAfter`, Graph's only supported delta filter) and
 * the message allowlist.
 */

const DELTA_TOOL = "list-mail-folder-messages-delta";
const PAGE_SIZE_MAX = 250;

const initialShape = {
  select: z
    .array(z.enum(MESSAGE_METADATA))
    .min(1)
    .max(MESSAGE_METADATA.length)
    .optional()
    .describe(
      `Initial call only (encoded in the token afterwards). Properties to track/return. Default: ${LIST_DEFAULT_SELECT.join(", ")}.`,
    ),
  changeType: z
    .enum(["created", "updated", "deleted"])
    .optional()
    .describe("Initial call only. Return only this kind of change."),
  receivedAfter: z
    .string()
    .refine(isRealUtcTimestamp)
    .optional()
    .describe(
      "Initial call only. UTC timestamp like 2024-01-31T00:00:00Z; only messages received at or after it (Graph's only supported delta filter, receivedDateTime ge).",
    ),
} satisfies z.ZodRawShape;

export function registerMailDeltaTool(
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
      pathParamNames: ["mailFolder-id"],
      initialShape,
      initialOnlyMessage:
        "select, changeType and receivedAfter apply to the initial call only; they are encoded in the token.",
      initialOnlyList: "select/changeType/receivedAfter",
      buildInitialQuery: (typed) => {
        const query: Record<string, string> = {
          $select: [
            ...new Set((typed["select"] as string[] | undefined) ?? LIST_DEFAULT_SELECT),
          ].join(","),
        };
        if (typed["changeType"] !== undefined) query["changeType"] = typed["changeType"] as string;
        if (typed["receivedAfter"] !== undefined) {
          query["$filter"] = `receivedDateTime ge ${typed["receivedAfter"] as string}`;
        }
        return query;
      },
      keepKeys: new Set(["id", ...MESSAGE_METADATA, "@odata.type", "@odata.etag"]),
      pageSizeMax: PAGE_SIZE_MAX,
    }),
  );
}
