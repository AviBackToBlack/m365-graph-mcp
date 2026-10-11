import { z } from "zod";
import type { EndpointCatalog } from "../catalog.js";
import { errorResult, successResult } from "../envelope.js";
import { validateBatch, type BatchTargets } from "../graph/batch.js";
import type { GraphClient } from "../graph/client.js";
import { GraphClientError } from "../graph/errors.js";
import { endpointVersion } from "../graph/url.js";
import type { ToolImplementation } from "../registry.js";

/**
 * Advertised input shape. Advisory only: `validateBatch` re-validates everything, so the
 * handler stays safe even if this schema were loosened.
 */
const inputSchema = {
  requests: z
    .array(
      z.strictObject({
        id: z.string().describe("Unique correlation id (letters, digits, _ or -; max 64)."),
        method: z.enum(["GET"]).describe("Only GET is accepted."),
        url: z
          .string()
          .describe("Relative Graph path starting with '/', e.g. /me/messages?$top=5."),
        headers: z
          .record(z.string(), z.string())
          .optional()
          .describe("Optional ConsistencyLevel or Prefer header only."),
      }),
    )
    .min(1)
    .max(20),
} satisfies z.ZodRawShape;

/**
 * `graph-batch`. `targets` are the endpoints a subrequest may target, each with its query
 * policy; in production it is `batchTargets` over the implemented tools (`createHandlers`).
 */
export function createBatchTool(
  client: GraphClient,
  catalog: EndpointCatalog,
  targets: BatchTargets,
): ToolImplementation {
  const batchEntry = catalog.find((e) => e.method === "post" && e.pathPattern === "/$batch");
  if (batchEntry === undefined) throw new Error("Catalog has no /$batch entry.");
  const version = endpointVersion(batchEntry);
  return {
    inputSchema,
    handler: async (args) => {
      try {
        const validated = validateBatch(args, targets, version);
        const result = await client.batch(validated);
        const meta: Record<string, unknown> = { clientRequestId: result.clientRequestId };
        if (result.requestId !== undefined) meta["requestId"] = result.requestId;
        return successResult({ responses: result.responses }, meta);
      } catch (err) {
        if (err instanceof GraphClientError) return err.toResult();
        return errorResult("graph_error", "Tool graph-batch failed unexpectedly.");
      }
    },
  };
}
