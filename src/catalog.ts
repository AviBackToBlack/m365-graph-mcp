import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { z } from "zod";

export const endpointDefinitionSchema = z
  .object({
    pathPattern: z.string().min(1).startsWith("/"),
    method: z.enum(["get", "post"]),
    toolName: z.string().min(1),
    readOnly: z.boolean().optional(),
  })
  .passthrough();

export const endpointCatalogSchema = z
  .array(endpointDefinitionSchema)
  .superRefine((entries, ctx) => {
    const seen = new Map<string, number>();

    for (const [index, entry] of entries.entries()) {
      const previous = seen.get(entry.toolName);
      if (previous !== undefined) {
        ctx.addIssue({
          code: "custom",
          path: [index, "toolName"],
          message: `duplicate toolName ${entry.toolName}; first seen at index ${previous}`,
        });
      } else {
        seen.set(entry.toolName, index);
      }
    }
  });

export type EndpointDefinition = z.infer<typeof endpointDefinitionSchema>;
export type EndpointCatalog = z.infer<typeof endpointCatalogSchema>;

// Names that would read as a generic proxy parameter are never exposed verbatim.
// Today only get-sharepoint-site-by-path has such a placeholder ({path}, a site-relative path).
const RESERVED_PARAM_RENAMES: Readonly<Record<string, string>> = {
  path: "relative-path",
  url: "resource-url-segment",
  method: "method-segment",
};

/** Exposed input-parameter name for a catalog path placeholder. */
export function exposedParamName(placeholder: string): string {
  return RESERVED_PARAM_RENAMES[placeholder] ?? placeholder;
}

// Resolved relative to this module (src/ or dist/ -> package root), never process.cwd():
// Claude Desktop launches servers with an arbitrary working directory.
export const DEFAULT_CATALOG_PATH = resolve(
  dirname(fileURLToPath(import.meta.url)),
  "..",
  "endpoints.json",
);

export async function loadEndpointCatalog(
  catalogPath = DEFAULT_CATALOG_PATH,
): Promise<EndpointCatalog> {
  const raw = await readFile(catalogPath, "utf8");
  const parsed: unknown = JSON.parse(raw);
  return endpointCatalogSchema.parse(parsed);
}
