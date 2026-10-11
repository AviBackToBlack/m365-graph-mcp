import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { isRecord, pickScalars } from "./chats.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";
import { withSanitizer } from "./teams.js";

/**
 * SharePoint site drives (SM-13 part B): `list-sharepoint-site-drives` and
 * `get-sharepoint-site-drive-by-id` (both v1.0 GET). The site `items` routes were removed from the
 * catalog by D32: Microsoft documents the `site` relationship `items` as "Used to address any item
 * contained in this site. This collection can't be enumerated." and publishes no API page for them.
 *
 * Delegated permissions (current Microsoft docs, October 2026): drive-list and drive-get both
 * document `Files.Read` as least privileged for work/school AND personal accounts, with
 * `Files.Read.All`, `Sites.Read.All` (and the ReadWrite variants) higher. Neither page mentions
 * `Sites.Selected`. D33 sets the catalog to `Files.Read.All` or `Sites.Read.All` or `Sites.Selected`
 * (D30); `Files.Read` (the signed-in user's own files) is live-checked in SM-13 part D before it is
 * considered. The server never pre-checks scopes; Graph decides, and a 403 is a `capability_gap`.
 *
 * List: drive-list documents `$expand`, `$select`, `$skipToken`, `$top` and `$orderby`. Only the
 * documented `$top` (typed `top`, sent on the wire, pages followed) is offered; no `$select`,
 * `$expand` or `$orderby`, so `system`-faceted (hidden) drives are never requested. Get: only
 * `$select` is documented and it is not offered, so no query option exists. `site-id` uses the
 * shared site shape, `drive-id` the opaque id shape (`b!...`) in `graph/placeholders.ts`.
 *
 * Every response (single, list item, batch body) is projected to an allowlist of scalar drive
 * metadata: `id`, `name`, `description`, `driveType`, `webUrl`, `createdDateTime`,
 * `lastModifiedDateTime`. `owner`, `createdBy`, `lastModifiedBy` (identities), `quota`, `root`,
 * `items`, `list`, `sharepointIds`, `system`, `eTag` and every unknown property are dropped.
 * No catalog tool lists or downloads drive content (D29).
 */

const DRIVE_KEYS = [
  "name",
  "description",
  "driveType",
  "webUrl",
  "createdDateTime",
  "lastModifiedDateTime",
] as const;

/** Allowlist projection for one drive (single, collection item, batch body). */
export function sanitizeDrive(item: unknown): unknown {
  if (!isRecord(item)) return undefined;
  return pickScalars(item, ["id", ...DRIVE_KEYS]);
}

const listContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      // The page documents `$top` (no stated maximum): sent on the wire, pages followed to the limit.
      top: { default: 25, max: 200, pageMax: 100 },
      // Load-bearing on a collection contract: it widens the inner batch keep-set
      // (`projectCollection`) so drive scalars survive until `sanitizeDrive` runs.
      projectEntity: { keep: DRIVE_KEYS },
    }),
    true,
    sanitizeDrive,
  );

const driveContract = (): QueryContract =>
  withSanitizer(defineQueryContract({ projectEntity: { keep: DRIVE_KEYS } }), false, sanitizeDrive);

export function registerSharePointDriveTools(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const patternOf = (toolName: string): string => {
    const entry = catalog.find((e) => e.toolName === toolName);
    if (entry === undefined || entry.method !== "get" || entry["apiVersion"] === "beta") {
      throw new Error(`Catalog has no v1.0 GET entry for ${toolName}.`);
    }
    return entry.pathPattern;
  };
  handlers.set(
    "list-sharepoint-site-drives",
    listTool(
      client,
      "list-sharepoint-site-drives",
      patternOf("list-sharepoint-site-drives"),
      listContract(),
    ),
  );
  handlers.set(
    "get-sharepoint-site-drive-by-id",
    getTool(
      client,
      "get-sharepoint-site-drive-by-id",
      patternOf("get-sharepoint-site-drive-by-id"),
      driveContract(),
    ),
  );
}
