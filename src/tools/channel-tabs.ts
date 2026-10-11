import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { isRecord, pickScalars } from "./chats.js";
import { withProjection } from "./chat-messages.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";
import { withSanitizer } from "./teams.js";

/**
 * Teams channel tabs and files folder (SM-12 part C): `list-channel-tabs` and
 * `get-channel-files-folder` (both v1.0 GET).
 *
 * Delegated permissions (current Microsoft docs, October 2026; see PERMISSIONS.md): the tabs page
 * documents `TeamsTab.Read.All` as least privileged (work/school only), matching the catalog. The
 * files folder page documents `Files.Read.All` as least privileged (personal accounts: `Files.Read`);
 * the catalog uses it since D28. The server never
 * pre-checks scopes; Graph decides, and a 403 is a `capability_gap`.
 *
 * Tabs: the page documents `$filter`, `$select` and `$expand` but no `$top` and no ordering. Only a
 * local `top` limit (pages followed, no `$top` sent) and the typed `expand: ["teamsApp"]` are
 * offered; no `$filter`/`$select`. A tab is projected to `id`, `displayName`, `webUrl` (Teams deep
 * link), `sortOrderIndex` and, when expanded, `teamsApp` reduced to `id`, `displayName`,
 * `distributionMethod`. The whole `configuration` object (`contentUrl`, `websiteUrl`, `removeUrl`,
 * `entityId`, and open-type custom properties) is NOT exposed: those are app-defined URLs and
 * identifiers that can embed tokens or tenant data, and the tab's own `webUrl` already links to it.
 *
 * Files folder: the page documents no OData options at all. The driveItem is projected to an
 * allowlist of folder metadata: `id`, `name`, `webUrl`, `createdDateTime`, `lastModifiedDateTime`,
 * `size`, `folder.childCount` and `parentReference` reduced to `driveId`, `driveType`, `id`.
 * Identity sets (`createdBy`, `lastModifiedBy`), `fileSystemInfo`, `eTag`/`cTag`, `shared`,
 * `@microsoft.graph.downloadUrl`, `content`, permissions and every unknown property are dropped.
 * No catalog tool lists or downloads the folder's files: browsing channel files is outside this
 * tool surface (D29).
 */

const TAB_KEYS = ["displayName", "webUrl", "sortOrderIndex"] as const;
const TEAMS_APP_KEYS = ["id", "displayName", "distributionMethod"] as const;

/**
 * Allowlist projection for one tab (single, collection item, batch body). `teamsApp` is kept only
 * when that request asked for it (`withApp`, from the wire `$expand`); an over-returned one is dropped.
 */
export function sanitizeTab(item: unknown, withApp = false): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, ["id", ...TAB_KEYS]);
  if (withApp && Object.hasOwn(item, "teamsApp")) {
    const app = item["teamsApp"];
    if (app === null) out["teamsApp"] = null;
    else if (isRecord(app)) out["teamsApp"] = pickScalars(app, TEAMS_APP_KEYS);
  }
  return out;
}

const FOLDER_KEYS = [
  "id",
  "name",
  "webUrl",
  "createdDateTime",
  "lastModifiedDateTime",
  "size",
] as const;
const PARENT_KEYS = ["driveId", "driveType", "id"] as const;

/** Last-line sanitizer for the files-folder driveItem. */
export function sanitizeFilesFolder(item: unknown): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, FOLDER_KEYS);
  const folder = item["folder"];
  if (isRecord(folder)) out["folder"] = pickScalars(folder, ["childCount"]);
  const parent = item["parentReference"];
  if (isRecord(parent)) out["parentReference"] = pickScalars(parent, PARENT_KEYS);
  return out;
}

const wantsApp = (query: Readonly<Record<string, unknown>>): boolean => {
  const expand = query["$expand"];
  return typeof expand === "string" && expand.split(",").includes("teamsApp");
};

// Query-aware projection (as for pinned chat messages): `teamsApp` survives only when the request's
// own `$expand` named it, on the single path and in batch alike.
const tabsContract = (): QueryContract =>
  withProjection(
    defineQueryContract({
      // The page documents no `$top`: `top` is a local limit and pages are followed.
      top: { default: 50, max: 200, wire: false },
      expand: { allowed: ["teamsApp"] },
      // Load-bearing on a collection contract: it widens the inner batch keep-set
      // (`projectCollection`) so tab scalars survive until `sanitizeTab` runs. Batch parity tests pin it.
      projectEntity: { keep: TAB_KEYS },
    }),
    true,
    (item, query) => sanitizeTab(item, wantsApp(query)),
  );

const filesFolderContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      projectEntity: {
        keep: [...FOLDER_KEYS.filter((k) => k !== "id"), "folder", "parentReference"],
      },
    }),
    false,
    sanitizeFilesFolder,
  );

export function registerChannelTabTools(
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
    "list-channel-tabs",
    listTool(client, "list-channel-tabs", patternOf("list-channel-tabs"), tabsContract()),
  );
  handlers.set(
    "get-channel-files-folder",
    getTool(
      client,
      "get-channel-files-folder",
      patternOf("get-channel-files-folder"),
      filesFolderContract(),
    ),
  );
}
