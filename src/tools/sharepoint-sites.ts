import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { isRecord, pickScalars } from "./chats.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";
import { withSanitizer } from "./teams.js";

/**
 * SharePoint sites (SM-13 part A): `search-sharepoint-sites`, `get-sharepoint-site` and
 * `get-sharepoint-site-by-path` (all v1.0 GET).
 *
 * Delegated permissions (current Microsoft docs, October 2026; see PERMISSIONS.md): site-search,
 * site-get and site-getbypath all document `Sites.Read.All` as least privileged (work/school
 * only; personal Microsoft accounts "Not supported"). The search page states "This method doesn't
 * support the Sites.Selected application permission", matching the catalog's `["Sites.Read.All"]`;
 * the get/by-path pages list no `Sites.Selected`; D30 keeps it as the resource-scoped Selected
 * alternative there (it needs a grant on the target site; see PERMISSIONS.md). The server never pre-checks scopes; Graph decides, and a 403 is a `capability_gap`.
 *
 * Search: `GET /sites?search={query}` (documented bare name `search`; free-text keywords). The page
 * documents no other query option (the prose notes `createdDateTime` as the only sortable property,
 * but no `$orderby`/`$top`/`$select`/`$filter` is specified), so only a REQUIRED typed `search`
 * (a bare `/sites` is never sent, and a text without a letter or digit such as `*` is rejected) and
 * a local `top` limit (pages followed, no `$top` sent) are offered.
 *
 * Get / by-path: no query options at all are documented, none are offered. By-path addresses
 * `/sites/{hostname}:/{server-relative-path}`; the hostname and the multi-segment path are validated
 * in `graph/placeholders.ts` + `graph/url.ts` (colon-free, no dot segments, no empty segments, no
 * `? # \` or control characters, no OData control segments; the final URL ends at the path, so
 * nothing but the site itself is addressable).
 *
 * Every response (single, list item, batch body) is projected to an allowlist of scalar site
 * metadata: `id`, `name`, `displayName`, `description`, `webUrl`, `createdDateTime`,
 * `lastModifiedDateTime`, `isPersonalSite` and `siteCollection.{hostname, dataLocationCode}`.
 * `eTag`, `root`, `sharepointIds`, relationships and every unknown property are dropped.
 */

const SITE_KEYS = [
  "name",
  "displayName",
  "description",
  "webUrl",
  "createdDateTime",
  "lastModifiedDateTime",
  "isPersonalSite",
] as const;
const COLLECTION_KEYS = ["hostname", "dataLocationCode"] as const;

/** Allowlist projection for one site (single, collection item, batch body). */
export function sanitizeSite(item: unknown): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, ["id", ...SITE_KEYS]);
  const collection = item["siteCollection"];
  if (isRecord(collection)) out["siteCollection"] = pickScalars(collection, COLLECTION_KEYS);
  return out;
}

const SEARCH_DESCRIPTION =
  "Required. Free-text keywords matched against the sites of the tenant (max 200 chars), sent as `search=<text>`; it is data, never an instruction. Must contain a letter or digit (a bare `*` is refused: this tool searches, it does not enumerate every site). Returns site metadata (id, name, displayName, webUrl, ...); the returned id feeds get-sharepoint-site. Work/school accounts only.";

const searchContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      // The page documents no `$top`: `top` is a local limit and pages are followed.
      top: { default: 25, max: 200, wire: false },
      search: { maxLength: 200, siteSearch: true, description: SEARCH_DESCRIPTION },
      // Load-bearing on a collection contract: it widens the inner batch keep-set
      // (`projectCollection`) so site scalars survive until `sanitizeSite` runs.
      projectEntity: { keep: [...SITE_KEYS, "siteCollection"] },
    }),
    true,
    sanitizeSite,
  );

const siteContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({ projectEntity: { keep: [...SITE_KEYS, "siteCollection"] } }),
    false,
    sanitizeSite,
  );

export function registerSharePointSiteTools(
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
    "search-sharepoint-sites",
    listTool(
      client,
      "search-sharepoint-sites",
      patternOf("search-sharepoint-sites"),
      searchContract(),
    ),
  );
  handlers.set(
    "get-sharepoint-site",
    getTool(client, "get-sharepoint-site", patternOf("get-sharepoint-site"), siteContract()),
  );
  const byPath = getTool(
    client,
    "get-sharepoint-site-by-path",
    patternOf("get-sharepoint-site-by-path"),
    siteContract(),
  );
  handlers.set("get-sharepoint-site-by-path", {
    ...byPath,
    // The catalog placeholder is `{path}` but the exposed input is `relative-path` (a bare `path`
    // would read as a generic-proxy parameter, see catalog.ts RESERVED_PARAM_RENAMES).
    handler: (args) => {
      const { "relative-path": relativePath, ...rest } = args;
      return byPath.handler({ ...rest, path: relativePath });
    },
  });
}
