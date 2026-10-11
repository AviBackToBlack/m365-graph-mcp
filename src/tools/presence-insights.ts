import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { isRecord, pickScalars } from "./chats.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";
import { withSanitizer } from "./teams.js";

/**
 * SM-14 part A: `get-my-presence`, `get-user-presence` and `list-trending-insights` (all v1.0 GET).
 *
 * Delegated permissions (current Microsoft docs, October 2026; see PERMISSIONS.md): presence-get
 * documents `Presence.Read` as least privileged (`/me/presence`) and `Presence.Read.All` as the
 * higher one (`/users/{id}/presence`, per its two examples); personal Microsoft accounts are "Not
 * supported", and the page notes 1,500 requests per 30 seconds per application per tenant.
 * insights-list-trending documents `Sites.Read.All` (work/school only; personal "Not supported";
 * global service only, no national cloud). The server never pre-checks scopes; Graph decides, and
 * a 403 is a `capability_gap`.
 *
 * Presence: no query option is documented or offered. `get-user-presence` takes only a Microsoft
 * Entra user object ID GUID (D39): Learn says "replace {id} with the user's GUID". The GUID-only
 * rule is the route-scoped shape `/users/{user-id}/presence|user-id` in `graph/placeholders.ts`,
 * so the single-path tool and `graph-batch` refuse a UPN/email identically with zero requests,
 * while the shared `user-id` shape of the other `/users/{user-id}/...` routes is unchanged.
 *
 * Trending (`GET /me/insights/trending`): Learn documents `$filter` (e.g. on
 * `resourceVisualization/type` and `containerType`) but this typed surface omits it on purpose
 * (D40); `$orderby` is not documented. The only option is a local `top` limit (1-50, default 25):
 * no `$top` is sent and nextLink pages are followed through the guarded `GraphClient.paginate`
 * (continuations are bound to the original route and query by `validateGraphLink`).
 *
 * Every response (single, list item, batch body) is projected through an allowlist (D40).
 * Presence: `id`, `availability`, `activity`, `outOfOfficeSettings.isOutOfOffice`,
 * `workLocation.workLocationType`; `statusMessage`, `outOfOfficeSettings.message`,
 * `workLocation.placeId`/`source`, `sequenceNumber` and unknown properties are dropped (the status
 * and out-of-office messages are user-authored free text). Trending: `id`, `weight`,
 * `resourceVisualization.{title,type,mediaType,containerDisplayName,containerType,containerWebUrl}`
 * and `resourceReference.{webUrl,type}`; `previewImageUrl`, `previewText`, `resourceReference.id`
 * (a raw relative Graph resource path that callers must not treat as a request target) and
 * unknown properties are dropped. `weight` is kept whether Graph sends a number or a string
 * (Learn's example is a string).
 */

const PRESENCE_KEYS = ["availability", "activity"] as const;
const OOF_KEYS = ["isOutOfOffice"] as const;
const WORK_LOCATION_KEYS = ["workLocationType"] as const;

/** Allowlist projection for one presence (single, batch body). Non-objects yield undefined. */
export function sanitizePresence(item: unknown): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, ["id", ...PRESENCE_KEYS]);
  const oof = item["outOfOfficeSettings"];
  if (isRecord(oof)) out["outOfOfficeSettings"] = pickScalars(oof, OOF_KEYS);
  const location = item["workLocation"];
  if (isRecord(location)) out["workLocation"] = pickScalars(location, WORK_LOCATION_KEYS);
  return out;
}

const VISUALIZATION_KEYS = [
  "title",
  "type",
  "mediaType",
  "containerDisplayName",
  "containerType",
  "containerWebUrl",
] as const;
const REFERENCE_KEYS = ["webUrl", "type"] as const;

/** Allowlist projection for one trending insight (collection item, batch body item). */
export function sanitizeTrending(item: unknown): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, ["id", "weight"]);
  const visualization = item["resourceVisualization"];
  if (isRecord(visualization)) {
    out["resourceVisualization"] = pickScalars(visualization, VISUALIZATION_KEYS);
  }
  const reference = item["resourceReference"];
  if (isRecord(reference)) out["resourceReference"] = pickScalars(reference, REFERENCE_KEYS);
  return out;
}

const presenceContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      // No query option exists. The keep set is load-bearing: it lets the nested objects survive
      // the entity projection until `sanitizePresence` reduces them.
      projectEntity: { keep: [...PRESENCE_KEYS, "outOfOfficeSettings", "workLocation"] },
    }),
    false,
    sanitizePresence,
  );

const trendingContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      // `top` is only a local item limit: no `$top` is sent, pages are followed, and a batch
      // subrequest carrying `$top` (or any other option) is rejected.
      top: { default: 25, max: 50, wire: false },
      // Load-bearing on a collection contract: it widens the inner batch keep-set
      // (`projectCollection`) so these properties survive until `sanitizeTrending` runs.
      projectEntity: { keep: ["weight", "resourceVisualization", "resourceReference"] },
    }),
    true,
    sanitizeTrending,
  );

export function registerPresenceInsightTools(
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
  for (const name of ["get-my-presence", "get-user-presence"]) {
    handlers.set(name, getTool(client, name, patternOf(name), presenceContract()));
  }
  handlers.set(
    "list-trending-insights",
    listTool(
      client,
      "list-trending-insights",
      patternOf("list-trending-insights"),
      trendingContract(),
    ),
  );
}
