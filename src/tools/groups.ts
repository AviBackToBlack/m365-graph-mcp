import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import {
  REPORTS_TOP_DEFAULT,
  directoryCollectionTool,
  projectDirectoryObject,
} from "./directory.js";
import { eventListContract, eventWindowContract } from "./calendar.js";
import { USER_PROPERTIES } from "./identity.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";

/**
 * Group directory tools (SM-10 part A): `list-groups`, `get-group`, `list-group-members` and
 * `list-group-owners` (all v1.0 GET).
 *
 * Delegated permissions (current Microsoft docs; see PERMISSIONS.md): list/get keep
 * `Group.Read.All` (D12); members use `GroupMember.ReadBasic.All` with best-effort per-type
 * details (D13); owners use `GroupMember.Read.All`. None supports personal accounts. Hidden-membership groups
 * additionally need `Member.Read.Hidden`, which is not requested, so they surface as 403
 * (`capability_gap`).
 *
 * `list-groups` offers `top`, a typed `select` and a typed directory `search` (D9 pattern: the
 * server itself sets `$count` and `ConsistencyLevel: eventual`). `$filter`, `$orderby`,
 * `$expand`, `$skip` and `$count` are deliberately not offered.
 *
 * `get-group` offers only a typed `select` over properties Graph returns by default; the
 * `$select`-only Exchange-backed properties (accessType, unseen counts, ...) are excluded.
 *
 * Group calendar tools (SM-10 part B): `get-group-calendar-view` and `list-group-events` reuse
 * the user-calendar contracts from calendar.ts; both need only `Group.Read.All` (delegated,
 * work/school; no personal-account or application support for list events).
 *
 * Legacy group conversations/threads (SM-10 part C): `list-group-conversations` and
 * `list-group-threads` use `Group-Conversation.Read.All`, the documented least privilege (D16).
 * Typed `top` (local) and `select` only.
 *
 * Members and owners are `directoryObject` collections (user, group, orgContact, device,
 * servicePrincipal, ...). Each item is projected PER `@odata.type` and per request
 * (`projectGroupMember`). The pages' query-option note ("supported only with advanced query
 * parameters") contradicts their own header table, so no `$top` is sent: `top` is a local limit
 * with pages followed (as for directReports/memberOf), and `$filter`/`$search`/`$orderby`/OData
 * cast/`$count`/`$expand` are not offered. Members/owners of other types are not expanded.
 */

const GROUP_TYPE = "#microsoft.graph.group";
const USER_TYPE = "#microsoft.graph.user";
const ORG_CONTACT_TYPE = "#microsoft.graph.orgContact";

/**
 * Selectable group properties: only ones Graph returns by default on BOTH list and get.
 * Deliberately EXCLUDED: `$select`-only Exchange-backed properties (allowExternalSenders,
 * hideFrom*, isSubscribedByMail, unseen*, accessType, ...; get-only), licensing and
 * on-premises detail (assignedLicenses, onPremisesSecurityIdentifier, provisioning errors,
 * netBios/SAM names), `proxyAddresses`, `theme`, `preferredDataLocation`, `securityIdentifier`.
 */
export const GROUP_PROPERTIES = [
  "id",
  "displayName",
  "description",
  "mail",
  "mailEnabled",
  "mailNickname",
  "securityEnabled",
  "groupTypes",
  "visibility",
  "createdDateTime",
  "classification",
  "membershipRule",
  "membershipRuleProcessingState",
  "expirationDateTime",
  "renewedDateTime",
  "isAssignableToRole",
  "resourceProvisioningOptions",
  "preferredLanguage",
  "onPremisesSyncEnabled",
  "onPremisesLastSyncDateTime",
] as const;

/** Compact default; `groupTypes`/`mailEnabled`/`securityEnabled` let callers classify a group. */
export const GROUP_DEFAULT_PROPERTIES = [
  "displayName",
  "description",
  "mail",
  "mailEnabled",
  "securityEnabled",
  "groupTypes",
  "visibility",
  "createdDateTime",
  "membershipRule",
] as const;

/** `$search` is tokenized on these two group properties only (List groups). */
const GROUP_SEARCH_PROPERTIES = ["displayName", "description"] as const;

const GROUP_SET: ReadonlySet<string> = new Set(GROUP_PROPERTIES);

// Union of every projected member type's properties (user, orgContact via the user list, group).
// `id` is selectable too (an ID-only payload); the projector always keeps it anyway.
const MEMBER_SELECT = [...new Set(["id", ...USER_PROPERTIES, ...GROUP_PROPERTIES])] as [
  string,
  ...string[],
];

/**
 * Legacy group conversations/threads (SM-10 part C): `list-group-conversations` and
 * `list-group-threads`. Metadata only: no posts, bodies or attachments (the post/thread-detail
 * paths are not in the catalog). Properties are the documented v1.0 `conversation` /
 * `conversationThread` ones. `preview` (a short excerpt of the latest post, the analogue of the
 * mail list `bodyPreview`) and `uniqueSenders` (sender display names) are in the default, as the
 * mail list tools do for `bodyPreview`/`from`; thread `toRecipients`/`ccRecipients` need an
 * explicit `$select` at Graph ("Requires $select to retrieve") so they are opt-in and never
 * default. `$top` is not documented on either page, so `top` is a local limit with pages
 * followed (no `$top` is sent); `$filter`/`$search`/`$orderby`/`$skip`/`$expand` are not offered.
 */
const CONVERSATION_PROPERTIES = [
  "id",
  "topic",
  "hasAttachments",
  "lastDeliveredDateTime",
  "uniqueSenders",
  "preview",
] as const;

const THREAD_PROPERTIES = [
  ...CONVERSATION_PROPERTIES,
  "isLocked",
  "toRecipients",
  "ccRecipients",
] as const;

const THREAD_DEFAULT_PROPERTIES = [...CONVERSATION_PROPERTIES, "isLocked"] as const;

const legacyContract = (
  allowed: readonly [string, ...string[]],
  defaults: readonly [string, ...string[]],
): QueryContract =>
  defineQueryContract({
    top: { default: REPORTS_TOP_DEFAULT, max: 100, wire: false },
    select: { allowed, default: defaults, requiredInBatch: false },
  });

const listContract = (): QueryContract =>
  defineQueryContract({
    top: { default: 25, max: 200, pageMax: 100 },
    select: {
      allowed: GROUP_PROPERTIES,
      default: GROUP_DEFAULT_PROPERTIES,
      requiredInBatch: false,
    },
    search: {
      maxLength: 200,
      directoryProperties: GROUP_SEARCH_PROPERTIES,
      description: `Directory search: ONE "property:text" clause, property one of ${GROUP_SEARCH_PROPERTIES.join(", ")} (e.g. displayName:marketing). Text may not contain quotes, backslashes, colons, parentheses, ampersands or the KQL operators AND, OR, NOT. The server sends it as $search="property:text" and sets the ConsistencyLevel header and $count itself. Cannot be used through graph-batch. Results are search-ranked, not sortable.`,
    },
  });

const getContract = (): QueryContract =>
  defineQueryContract({
    select: {
      allowed: GROUP_PROPERTIES,
      default: GROUP_DEFAULT_PROPERTIES,
      requiredInBatch: false,
    },
    projectEntity: {},
  });

const membersContract = (): QueryContract =>
  defineQueryContract({
    top: { default: REPORTS_TOP_DEFAULT, max: 100, wire: false },
    // No default: nothing is sent unless the caller chooses (Graph's default set applies).
    select: { allowed: MEMBER_SELECT, requiredInBatch: false },
  });

/**
 * Keeps `id`, `@odata.type` and the item type's allowlisted properties (restricted to the
 * request's `$select` when one was sent, else the type's default set). `group` has the group
 * allowlist; `user` and `orgContact` follow `projectDirectoryObject`; any other or missing type
 * (device, servicePrincipal, limited-information objects, ...) keeps `id` and `@odata.type`
 * alone. Non-objects yield undefined. Graph-supplied
 * data is untrusted: nothing outside the allowlists survives.
 */
export function projectGroupMember(item: unknown, selectWire: string | undefined): unknown {
  if (typeof item !== "object" || item === null || Array.isArray(item)) return undefined;
  const obj = item as Record<string, unknown>;
  const type = obj["@odata.type"];
  if (type === USER_TYPE || type === ORG_CONTACT_TYPE) {
    return projectDirectoryObject(item, selectWire);
  }
  if (type !== GROUP_TYPE) {
    // Any other or missing type fails closed (as projectMembership): a typeless item could be a
    // group, device or service principal, so the user/orgContact common set does not apply.
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(obj)) {
      if (k === "id" || k === "@odata.type") out[k] = v;
    }
    return out;
  }
  const wanted =
    selectWire === undefined || selectWire.length === 0
      ? GROUP_DEFAULT_PROPERTIES
      : selectWire.split(",");
  const keys = new Set(["id", ...wanted.filter((n) => GROUP_SET.has(n))]);
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (keys.has(k) || k === "@odata.type") out[k] = v;
  }
  return out;
}

export function registerGroupTools(
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
    "list-groups",
    listTool(client, "list-groups", patternOf("list-groups"), listContract()),
  );
  handlers.set("get-group", getTool(client, "get-group", patternOf("get-group"), getContract()));
  for (const name of ["list-group-members", "list-group-owners"]) {
    handlers.set(
      name,
      directoryCollectionTool(client, name, patternOf(name), membersContract(), projectGroupMember),
    );
  }
  // SM-10 part B: the user-calendar list/window contracts unchanged (body-free projection,
  // typed `timeZone`, bounded window), bound to a group's default calendar. The catalog does
  // not flag `supportsExpandExtendedProperties` on these entries, so no `extendedPropertyId`.
  handlers.set(
    "get-group-calendar-view",
    listTool(
      client,
      "get-group-calendar-view",
      patternOf("get-group-calendar-view"),
      eventWindowContract(false),
    ),
  );
  handlers.set(
    "list-group-events",
    listTool(client, "list-group-events", patternOf("list-group-events"), eventListContract(false)),
  );
  // SM-10 part C: legacy conversations and threads of a group (metadata only).
  handlers.set(
    "list-group-conversations",
    listTool(
      client,
      "list-group-conversations",
      patternOf("list-group-conversations"),
      legacyContract(CONVERSATION_PROPERTIES, CONVERSATION_PROPERTIES),
    ),
  );
  handlers.set(
    "list-group-threads",
    listTool(
      client,
      "list-group-threads",
      patternOf("list-group-threads"),
      legacyContract(THREAD_PROPERTIES, THREAD_DEFAULT_PROPERTIES),
    ),
  );
}
