import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { isRecord, pickScalars, projectChatMember } from "./chats.js";
import { directoryCollectionTool } from "./directory.js";
import { getTool, listTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";

/**
 * Teams and channels (SM-12 part A): `get-team`, `list-team-channels`, `get-team-channel` and
 * `list-team-members` (all v1.0 GET). Channel messages, replies, hosted contents, tabs, the files
 * folder and the capability matrix are later parts.
 *
 * Delegated permissions (current Microsoft docs, October 2026; see PERMISSIONS.md): get-team
 * `Team.ReadBasic.All`; list/get channel `Channel.ReadBasic.All`; list members
 * `TeamMember.Read.All`. None supports personal Microsoft accounts.
 *
 * Offered options are typed only. `get-team` and both channel tools take a typed `select` over an
 * allowlist of scalar properties; the channel `email` and the `summary` objects are opt-in because
 * Graph documents them as expensive and `$select`-only respectively. `list-team-channels` and
 * `list-team-members` take a local `top` (no `$top` is sent: the channel page documents none, and the
 * member contract mirrors list-chat-members so `graph-batch` and the single tool clip identically).
 * Not offered anywhere: `$expand`, `$filter` (no raw OData; a typed `membershipType` filter is
 * deliberately left out, see PERMISSIONS.md), `$search`, `$orderby`, `$skip`, `$count`.
 *
 * Every response is projected locally (single path and batch alike): scalars only, `summary` reduced
 * to its counts, and the team `*Settings` objects, `internalId`, `classSettings`, channel
 * `moderationSettings`, `migrationMode` and every unknown property or annotation dropped.
 * Members are projected per `@odata.type` exactly like chat members (`projectChatMember`).
 */

const TEAM_PROPERTIES = [
  "displayName",
  "description",
  "classification",
  "specialization",
  "visibility",
  "webUrl",
  "isArchived",
  "createdDateTime",
  "tenantId",
  "summary",
] as const;
const TEAM_DEFAULT_PROPERTIES = [
  "displayName",
  "description",
  "classification",
  "specialization",
  "visibility",
  "webUrl",
  "isArchived",
  "createdDateTime",
] as const;

/** `layoutType` is always null in the channel LIST (documented known issue), so list omits it. */
const CHANNEL_LIST_PROPERTIES = [
  "displayName",
  "description",
  "createdDateTime",
  "membershipType",
  "isArchived",
  "isFavoriteByDefault",
  "webUrl",
  "email",
] as const;
/** Everything except `email` (documented as expensive to populate) is in the default. */
const CHANNEL_LIST_DEFAULT_PROPERTIES = [
  "displayName",
  "description",
  "createdDateTime",
  "membershipType",
  "isArchived",
  "isFavoriteByDefault",
  "webUrl",
] as const;
const CHANNEL_GET_PROPERTIES = [...CHANNEL_LIST_PROPERTIES, "layoutType", "summary"] as const;
const CHANNEL_GET_DEFAULT_PROPERTIES = [...CHANNEL_LIST_DEFAULT_PROPERTIES, "layoutType"] as const;

/** Scalar properties that can survive projection (a union; the request's select narrows it). */
const SCALAR_KEYS = [
  "id",
  ...new Set([...TEAM_PROPERTIES, ...CHANNEL_GET_PROPERTIES].filter((k) => k !== "summary")),
] as readonly string[];
/** `teamSummary`: owner/member/guest counts only. */
const TEAM_SUMMARY = ["ownersCount", "membersCount", "guestsCount"] as const;
/** `channelSummary` additionally carries `hasMembersFromOtherTenants` (channel-only). */
const CHANNEL_SUMMARY = [...TEAM_SUMMARY, "hasMembersFromOtherTenants"] as const;

/**
 * Last-line sanitizer applied AFTER the contract's own per-request projection: scalars only (an
 * object where a scalar is expected is dropped) and `summary` reduced to its count fields.
 * Non-objects yield undefined.
 */
export function sanitizeTeamsEntity(item: unknown, kind: "team" | "channel"): unknown {
  if (!isRecord(item)) return undefined;
  const out = pickScalars(item, SCALAR_KEYS);
  if (Object.hasOwn(item, "summary")) {
    const summary = item["summary"];
    if (summary === null) out["summary"] = null;
    else if (isRecord(summary)) {
      out["summary"] = pickScalars(summary, kind === "team" ? TEAM_SUMMARY : CHANNEL_SUMMARY);
    }
  }
  return out;
}

/**
 * Composes `sanitize` over EVERY projection path of `base` (single entity, collection item, batch
 * body or batch collection items). The base projection runs first, so the request's effective
 * `$select` still trims allowlisted-but-unselected fields identically on both paths.
 */
export function withSanitizer(
  base: QueryContract,
  collection: boolean,
  sanitize: (item: unknown) => unknown,
): QueryContract {
  const inner = base.batchPolicy.project;
  return {
    ...base,
    projectItem: (item) => sanitize(base.projectItem(item)),
    projectResponse: (item, query) => sanitize(base.projectResponse(item, query)),
    itemProjector: (query) => {
      const project = base.itemProjector(query);
      return (item) => sanitize(project(item));
    },
    batchPolicy: {
      check: base.batchPolicy.check,
      project: (body, pairs) => {
        const out = inner === undefined ? body : inner(body, pairs);
        if (!collection) return sanitize(out) ?? {};
        if (!isRecord(out) || !Array.isArray(out["value"])) return {};
        const value = out["value"].map(sanitize).filter((i) => i !== undefined);
        return { ...out, value };
      },
    },
  };
}

const teamContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      select: {
        allowed: TEAM_PROPERTIES,
        default: TEAM_DEFAULT_PROPERTIES,
        requiredInBatch: false,
      },
      projectEntity: {},
    }),
    false,
    (item) => sanitizeTeamsEntity(item, "team"),
  );

const channelListContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      // The page documents no `$top`: `top` is a local limit and pages are followed.
      top: { default: 50, max: 200, wire: false },
      select: {
        allowed: CHANNEL_LIST_PROPERTIES,
        default: CHANNEL_LIST_DEFAULT_PROPERTIES,
        requiredInBatch: true,
      },
    }),
    true,
    (item) => sanitizeTeamsEntity(item, "channel"),
  );

const channelGetContract = (): QueryContract =>
  withSanitizer(
    defineQueryContract({
      select: {
        allowed: CHANNEL_GET_PROPERTIES,
        default: CHANNEL_GET_DEFAULT_PROPERTIES,
        requiredInBatch: true,
      },
      projectEntity: {},
    }),
    false,
    (item) => sanitizeTeamsEntity(item, "channel"),
  );

const membersContract = (): QueryContract =>
  // No `$top` on the wire: `top` is a local limit (pages followed), as for list-chat-members.
  defineQueryContract({ top: { default: 100, max: 250, wire: false } });

export function registerTeamTools(
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
  handlers.set("get-team", getTool(client, "get-team", patternOf("get-team"), teamContract()));
  handlers.set(
    "list-team-channels",
    listTool(client, "list-team-channels", patternOf("list-team-channels"), channelListContract()),
  );
  handlers.set(
    "get-team-channel",
    getTool(client, "get-team-channel", patternOf("get-team-channel"), channelGetContract()),
  );
  handlers.set(
    "list-team-members",
    directoryCollectionTool(
      client,
      "list-team-members",
      patternOf("list-team-members"),
      membersContract(),
      projectChatMember,
    ),
  );
}
