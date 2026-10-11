import type { EndpointCatalog } from "../catalog.js";
import { successResult } from "../envelope.js";
import type { ToolImplementation } from "../registry.js";
import type { GraphClient, GraphRequest } from "../graph/client.js";
import { REPORTS_TOP_DEFAULT, directoryCollectionTool } from "./directory.js";
import { failure, pathParams } from "./mail.js";
import { defineQueryContract, nextSkipFor, type QueryContract } from "./query-contract.js";

/** People offset paging: the largest `skip` accepted (and so the largest nextSkip offered). */
const PEOPLE_SKIP_MAX = 1000;

/**
 * People and membership tools (SM-9 part C): `list-relevant-people` (`GET /me/people`, v1.0,
 * `person`) and `list-my-memberships` (`GET /me/memberOf`, v1.0, `directoryObject`).
 *
 * Delegated permissions (current Microsoft docs; see PERMISSIONS.md): people `People.Read`
 * (matches the catalog; work/school and personal accounts); memberOf `User.Read` (D10, work/school
 * only): per-type details are best-effort and an unreadable type degrades to id + @odata.type.
 *
 * `list-relevant-people` offers `top`, `skip`, a typed `select` and a plain fuzzy `search`
 * (never combined with `skip`, as documented). `$filter`, `$orderby` and the
 * `X-PeopleQuery-QuerySources` header are deliberately not offered.
 *
 * `list-my-memberships` returns `directoryObject`s of several types, so each item is projected
 * PER `@odata.type` and per request (`projectMembership`). The memberOf page documents no `$top`
 * (and `$filter`, `$search`, `$orderby`, OData cast and `$count` all need advanced query mode),
 * so `top` is a local limit with pages followed, as for directReports, and none of those
 * options is offered. A "groups only" cast is another path (`/me/memberOf/microsoft.graph.group`)
 * that is not in the catalog; callers can read `@odata.type` instead.
 */

/**
 * Selectable v1.0 `person` properties. Deliberately EXCLUDED: `personNotes` (the user's private
 * free-form notes about a person), `birthday`, `yomiCompany`, `profession`, `websites`,
 * `postalAddresses` (rarely populated, not needed to identify or contact a person) and
 * `isFavorite`.
 */
export const PERSON_PROPERTIES = [
  "id",
  "displayName",
  "givenName",
  "surname",
  "jobTitle",
  "companyName",
  "department",
  "officeLocation",
  "userPrincipalName",
  "scoredEmailAddresses",
  "phones",
  "personType",
  "imAddress",
] as const;

export const PERSON_DEFAULT_PROPERTIES = [
  "displayName",
  "scoredEmailAddresses",
  "jobTitle",
  "department",
  "officeLocation",
  "personType",
] as const;

const peopleContract = (): QueryContract =>
  defineQueryContract({
    top: { default: 10, max: 100, offsetPaging: true },
    select: {
      allowed: PERSON_PROPERTIES,
      default: PERSON_DEFAULT_PROPERTIES,
      requiredInBatch: false,
    },
    skip: { max: PEOPLE_SKIP_MAX },
    search: {
      maxLength: 100,
      plainName: true,
      description:
        'Fuzzy name or e-mail search over the relevant people (plain text such as tyler lee or tylerle@example.com; typos are tolerated). Allowed characters: letters, digits, space and . @ \' _ - only (no quotes, colons or operators). The server sends it as $search="text". It is not KQL and cannot be combined with skip. Results are relevance-ranked.',
    },
  });

/** Directory object types `/me/memberOf` returns that this tool projects. */
const GROUP_TYPE = "#microsoft.graph.group";
const ROLE_TYPE = "#microsoft.graph.directoryRole";
const ADMIN_UNIT_TYPE = "#microsoft.graph.administrativeUnit";

interface MembershipShape {
  allowed: ReadonlySet<string>;
  defaults: readonly string[];
}

const GROUP_SHAPE: MembershipShape = {
  allowed: new Set([
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
  ]),
  defaults: ["displayName", "description", "mail", "mailEnabled", "securityEnabled", "groupTypes"],
};
const ROLE_SHAPE: MembershipShape = {
  allowed: new Set(["displayName", "description", "roleTemplateId"]),
  defaults: ["displayName", "description", "roleTemplateId"],
};
const ADMIN_UNIT_SHAPE: MembershipShape = {
  allowed: new Set([
    "displayName",
    "description",
    "visibility",
    "membershipType",
    "membershipRule",
    "membershipRuleProcessingState",
    "isMemberManagementRestricted",
  ]),
  defaults: ["displayName", "description", "visibility"],
};
// `id` is selectable too (an ID-only payload); the projector always keeps it anyway.
const MEMBERSHIP_SELECT = [
  ...new Set([
    "id",
    ...[GROUP_SHAPE, ROLE_SHAPE, ADMIN_UNIT_SHAPE].flatMap((shape) => [...shape.allowed]),
  ]),
] as [string, ...string[]];

const NO_SHAPE: MembershipShape = { allowed: new Set(), defaults: [] };

/**
 * Keeps `id`, `@odata.type` and the item type's allowlisted properties (restricted to the
 * request's `$select` when one was sent, else the type's default set). `group`,
 * `directoryRole` and `administrativeUnit` have their own allowlists; any other or missing type
 * (limited-information objects, devices, ...) keeps `id` and `@odata.type` alone. Non-objects
 * yield undefined. Graph-supplied data is untrusted: nothing outside the allowlists survives.
 */
export function projectMembership(item: unknown, selectWire: string | undefined): unknown {
  if (typeof item !== "object" || item === null || Array.isArray(item)) return undefined;
  const obj = item as Record<string, unknown>;
  const type = obj["@odata.type"];
  const shape =
    type === GROUP_TYPE
      ? GROUP_SHAPE
      : type === ROLE_TYPE
        ? ROLE_SHAPE
        : type === ADMIN_UNIT_TYPE
          ? ADMIN_UNIT_SHAPE
          : NO_SHAPE;
  const wanted =
    selectWire === undefined || selectWire.length === 0 ? shape.defaults : selectWire.split(",");
  const keys = new Set(["id", ...wanted.filter((n) => shape.allowed.has(n))]);
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (keys.has(k) || k === "@odata.type") out[k] = v;
  }
  return out;
}

const membershipsContract = (): QueryContract =>
  defineQueryContract({
    top: { default: REPORTS_TOP_DEFAULT, max: 100, wire: false },
    // No default: nothing is sent unless the caller chooses (Graph's default set applies).
    select: { allowed: MEMBERSHIP_SELECT, requiredInBatch: false },
  });

/**
 * `/me/people` pages by OFFSET (`$top`/`$skip`); Graph does not reliably return a nextLink,
 * so a missing link does not prove the listing is complete. Like `listTool`, but a result
 * that fills `top` reports `hasMore: true` with `nextSkip` (pass it back as `skip`); a short
 * page means no more. `hasMore` can therefore be true when exactly `top` people remain.
 * nextLinks, when present, are still followed and validated by `paginate`.
 */
function peopleTool(
  client: GraphClient,
  toolName: string,
  pathPattern: string,
  contract: QueryContract,
): ToolImplementation {
  return {
    inputSchema: contract.shape,
    batchPolicy: contract.batchPolicy,
    handler: async (args) => {
      try {
        const { query, headers, limit } = contract.toRequest(args);
        const base: GraphRequest = {
          pathPattern,
          pathParams: pathParams(pathPattern, args),
          query,
          headers,
        };
        const result = await client.paginate(base, { maxItems: limit });
        const project = contract.itemProjector(query);
        const items = result.items.map(project).filter((item) => item !== undefined);
        const hasMore = result.truncated || result.items.length >= limit;
        const meta: Record<string, unknown> = { count: items.length, hasMore, pages: result.pages };
        if (hasMore && args["search"] === undefined) {
          const skip = typeof args["skip"] === "number" ? args["skip"] : 0;
          // Omitted past the skip maximum: the caller could not pass it back.
          const next = nextSkipFor(skip, result.items.length, PEOPLE_SKIP_MAX);
          if (next !== undefined) meta["nextSkip"] = next;
        }
        return successResult(items, meta);
      } catch (err) {
        return failure(toolName, err);
      }
    },
  };
}

export function registerPeopleTools(
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
    "list-relevant-people",
    peopleTool(client, "list-relevant-people", patternOf("list-relevant-people"), peopleContract()),
  );
  handlers.set(
    "list-my-memberships",
    directoryCollectionTool(
      client,
      "list-my-memberships",
      patternOf("list-my-memberships"),
      membershipsContract(),
      projectMembership,
    ),
  );
}
