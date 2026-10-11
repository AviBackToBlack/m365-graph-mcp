import type { EndpointCatalog } from "../catalog.js";
import { successResult } from "../envelope.js";
import { GraphClientError } from "../graph/errors.js";
import type { GraphClient, GraphRequest } from "../graph/client.js";
import type { BatchPolicy } from "../graph/batch.js";
import type { ToolImplementation } from "../registry.js";
import { USER_DEFAULT_PROPERTIES, USER_PROPERTIES } from "./identity.js";
import { failure, listTool, pathParams } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";

/**
 * Directory tools (SM-9 part B): `list-users` and the manager / direct-report navigations.
 *
 * Delegated permissions (current Microsoft docs; `endpoints.json` scopes are non-normative and
 * unchanged, see PERMISSIONS.md): manager `User.Read.All`; directReports `User.Read` and
 * `User.ReadBasic.All`; list-users `User.ReadBasic.All`. None supports personal accounts.
 *
 * `list-users` returns `user` entities, so it uses the plain list contract (typed `select`,
 * bounded `top`, and a typed directory `search` for which the server itself sets the required
 * advanced-query headers). `$filter`/`$orderby`/`$expand`/`$skip` are deliberately not offered.
 *
 * Manager and directReports are typed `directoryObject`: each item may be a `user`, an
 * `orgContact`, or (limited information) a bare id + `@odata.type`. Responses are therefore
 * projected PER ITEM TYPE and per request (see `projectDirectoryObject`). No `$select` is sent
 * unless the caller chose one (a user-only `$select` on a directoryObject navigation is the
 * documented-but-risky path; Graph's own default set is what comes back otherwise).
 */

const USER_TYPE = "#microsoft.graph.user";
const ORG_CONTACT_TYPE = "#microsoft.graph.orgContact";

/** orgContact properties (v1.0) that overlap the user allowlist's names. */
const ORG_CONTACT_PROPERTIES: ReadonlySet<string> = new Set([
  "displayName",
  "givenName",
  "surname",
  "jobTitle",
  "department",
  "companyName",
  "mail",
  "mailNickname",
  "proxyAddresses",
  "onPremisesSyncEnabled",
  "onPremisesLastSyncDateTime",
]);
const ORG_CONTACT_DEFAULTS = [
  "displayName",
  "givenName",
  "surname",
  "jobTitle",
  "department",
  "companyName",
  "mail",
  "mailNickname",
] as const;
const USER_SET: ReadonlySet<string> = new Set(USER_PROPERTIES);
const COMMON_SET: ReadonlySet<string> = new Set(
  USER_PROPERTIES.filter((p) => ORG_CONTACT_PROPERTIES.has(p)),
);

/** Directory `$search` is one `property:text` clause over these (users: documented example). */
const USER_SEARCH_PROPERTIES = ["displayName", "mail", "userPrincipalName"] as const;

/**
 * Keeps `id`, `@odata.type` and the type's allowlisted properties (restricted to the request's
 * `$select` when one was sent, else the type's default set). `user` -> user allowlist;
 * `orgContact` -> orgContact allowlist; a missing type -> the common subset only; any other
 * type (group, device, ...) -> `id` and `@odata.type` alone. Non-objects yield undefined.
 * Graph-supplied data is untrusted: nothing outside the allowlists survives.
 */
export function projectDirectoryObject(item: unknown, selectWire: string | undefined): unknown {
  if (typeof item !== "object" || item === null || Array.isArray(item)) return undefined;
  const obj = item as Record<string, unknown>;
  const type = obj["@odata.type"];
  let allowed: ReadonlySet<string>;
  let defaults: readonly string[];
  if (type === USER_TYPE) {
    allowed = USER_SET;
    defaults = USER_DEFAULT_PROPERTIES;
  } else if (type === ORG_CONTACT_TYPE) {
    allowed = ORG_CONTACT_PROPERTIES;
    defaults = ORG_CONTACT_DEFAULTS;
  } else if (type === undefined) {
    allowed = COMMON_SET;
    defaults = ORG_CONTACT_DEFAULTS;
  } else {
    allowed = new Set();
    defaults = [];
  }
  const wanted =
    selectWire === undefined || selectWire.length === 0 ? defaults : selectWire.split(",");
  const keys = new Set(["id", ...wanted.filter((n) => allowed.has(n))]);
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (keys.has(k) || k === "@odata.type") out[k] = v;
  }
  return out;
}

const selectContract = (spec: { top?: boolean }): QueryContract =>
  defineQueryContract({
    ...(spec.top === true ? { top: { default: REPORTS_TOP_DEFAULT, max: 100, wire: false } } : {}),
    // No default: nothing is sent unless the caller chooses (see module comment).
    select: { allowed: USER_PROPERTIES, requiredInBatch: false },
  });

const usersContract = (): QueryContract =>
  defineQueryContract({
    top: { default: 25, max: 200, pageMax: 100 },
    select: { allowed: USER_PROPERTIES, default: USER_DEFAULT_PROPERTIES, requiredInBatch: false },
    search: {
      maxLength: 200,
      directoryProperties: USER_SEARCH_PROPERTIES,
      description: `Directory search: ONE "property:text" clause, property one of ${USER_SEARCH_PROPERTIES.join(", ")} (e.g. displayName:john). Text may not contain quotes, backslashes, colons, parentheses, ampersands or the KQL operators AND, OR, NOT. The server sends it as $search="property:text" and sets the ConsistencyLevel header and $count itself. Cannot be used through graph-batch. Results are search-ranked, not sortable.`,
    },
  });

const selectOf = (pairs: ReadonlyArray<readonly [string, string]>): string | undefined =>
  pairs.find(([n]) => n === "$select")?.[1];

function managerTool(
  client: GraphClient,
  toolName: string,
  pathPattern: string,
): ToolImplementation {
  const contract = selectContract({});
  const batchPolicy: BatchPolicy = {
    check: contract.batchPolicy.check,
    // Same projector as the single path; a non-object body becomes {} (never raw).
    project: (body, pairs) => projectDirectoryObject(body, selectOf(pairs)) ?? {},
  };
  return {
    inputSchema: contract.shape,
    batchPolicy,
    handler: async (args) => {
      try {
        const { query, headers } = contract.toRequest(args);
        const res = await client.request({
          pathPattern,
          pathParams: pathParams(pathPattern, args),
          query,
          headers,
        });
        const select = query["$select"];
        const projected = projectDirectoryObject(
          res.data,
          typeof select === "string" ? select : undefined,
        );
        if (projected === undefined) {
          throw new GraphClientError("graph_error", "Graph returned an unexpected response shape.");
        }
        const meta: Record<string, unknown> = { clientRequestId: res.clientRequestId };
        if (res.requestId !== undefined) meta["requestId"] = res.requestId;
        return successResult(projected, meta);
      } catch (err) {
        return failure(toolName, err);
      }
    },
  };
}

/** directReports `top` default; also the batch clip, since `$top` never reaches the wire. */
export const REPORTS_TOP_DEFAULT = 25;

/**
 * A directoryObject navigation collection with a local `top` limit and no `$top` on the wire.
 * Directory reports and `list-my-memberships` share it; `project` is the per-item-type
 * projector and `contract` the typed options (`top` there must use `wire: false`).
 */
export function directoryCollectionTool(
  client: GraphClient,
  toolName: string,
  pathPattern: string,
  contract: QueryContract,
  project: (item: unknown, selectWire: string | undefined) => unknown,
): ToolImplementation {
  // `$top` is never on the wire, so a batch page is clipped to THIS contract's `top` default
  // (the limit an argument-less single call gets), keeping both paths in step per tool.
  const batchLimit = contract.toRequest({}).limit;
  const batchPolicy: BatchPolicy = {
    check: contract.batchPolicy.check,
    project: (body, pairs) => {
      if (typeof body !== "object" || body === null || Array.isArray(body)) return {};
      const { value } = body as Record<string, unknown>;
      if (!Array.isArray(value)) return {};
      const limit = batchLimit;
      const select = selectOf(pairs);
      const items = value
        .slice(0, limit)
        .map((i) => project(i, select))
        .filter((i) => i !== undefined);
      const more =
        value.length > limit || (body as Record<string, unknown>)["@odata.nextLink"] !== undefined;
      return more ? { value: items, hasMore: true } : { value: items };
    },
  };
  return {
    inputSchema: contract.shape,
    batchPolicy,
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
        const select = query["$select"];
        const items = result.items
          .map((i) => project(i, typeof select === "string" ? select : undefined))
          .filter((i) => i !== undefined);
        return successResult(items, {
          count: items.length,
          hasMore: result.truncated,
          pages: result.pages,
        });
      } catch (err) {
        return failure(toolName, err);
      }
    },
  };
}

function directReportsTool(
  client: GraphClient,
  toolName: string,
  pathPattern: string,
): ToolImplementation {
  return directoryCollectionTool(
    client,
    toolName,
    pathPattern,
    selectContract({ top: true }),
    projectDirectoryObject,
  );
}

export function registerDirectoryTools(
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
    "list-users",
    listTool(client, "list-users", patternOf("list-users"), usersContract()),
  );
  for (const name of ["get-my-manager", "get-user-manager"]) {
    handlers.set(name, managerTool(client, name, patternOf(name)));
  }
  for (const name of ["list-my-direct-reports", "list-user-direct-reports"]) {
    handlers.set(name, directReportsTool(client, name, patternOf(name)));
  }
}
