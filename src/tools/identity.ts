import type { EndpointCatalog } from "../catalog.js";
import type { GraphClient } from "../graph/client.js";
import type { ToolImplementation } from "../registry.js";
import { getTool } from "./mail.js";
import { defineQueryContract, type QueryContract } from "./query-contract.js";

/**
 * Identity tools (SM-9 part A): `get-current-user`, `get-my-profile`, `get-mailbox-settings`.
 * All three are single-entity GETs for the signed-in user, built from the shared query
 * contract with strict entity projection (only documented properties leave the tool; a
 * non-object response fails closed; `@odata.context` is dropped).
 *
 * Delegated permissions (current Microsoft docs; `endpoints.json` scopes are non-normative,
 * see PERMISSIONS.md): `User.Read` for `/me` and `/me/profile`, `MailboxSettings.Read` for
 * `/me/mailboxSettings`.
 */

/** The eleven properties `GET /me` returns by default (v1.0 `user: get`). */
export const USER_DEFAULT_PROPERTIES = [
  "id",
  "displayName",
  "givenName",
  "surname",
  "userPrincipalName",
  "mail",
  "jobTitle",
  "businessPhones",
  "mobilePhone",
  "officeLocation",
  "preferredLanguage",
] as const;

/**
 * Selectable v1.0 `user` properties. Deliberately EXCLUDED: `mailboxSettings` (own tool and
 * permission), `passwordProfile`/`passwordPolicies`, `customSecurityAttributes`,
 * `employeeLeaveDateTime`, `signInActivity` and `accountEnabled` (extra roles/permissions), licensing, security
 * identifiers and identities, extension properties, and the SharePoint-backed `aboutMe`,
 * `skills`, ... (`$select`-only and not reliably available; not needed for identity).
 */
export const USER_PROPERTIES = [
  ...USER_DEFAULT_PROPERTIES,
  "mailNickname",
  "faxNumber",
  "department",
  "companyName",
  "employeeId",
  "employeeType",
  "employeeHireDate",
  "employeeOrgData",
  "userType",
  "createdDateTime",
  "city",
  "state",
  "country",
  "streetAddress",
  "postalCode",
  "usageLocation",
  "imAddresses",
  "proxyAddresses",
  "onPremisesSamAccountName",
  "onPremisesUserPrincipalName",
  "onPremisesDomainName",
  "onPremisesSyncEnabled",
  "onPremisesLastSyncDateTime",
] as const;

/** The nine documented v1.0 `mailboxSettings` properties; nothing else is returned. */
export const MAILBOX_SETTINGS_PROPERTIES = [
  "archiveFolder",
  "automaticRepliesSetting",
  "dateFormat",
  "delegateMeetingMessageDeliveryOptions",
  "language",
  "timeFormat",
  "timeZone",
  "userPurpose",
  "workingHours",
] as const;

/** The relationships of the beta `profile` resource (its only structural property is `id`). */
export const PROFILE_RELATIONSHIPS = [
  "accounts",
  "addresses",
  "anniversaries",
  "awards",
  "certifications",
  "educationalActivities",
  "emails",
  "interests",
  "languages",
  "names",
  "notes",
  "patents",
  "phones",
  "positions",
  "projects",
  "publications",
  "responsibilities",
  "skills",
  "webAccounts",
  "websites",
] as const;

// `user: get` documents `$select` only; unselected, Graph returns exactly the default set,
// which is also the project default, so no data is hidden by it.
const userContract = (): QueryContract =>
  defineQueryContract({
    select: {
      allowed: USER_PROPERTIES,
      default: USER_DEFAULT_PROPERTIES,
      requiredInBatch: false,
    },
    projectEntity: {},
  });

// `user: get mailboxSettings` says only that "some" OData options are supported and shows
// none, so no query option is offered and none is sent (the whole object is compact).
// The `/me/mailboxSettings/<setting>` sub-paths are separate Graph paths and not in the catalog.
const mailboxSettingsContract = (): QueryContract =>
  defineQueryContract({
    projectEntity: { keep: MAILBOX_SETTINGS_PROPERTIES },
  });

// `profile: get` (beta): a plain GET and `$select` are documented; the relationships are
// returned only when expanded (page example 2). `$select` is not offered (`id` is the only
// structural property); `$expand` is a typed list of the documented relationship names with
// no nested options. Beta-only, so it has no `batchPolicy` consumer (batch is v1.0).
const profileContract = (): QueryContract =>
  defineQueryContract({
    expand: { allowed: PROFILE_RELATIONSHIPS },
    projectEntity: {},
  });

export function registerIdentityTools(
  handlers: Map<string, ToolImplementation>,
  client: GraphClient,
  catalog: EndpointCatalog,
): void {
  const entryOf = (toolName: string) => {
    const entry = catalog.find((e) => e.toolName === toolName);
    if (entry === undefined || entry.method !== "get") {
      throw new Error(`Catalog has no GET entry for ${toolName}.`);
    }
    return entry;
  };

  handlers.set(
    "get-current-user",
    getTool(client, "get-current-user", entryOf("get-current-user").pathPattern, userContract()),
  );
  handlers.set(
    "get-mailbox-settings",
    getTool(
      client,
      "get-mailbox-settings",
      entryOf("get-mailbox-settings").pathPattern,
      mailboxSettingsContract(),
    ),
  );

  // Documented by Microsoft in the BETA API only (`profile: get`; no v1.0 page). The catalog
  // entry's `apiVersion: "beta"` is the only thing that selects `/beta/`, so refuse to
  // register against an entry that is not beta rather than silently serving v1.0.
  const profile = entryOf("get-my-profile");
  if (profile["apiVersion"] !== "beta") {
    throw new Error(
      "get-my-profile is a beta-only endpoint (profile: get is documented in beta only).",
    );
  }
  const contract = profileContract();
  const tool = getTool(client, "get-my-profile", profile.pathPattern, contract);
  // Deliberately no batchPolicy: graph-batch is a v1.0 batch and never targets a beta entry.
  handlers.set("get-my-profile", { inputSchema: contract.shape, handler: tool.handler });
}
