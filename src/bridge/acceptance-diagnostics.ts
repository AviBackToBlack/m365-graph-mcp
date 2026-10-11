/**
 * SM-15 acceptance diagnostics (D48): an operator-enabled, shape-only observation channel used to
 * attest properties of the bridge contract that D44's logging limits otherwise hide.
 *
 * This is NOT a debug logger. The surface is a fixed set of typed events whose parameters are
 * closed string-literal unions, booleans and one bounded integer. There is deliberately no method
 * that accepts a free string, and nothing token-, CDP-, URL- or page-derived can reach the output:
 * callers classify locally (see `classifyTokenShape`, `classifyGraphOrigin`,
 * `classifyGraphExplorerPage`) and hand over only the category. As defense in depth, every value is
 * mapped through a closed lookup table here and anything outside the set (for example a value forced
 * in with a cast) is dropped instead of printed.
 *
 * The only sink is the injected `write` function, which the bridge wires to its existing stderr
 * logger. This module imports nothing: no file system, no network, no environment access, no state.
 * The mode is off by default (`M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS`, exactly `1`); a disabled
 * instance does no work at all.
 */

export type AudienceCategory = "graph_url" | "graph_url_trailing_slash" | "graph_application_id";
/** Which client claim the token carries; `both`/`none` keep the label truthful when pinning is off. */
export type ClientClaimKind = "appid" | "azp" | "both" | "none";
export type TokenVersionCategory = "v1" | "v2" | "other_or_absent";
export type Presence = "present" | "absent";
export type TenantCheck = "matched" | "mismatch";

/** Shape of an accepted JWT-compatible token. Categories only; never claim values. */
export interface TokenShape {
  audience: AudienceCategory;
  clientClaim: ClientClaimKind;
  version: TokenVersionCategory;
  scp: Presence;
  roles: Presence;
  tenant: TenantCheck;
  /** Coarse remaining lifetime in whole minutes; clamped to 0..MAX_REMAINING_MINUTES when printed. */
  remainingMinutes: number;
}

export type SchemeClass = "https" | "other";
export type HostClass = "expected_graph_host" | "other";
export type PortClass = "default_https" | "other";

/** Classification of an outbound request URL against the exact Graph origin. */
export interface GraphRequestOrigin {
  exactOriginMatch: boolean;
  scheme: SchemeClass;
  host: HostClass;
  port: PortClass;
}

export type GraphExplorerPageShape = "canonical" | "localized";

/** Where the bearer for an evaluated Graph request was first observed (checklist 4/4a). */
export type AuthorizationObservation =
  "main_event" | "extra_info_before_main" | "extra_info_after_main";

/**
 * Whether a non-redirected Graph request's ExtraInfo event carried a bearer before evaluation
 * (checklist item 4), independent of where the bearer was first observed. `not_observed` means no
 * such ExtraInfo bearer arrived before the request was evaluated; it does not prove absence.
 * `present_after_response` is a follow-up line for a request already reported: its ExtraInfo bearer
 * arrived only after the response was received and the request was evaluated.
 */
export type ExtraInfoAuthorization = "present" | "not_observed" | "present_after_response";

export interface AcceptanceDiagnostics {
  /** False for the shared no-op instance; callers gate classification work on this. */
  readonly enabled: boolean;
  /** Announces that the mode is on (fixed text). */
  modeEnabled(): void;
  tokenShape(shape: TokenShape): void;
  graphRequestOrigin(origin: GraphRequestOrigin): void;
  graphExplorerPage(shape: GraphExplorerPageShape): void;
  authorizationObservation(observation: AuthorizationObservation): void;
  extraInfoAuthorization(value: ExtraInfoAuthorization): void;
  /** A redirect hop was observed on a Graph request. */
  redirectSeen(): void;
  /** The final redirect hop's bearer came from its own requestWillBeSent headers. */
  redirectFinalHopAuthorizationFromMainEvent(): void;
  /** ExtraInfo for a redirected request was not used (it cannot be tied to the final hop). */
  redirectExtraInfoDiscarded(): void;
}

export const MAX_REMAINING_MINUTES = 99_999;

/** Closed sets. Output values are looked up here, never interpolated from caller input. */
const AUDIENCES: ReadonlySet<string> = new Set([
  "graph_url",
  "graph_url_trailing_slash",
  "graph_application_id",
]);
const CLIENT_CLAIMS: ReadonlySet<string> = new Set(["appid", "azp", "both", "none"]);
const VERSIONS: ReadonlySet<string> = new Set(["v1", "v2", "other_or_absent"]);
const PRESENCE: ReadonlySet<string> = new Set(["present", "absent"]);
const TENANT_CHECKS: ReadonlySet<string> = new Set(["matched", "mismatch"]);
const SCHEMES: ReadonlySet<string> = new Set(["https", "other"]);
const HOSTS: ReadonlySet<string> = new Set(["expected_graph_host", "other"]);
const PORTS: ReadonlySet<string> = new Set(["default_https", "other"]);
const PAGE_SHAPES: ReadonlySet<string> = new Set(["canonical", "localized"]);
const AUTH_OBSERVATIONS: ReadonlySet<string> = new Set([
  "main_event",
  "extra_info_before_main",
  "extra_info_after_main",
]);
const EXTRA_INFO_AUTH: ReadonlySet<string> = new Set([
  "present",
  "not_observed",
  "present_after_response",
]);

/** Fixed line prefix; the bridge logger adds its own `m365-graph-bridge: ` prefix. */
export const ACCEPTANCE_PREFIX = "acceptance: ";

function inSet(set: ReadonlySet<string>, value: unknown): value is string {
  return typeof value === "string" && set.has(value);
}

function boundedMinutes(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  return Math.min(MAX_REMAINING_MINUTES, Math.max(0, Math.floor(value)));
}

const NOOP = (): void => undefined;

/** The shared disabled instance: every method is a no-op and does no work. */
export const DISABLED_ACCEPTANCE_DIAGNOSTICS: AcceptanceDiagnostics = Object.freeze({
  enabled: false,
  modeEnabled: NOOP,
  tokenShape: NOOP,
  graphRequestOrigin: NOOP,
  graphExplorerPage: NOOP,
  authorizationObservation: NOOP,
  extraInfoAuthorization: NOOP,
  redirectSeen: NOOP,
  redirectFinalHopAuthorizationFromMainEvent: NOOP,
  redirectExtraInfoDiscarded: NOOP,
});

/**
 * Creates the enabled instance. `write` receives complete fixed-format lines (without the bridge's
 * own prefix) and must be the bridge's stderr logger; nothing else may be passed.
 */
export function createAcceptanceDiagnostics(write: (line: string) => void): AcceptanceDiagnostics {
  const emit = (text: string): void => write(`${ACCEPTANCE_PREFIX}${text}`);
  return Object.freeze({
    enabled: true,
    modeEnabled(): void {
      emit("mode=enabled");
    },
    tokenShape(shape: TokenShape): void {
      const minutes = boundedMinutes(shape?.remainingMinutes);
      if (
        !inSet(AUDIENCES, shape?.audience) ||
        !inSet(CLIENT_CLAIMS, shape.clientClaim) ||
        !inSet(VERSIONS, shape.version) ||
        !inSet(PRESENCE, shape.scp) ||
        !inSet(PRESENCE, shape.roles) ||
        !inSet(TENANT_CHECKS, shape.tenant) ||
        minutes === undefined
      ) {
        return;
      }
      emit(
        `token_shape audience=${shape.audience} client_claim=${shape.clientClaim} ` +
          `token_version=${shape.version} scp=${shape.scp} roles=${shape.roles} ` +
          `tenant=${shape.tenant} remaining_minutes=${minutes}`,
      );
    },
    graphRequestOrigin(origin: GraphRequestOrigin): void {
      if (
        typeof origin?.exactOriginMatch !== "boolean" ||
        !inSet(SCHEMES, origin.scheme) ||
        !inSet(HOSTS, origin.host) ||
        !inSet(PORTS, origin.port)
      ) {
        return;
      }
      emit(
        `graph_request_exact_origin_match=${origin.exactOriginMatch ? "yes" : "no"} ` +
          `graph_request_scheme=${origin.scheme} graph_request_host=${origin.host} ` +
          `graph_request_port=${origin.port}`,
      );
    },
    graphExplorerPage(shape: GraphExplorerPageShape): void {
      if (!inSet(PAGE_SHAPES, shape)) return;
      emit(
        shape === "localized"
          ? "graph_explorer_page_shape=localized locale_segment_shape=xx-xx"
          : "graph_explorer_page_shape=canonical",
      );
    },
    authorizationObservation(observation: AuthorizationObservation): void {
      if (!inSet(AUTH_OBSERVATIONS, observation)) return;
      emit(`authorization_observation=${observation}`);
    },
    extraInfoAuthorization(value: ExtraInfoAuthorization): void {
      if (!inSet(EXTRA_INFO_AUTH, value)) return;
      emit(`extra_info_authorization=${value}`);
    },
    redirectSeen(): void {
      emit("redirect_seen=yes");
    },
    redirectFinalHopAuthorizationFromMainEvent(): void {
      emit("redirect_final_hop_authorization=main_event");
    },
    redirectExtraInfoDiscarded(): void {
      emit("redirect_extra_info_discarded=yes");
    },
  });
}
