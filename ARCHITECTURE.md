# m365-graph-mcp — Architecture

## Purpose

`m365-graph-mcp` is a local, retrieval-only MCP server exposing a deliberately fixed subset of Microsoft Graph to Claude Desktop.

It is not a general Graph proxy and not a general Microsoft 365 automation server.

## Development and runtime environments

- Development and CI may use Windows or Linux; neither is required at runtime.
- The production/runtime target is macOS with Claude Desktop.
- Runtime code must therefore remain portable to macOS even when all active development happens on Windows.
- Windows-specific development helpers are acceptable only when isolated from production/runtime code.
- Runtime paths, process handling, shell behavior, browser integration, and packaging must use cross-platform or macOS-correct abstractions rather than assuming Windows drive letters, separators, registry state, PowerShell, or Windows browser-profile locations.
- Final release acceptance is performed on a clean macOS machine with Claude Desktop.

## Runtime shape

```text
Claude Desktop (macOS)
      |
      | stdio MCP
      v
m365-graph-mcp
      |
      | short-lived delegated Graph access token
      v
TokenProvider
      |
      +-- bootstrap: manually supplied access token
      |
      `-- later: Graph Explorer browser token bridge
      |
      v
https://graph.microsoft.com
```

## Technology baseline

- TypeScript on current supported Node.js LTS.
- stdio MCP transport only.
- Native `fetch` for Graph HTTP unless an SDK is proven necessary.
- No HTTP MCP listener.
- No OAuth implementation in the MCP server.
- No client secret, refresh token, PRT, browser cookie, or MSAL cache owned by the MCP server.
- Exact dependencies and lockfile are pinned at bootstrap; no runtime package resolution.

## Source-of-truth hierarchy

1. `endpoints.json`: normative list of capabilities/tools.
2. `TOOL_SURFACE.md`: generated human-readable representation.
3. Project contracts such as `SECURITY_INVARIANTS.md` and this architecture: normative runtime/security design.
4. Current official Microsoft documentation: normative endpoint semantics and permissions.
5. `SOFTERIA_REFERENCE_POLICY.md`: mandatory prior-art reconnaissance policy.
6. Softeria `ms-365-mcp-server`: mandatory engineering reference for relevant prior art and regressions, but not a runtime dependency and not normative.
7. Tests: executable enforcement of the architecture/security invariants.

Every roadmap unit must complete the Softeria reconnaissance gate before implementation. The gate is exhaustive within the scope of the current unit, uses the pinned baseline plus a recorded current-upstream SHA, and requires ADOPT / ADAPT / REJECT decisions for material prior-art approaches. See `SOFTERIA_REFERENCE_POLICY.md`.

## Core components

### MCP server

Registers only tools present in `endpoints.json`.

No generic `graph_request`, arbitrary URL, arbitrary path, or arbitrary HTTP method tool may exist.

### GraphClient

Responsibilities:
- fixed origin `https://graph.microsoft.com`;
- version routing (`v1.0` by default, explicit `beta` only where the catalog requires it);
- bearer injection;
- request timeout;
- bounded response body size;
- Graph error normalization;
- 429/Retry-After handling;
- bounded pagination;
- safe `@odata.nextLink` handling by validating/rebuilding against the fixed Graph origin;
- no redirects to non-Graph hosts.

### TokenProvider

Small interface, conceptually:

```ts
interface TokenProvider {
  getAccessToken(): Promise<string>;
}
```

The Graph client does not know how authentication is performed.

Manual provider (default, `M365_GRAPH_TOKEN_SOURCE` unset or `manual`): manually supplied short-lived bearer, preferably from an environment variable or OS-protected local secret handoff.

Bridge provider (`M365_GRAPH_TOKEN_SOURCE=bridge`, SM-15, D45): Graph Explorer token bridge reached over a local-only, user-restricted IPC broker (Unix domain socket on macOS; never TCP).

### Tool implementations

Prefer thin domain modules:
- mail
- calendar
- identity/org
- Teams chats
- Teams channels
- groups
- SharePoint/search
- delta/incremental retrieval

Tool handlers should normalize Graph responses and avoid returning unnecessary fields or giant bodies by default.

## Read-only semantics

Read-only is a capability property, not a CLI flag.

Most tools are GET.

The only allowed POSTs are the five entries explicitly present in `endpoints.json`:
- Graph batch (GET subrequests only)
- calendar getSchedule
- findMeetingTimes
- Microsoft Search query
- Copilot Retrieval

No mutation endpoint is implemented anywhere.

## Graph Explorer bridge (SM-15)

The production bridge runs on macOS alongside Claude Desktop. Development/prototyping may occur on Windows, but implementation must not depend on Windows-only browser storage locations or process semantics.

The browser remains the owner of:
- corporate SSO;
- MFA;
- Conditional Access;
- MSAL session and token renewal.

The bridge exposes only a current short-lived Graph access token over a local-only, user-restricted IPC broker (D43): a Unix domain socket on macOS in a bridge-owned `0700` directory (socket `0600`), never TCP. The broker offers only `getToken` and a token-free `status`; filesystem permissions are the boundary (a per-start secret would not protect against same-UID processes and is not used). `BridgeTokenProvider` is the MCP-side `TokenProvider` that talks to it and maps every non-ok outcome to `reauth_required`.

Preferred design:
- dedicated persistent Chrome/Edge profile on the production macOS host, launched and owned by the bridge;
- Graph Explorer authenticated interactively by the user;
- dependency-free direct CDP over `--remote-debugging-pipe` (D42; fd 3/4 on POSIX; never a debugging port, WebSocket or TCP), with a fixed method allowlist, observes an outbound request to the exact `https://graph.microsoft.com` origin and extracts only the short-lived bearer;
- the candidate passes the local token acceptance sanity gate (D44) before it is served: bridge-owned source, exact origin, non-401 response, decodable JWT, Graph audience, pinned `tid`, delegated `scp` only, at least 5 minutes remaining, optional client pin. This is not cryptographic validation;
- opaque/undecodable tokens return `unsupported_token_format` and are never served;
- do not scrape refresh tokens, cookies, PRTs, client secrets, or encrypted MSAL cache.

If interaction is required, return a clear re-auth-required state rather than attempting to bypass policy.

### Bridge components (SM-15 parts A and B)

- `src/bridge/cdp-pipe.ts`: dependency-free CDP client. NUL-terminated JSON, written to the browser's fd 3 and read from fd 4. Fixed command allowlist (`Target.setDiscoverTargets`, `Target.getTargets`, `Target.attachToTarget` (flatten), `Target.detachFromTarget`, `Network.enable` (session-scoped), `Browser.close` for clean shutdown) enforced before anything is written; fixed consumed-event set; 8 MiB frame cap (oversize closes the pipe); per-command timeout; EOF/error rejects pending commands.
- `src/bridge/browser.ts`: browser discovery (`M365_GRAPH_BROWSER_PATH`, else fixed macOS Chrome/Edge paths) and launch (`shell: false`, stdio `ignore, ignore, pipe, pipe, pipe`, dedicated `--user-data-dir` under the `0700` bridge directory, `--remote-debugging-pipe`, never a debugging port). stderr is drained and discarded unread. A live `SingletonLock` or an immediate browser exit is `profile_busy`; locks are never deleted.
- `src/bridge/capture.ts`: the `TokenSource`. Attaches only to the page target whose URL is exactly the Graph Explorer page (https, `developer.microsoft.com`, path `/graph/graph-explorer` or `/<locale>/graph/graph-explorer` where `<locale>` is exactly `xx-xx` (two ASCII letters, hyphen, two ASCII letters, case-insensitive), optional trailing slash; no extra segments, no suffix matching), enables Network on that session only, extracts only a single `Authorization: Bearer` value from requests to the exact Graph origin, correlates the response status (`requestId`, bounded pending map with a TTL), then runs the D44 gate. A token replaces the held one only if it differs and expires no earlier than the held one (equal expiry counts as a renewal), or the held one is no longer servable; a 401 for the held token clears it.
  - Claim lifecycle: if the claimed page's `targetInfoChanged` reports a non-Graph-Explorer URL, the claim is released (`Target.detachFromTarget`, session and pending entries dropped) and `Target.getTargets` is rescanned to attach another eligible page; a rescan also follows `targetDestroyed`/`detachedFromTarget`. A still-valid held token is kept while unclaimed. If the page returns to Graph Explorer it is eligible again. While attaching, the latest URL is re-checked before the session is committed; a page that navigated away is detached, not claimed.
  - ExtraInfo ordering: `requestWillBeSentExtraInfo` may precede `requestWillBeSent`. Only the extracted bearer candidate (never the header map) is buffered by `requestId` (bounded, same TTL; dropped on `loadingFailed`) and merged when the main event arrives for an exact-Graph-origin URL. Redirects: a `requestWillBeSent` carrying `redirectResponse` resets everything known for that `requestId`; ExtraInfo carries no URL or hop marker and cannot be tied to the final hop, so for a redirected request only the bearer in that final hop's own `request.headers` counts (tradeoff: a redirected request whose Authorization appears only in ExtraInfo is not captured).
- `src/bridge/run.ts` / `src/bridge/main.ts` (bin `m365-graph-bridge`): config, directories, broker start (an already-running bridge is detected before another browser is launched), browser launch, capture, SIGINT/SIGTERM shutdown. Exit codes: 0 clean, 2 config, 3 browser, 4 profile busy, 5 broker, 6 browser exited, 7 CDP setup, 8 CDP pipe closed while running (browser closed/killed, EOF, oversize frame; the browser is stopped and the broker removed). An asynchronous browser `error` event after spawn is exit 3 (browser), not 4. SIGINT/SIGTERM handlers are installed before startup; a signal during startup tears down whatever was created (broker/socket, browser) and exits 0.
- Dev-only: `M365_GRAPH_BRIDGE_DEV_PIPE=\\.\pipe\<name>` (honored only on Windows) lets the dev loop run the broker end-to-end with the fake CDP peer. It does not change the production path.

### SM-15 acceptance diagnostics (D48)

An opt-in, shape-only observation mode makes otherwise unobservable acceptance properties attestable on macOS (checklist items 1, 2 exact URL, 4 ExtraInfo, 4a, 4b) without widening D44 logging in general. It is acceptance observability, not a debug logger.

- Enablement: `M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS`, parsed in `src/bridge/config.ts` without trimming: unset or exactly empty is off, exactly `1` is on, any other value is a `BridgeConfigError` (bridge exit 2; the message names the variable, never the value). It is never auto-enabled and is neither read by the MCP server (`selectTokenProvider`) nor exposed through the broker or any MCP tool.
- Abstraction boundary: `src/bridge/acceptance-diagnostics.ts` is the only place that formats these lines. Its fixed set of typed methods take only string-literal unions, booleans and one bounded integer (remaining lifetime in minutes). It accepts no free string, maps every value through a closed lookup table (anything outside is dropped), imports nothing (no file system, network or environment access), and its sink is the bridge's existing stderr logger (`log` in `run.ts`, prefixed `acceptance: `). A disabled instance is a frozen no-op, and call sites gate classification work on `enabled`.
- Call sites classify locally, where the data already is, and hand over only the category: `classifyTokenShape` runs on the decoded claims on the final accepting return of `evaluateCandidate` (rejections keep their ordinary D44 reason categories and get no shape); `classifyGraphOrigin` and `classifyGraphExplorerPage` are the same matchers that decide acceptance (`isExactGraphOrigin` and `isGraphExplorerUrl` are derived from them, so there is no second matcher); `capture.ts` records where the bearer was first seen (main event, ExtraInfo before or after the main event) and redirect observations. Diagnostics never alter capture or gate decisions; with the mode off, behavior and stderr output are byte-identical to a build without it.
- Output vocabulary is closed (see `docs/macos-sm15-acceptance.md` for the per-item mapping): `mode=enabled`, `token_shape ...`, `graph_request_exact_origin_match=...`, `graph_explorer_page_shape=canonical|localized [locale_segment_shape=xx-xx]`, `authorization_observation=...`, `extra_info_authorization=present|not_observed|present_after_response`, `redirect_seen=yes`, `redirect_final_hop_authorization=main_event`, `redirect_extra_info_discarded=yes`.

### MCP token source selection (SM-15 part C, D45/D46)

`src/auth/select-token-provider.ts` (called from `src/index.ts`) picks the `TokenProvider` only from `M365_GRAPH_TOKEN_SOURCE`:
- unset or exactly `manual`: `ManualTokenProvider` (unchanged; reads `M365_GRAPH_ACCESS_TOKEN` and the optional `M365_TENANT_ID` pin);
- exactly `bridge`: `BridgeTokenProvider` on the endpoint from `resolveBridgeEndpoint` in `src/bridge/config.ts`, the one implementation `parseBridgeConfig` also uses (`M365_GRAPH_BRIDGE_DIR`, else the macOS default; Windows-only `M365_GRAPH_BRIDGE_DEV_PIPE`; never TCP);
- any other value, including empty, whitespace-only, case and padding variants, is a `TokenError` `auth_misconfigured`, so `main()` fails and the server does not start. Messages name the variable and never echo its value.

Rules: never auto-detected (not from the socket or directory, browser state, a present manual token or the OS); bridge mode never falls back to a manual bearer and manual mode never reads bridge configuration or contacts the bridge. Startup failures: bad selector, invalid endpoint configuration (including the dev pipe off Windows), and invalid manual configuration. A bridge that is not running is not a startup failure; token requests return `reauth_required`. Bridge mode neither requires nor reads `M365_GRAPH_TENANT_ID` and does no claim parsing: tenant binding is the bridge's D44 gate. `M365_GRAPH_TENANT_ID` is bridge-process config, `M365_TENANT_ID` is manual-mode config, with no aliasing or conflict logic between them.

Renewal limitation: the bridge only observes. A renewed token is captured only when Graph Explorer itself issues a Graph request (a user action or Graph Explorer's own calls). When the held token drops below the 5-minute serve floor the broker reports `reauth_required` until a fresh acceptable token is observed. Opaque tokens yield `unsupported_token_format`.

Operator run instructions (macOS): export `M365_GRAPH_TENANT_ID` (required), optionally `M365_GRAPH_BROWSER_PATH`, `M365_GRAPH_BRIDGE_DIR`, and the client pin variables; run `m365-graph-bridge`; sign in to Graph Explorer in the browser window it opens and run any query so a Graph request is observed; keep the bridge running while using the MCP server, and start the MCP server with `M365_GRAPH_TOKEN_SOURCE=bridge` (and the same `M365_GRAPH_BRIDGE_DIR`, if you set one for the bridge). The live acceptance checklist is in `PERMISSIONS.md`.

### Bridge operator security contract (D52)

These are accepted residual risks (`THREAT_MODEL.md` R1-R4); the operator is responsible for the practices below. No code enforces them.

- **The bearer is not retrieval-only (R1).** Retrieval-only is a property of this MCP implementation, not of the bearer. The bridge serves whatever delegated token Graph Explorer holds (D44 does not inspect scopes), so any other component that obtains the bearer is not restricted by the MCP catalog.
  - Use the dedicated Graph Explorer bridge profile only for this project.
  - Do not intentionally consent mutation permissions in that profile.
  - Before production use, review the permissions consented to Graph Explorer and remove any this project does not need.
  - Do not use that profile for unrelated mutation experiments.
- **Bearer lifetime is not revocation (R2).** The observed bearer was long-lived (about 24 hours, D49), and CAE-capable access tokens may have long lifetimes. The bridge serves a token until the 5-minute-before-`exp` floor unless it is renewed or rejected externally; there is no local maximum token age. Revocation propagation is controlled by Microsoft and resource-side behavior, and the bridge has no stronger local revocation guarantee.
- **The bridge profile is a credential (R3).**
  - FileVault is required on the production Mac.
  - The profile is dedicated to Graph Explorer/bridge authentication: no browser extensions and no general browsing in it.
  - Treat the profile as a credential-bearing asset even when the bridge is stopped.
  - To retire it: (1) sign out of the corporate/Graph Explorer session; (2) close the bridge-owned browser; (3) stop the bridge; (4) delete `<bridge dir>/profile`.
- **No MCP-side invalidation (R4).** A Graph 401 during a tool call returns `reauth_required`, but the broker keeps offering the held token (D43: only `getToken` and `status`) until Graph Explorer itself receives a 401 for it, a newer token is captured, or it reaches the serve floor. Repeated `reauth_required` results in that window are expected.
