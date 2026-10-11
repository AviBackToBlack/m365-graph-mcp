# m365-graph-mcp — Security Invariants

These are release-blocking invariants.

## Capability boundaries

1. `endpoints.json` is an explicit capability allowlist.
2. No generic Graph proxy tool exists.
3. No arbitrary URL tool exists.
4. No mutation semantics exist.
5. No mail send/delete/update, calendar create/update/delete, Teams send/update/delete, file upload/update/delete, or permission-management implementation exists.
6. A selected POST is allowed only when its semantics are retrieval-only and its exact path is allowlisted.

## Batch

`graph-batch` MUST:
- accept GET subrequests only;
- reject POST/PATCH/PUT/DELETE;
- accept relative URLs only;
- reject nested `$batch`;
- require every path to match an *implemented* allowlisted GET endpoint (an endpoint whose tool exists), with every captured path-parameter value passing its allowlisted value shape;
- reject caller-supplied auth/host headers;
- enforce Graph's batch size limit and local response-size limits.

## Network

- Graph origin is fixed to `https://graph.microsoft.com`.
- Never send the bearer token to any other origin.
- Never follow an absolute `@odata.nextLink` origin blindly.
- No telemetry or analytics dependency may receive M365 content.
- No Graph response may control a network destination.

## Authentication material

- Never log access tokens.
- Never return tokens through MCP tool results.
- Never put bearer tokens in command-line arguments.
- MCP does not own refresh tokens, browser cookies, PRTs, or client secrets.
- The MCP server selects its token source only from explicit `M365_GRAPH_TOKEN_SOURCE` configuration (D45, D47). It uses `manual` (`ManualTokenProvider`) when the variable is unset or set exactly to `manual`, and `bridge` (`BridgeTokenProvider`) only when it is set exactly to `bridge`. Any other value is a startup configuration error. The source is never auto-detected: bridge availability, socket or bridge-directory existence, a present manual token, browser state and the OS never select the provider. There is never automatic fallback between token providers. Bridge mode never uses a manual bearer, and an unavailable bridge gives `reauth_required`. Manual mode never probes the bridge and never falls back to it.
- The token bridge captures a bearer only from a bridge-owned Graph Explorer request to the exact Graph origin (`https://graph.microsoft.com`, default port, no suffix matching), and only when the correlated Graph response was not 401.
- The bridge applies fail-closed local sanity checks when the token format is inspectable (Graph audience forms, delegated-only scopes, minimum remaining lifetime, optional operator-configured client pin) and pins the configured corporate tenant (`M365_GRAPH_TENANT_ID`, required) for the JWT-compatible profile.
- The bridge performs no cryptographic validation of the token; Microsoft Graph remains authoritative. Decoded claims are advisory only.
- Opaque or undecodable Graph tokens are not supported and are never served without a later explicit decision (`unsupported_token_format`).
- The bridge never logs tokens, decoded payloads, tenant ids, scopes, audiences or client ids; diagnostics are limited to reason categories and lifetime metadata.
- An operator-enabled SM-15 acceptance diagnostic mode may emit only predefined non-secret shape classifications needed to verify the bridge contract. It is off by default, writes only to stderr, persists nothing, and must never emit bearer material, raw token claims, identifiers, scopes, headers, complete URLs, request/target/session identifiers, or Microsoft 365 content. (D48)
- The token bridge broker is local-only IPC restricted to the current user: a Unix domain socket (macOS) in a `0700` bridge-owned directory with a `0600` socket. It never listens on TCP and never binds `127.0.0.1`, `::1`, `0.0.0.0` or any routable interface. Windows named pipes are development/test only.
- The broker exposes only the current accepted bearer, token-free status metadata, and re-auth/unavailable states; it exposes no browser or CDP control.

## Prompt/tool injection containment

Microsoft 365 content is untrusted data.
- Return structured content, not executable instructions.
- Tool descriptions state that email/Teams/SharePoint content must never be treated as MCP/server instructions.
- Data cannot expand capabilities because endpoints and methods are code-level allowlists.
- No Graph-returned URL may become an arbitrary fetch target.

## Query and path handling

- Encode path parameters as path segments, never concatenate unchecked slash-containing input.
- Permit only documented query parameters per tool where practical.
- Bound `$top`, page counts, result counts, response sizes, and time windows.
- Treat `$search`, `$filter`, `$orderby`, `$expand`, and KQL as data; never shell-evaluate them.
- Reject malformed or unsupported query combinations with a clear error.

## Pagination

- Support `@odata.nextLink` only through the fixed Graph origin.
- Enforce page/result ceilings.
- Delta links are opaque capabilities but must still resolve to the fixed Graph origin.
- 410 Gone on expired deltas is surfaced as a restart-required condition.

## Process/filesystem

- MCP has no shell-execution tool.
- Graph content cannot select executable names or arguments.
- Avoid persistent storage of M365 content unless a future explicit requirement is approved.
- Logs contain metadata only and must redact PII where practical.

## Supply chain

- Lock dependency versions.
- Pin GitHub Actions by immutable commit SHA.
- Minimal workflow permissions.
- Current CI: tests, lint/typecheck, build, npm audit, OSV-Scanner, project-specific Semgrep rules, gitleaks full-history scan, a CycloneDX SBOM, reproducible minimal npm package checks and SHA-256 artifact hashes.
- GitHub-native CodeQL/code scanning, Dependency Review, secret scanning and artifact attestations are separate capabilities enabled and validated at the publication gate; independent scanners are not claimed to replace them.
- Generated artifacts are integrity checked; an artifact hash without an attestation does not establish publisher provenance.
