# macOS bridge acceptance and safe-operation runbook

This is a **non-secret verification procedure**, not a place to store live Graph Explorer data. The supported production design uses Claude Desktop on macOS, Node.js 22.12+, and a local `m365-graph-bridge`.

## Prepare a trusted build

Use a reviewed distribution from the canonical source repository. For a release tarball, verify its published SHA-256 and SBOM before installing. If no release is available, build from a trusted source checkout using `npm ci`, `npm run check`, and `npm run build:release`.

Do not copy a personal/corporate browser profile into the build directory. Use FileVault on the host.

## Configure the bridge

- Use a **dedicated browser profile**, without extensions or general browsing.
- Configure `M365_GRAPH_TENANT_ID` for the bridge. It is the required corporate tenant pin, not the server's optional manual-mode `M365_TENANT_ID`.
- If necessary, configure `M365_GRAPH_BROWSER_PATH` for the selected Chrome/Edge executable and `M365_GRAPH_BRIDGE_DIR` for a dedicated, user-owned directory.
- Start `m365-graph-bridge` from the verified build; complete Graph Explorer login/MFA **inside its browser**. Do not place tokens in shell scripts, terminal output, issue comments, or screenshots.
- Configure the MCP server with exactly `M365_GRAPH_TOKEN_SOURCE=bridge`. Never set a manual bearer as a fallback.

The broker must use a local Unix socket, not TCP. Expect owner-only directory/socket permissions and explicit `reauth_required` responses when the bridge is unavailable. A second bridge using an already locked profile must fail safely.

## Optional shape-only diagnostic verification

Set `M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS=1` **only** for the diagnostic run. Unset/empty disables it; any other non-empty value is an error. The mode writes fixed categories to stderr; it must not reveal raw bearer data, token claims, OAuth scopes, URL contents, request identifiers or private Graph responses.

Check that accepted tokens have a supported Graph audience category, delegated scope *presence*, an expected client-claim category, and a matching tenant category. Confirm target attachment, browser redirects and Graph request observations only by the documented **shape categories**, never by full URLs or raw headers.

A Graph 401 received by the server produces `reauth_required`; bridge-side invalidation depends on correlated browser traffic. Local test coverage of 401 behavior does not prove Microsoft-side revocation latency. Token ages may be much longer than a typical access-token refresh interval.

Stop the diagnostic run and relaunch without diagnostics for normal operation. Never persist stderr transcripts containing even category-only observations unless required by an approved privacy policy.

## Retire the browser profile

Sign out of the Graph Explorer corporate session, close the bridge-owned browser, stop the bridge, and remove the dedicated `<bridge dir>/profile` directory. Do not remove another browser's profile. The profile is a credential-bearing asset even when the bridge is stopped.
