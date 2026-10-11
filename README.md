# m365-graph-mcp

A deliberately bounded, retrieval-only [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server for selected Microsoft 365 data through Microsoft Graph.

The tool catalog contains **76 tools** (**71 GET**, **5 retrieval-only POST**). The exact catalog is in [`endpoints.json`](./endpoints.json) and is checked by [`TOOL_SURFACE.md`](./TOOL_SURFACE.md) and the repository tests. This is **not** a generic Microsoft Graph proxy. There are no mutation tools.

## Security and trust boundary

- MCP runs over local **stdio**, with a fixed `https://graph.microsoft.com` Graph origin and strictly typed tool input.
- Microsoft 365 content is untrusted data; do not execute or obey instructions embedded in retrieved content.
- **Read-only describes this server, not your Microsoft Graph access token.** A captured delegated token might carry write permissions. Keep tokens and the browser profile used to obtain them secure.
- Tool results pass through the MCP client to its configured model provider. Consider your organisation's privacy and data-handling rules before connecting confidential accounts.
- No access tokens, refresh tokens, browser cookies, PRTs, or session files should be committed, logged, pasted into issues, or sent to diagnostics.
- The token bridge's local Unix socket protects against other OS users, **not** against malicious programs running as the same account.
- The bridge profile is credential-bearing even when the bridge stops. On macOS use FileVault, a dedicated browser profile, no extensions, and no unrelated browsing.
- Access-token lifetime is **not** a revocation guarantee; observed Graph Explorer tokens can live for approximately a day.

Read [`SECURITY_INVARIANTS.md`](./SECURITY_INVARIANTS.md), [`THREAT_MODEL.md`](./THREAT_MODEL.md), and [`ARCHITECTURE.md`](./ARCHITECTURE.md) before production use.

## Platform and build

The runtime target is **macOS with Claude Desktop**. Source development and CI also use other platforms. Node.js **22.12 or newer** is required. A clean macOS installation remains part of the release acceptance process; a successful CI build alone is not a deployment certification.

From a source checkout:

```sh
npm ci
npm run check
npm run build:release
```

For a distribution artifact use `npm pack`; the repository's CI verifies the exact package contents, builds a CycloneDX SBOM and records SHA-256 hashes. Obtain release artifacts from an approved distribution channel and verify the corresponding integrity evidence. A source checkout is not the same as the minimal runtime package.

The package provides two Node executables:

- `m365-graph-mcp` — stdio MCP server (compiled `dist/index.js`).
- `m365-graph-bridge` — local Graph Explorer token bridge (compiled `dist/bridge/main.js`).

Configure the MCP client to launch the server with its process environment rather than storing tokens in a shared shell profile.

## Token source selection

The MCP server accepts exactly two choices via `M365_GRAPH_TOKEN_SOURCE`:

| Setting | Behaviour |
| --- | --- |
| Unset or exactly `manual` | Read a delegated bearer from `M365_GRAPH_ACCESS_TOKEN`; optionally enforce `M365_TENANT_ID` |
| Exactly `bridge` | Ask the local bridge over its user-restricted IPC endpoint; **never** fall back to a manual token |

All other values, including the empty string, are configuration errors. There is no automatic source detection or fallback.

In manual mode, `M365_GRAPH_ACCESS_TOKEN` must contain a valid, appropriately authorised delegated access token. Protect it as a credential. In bridge mode start `m365-graph-bridge` separately with the required `M365_GRAPH_TENANT_ID` pin. That variable configures the **bridge process**; `M365_TENANT_ID` applies only to **manual server mode**. The two names are intentionally not aliases.

Bridge mode uses a Unix-domain socket on macOS, under `M365_GRAPH_BRIDGE_DIR` or the macOS default `~/Library/Application Support/m365-bridge`. It does not use TCP. If the bridge is stopped or needs reauthentication, tool calls return a typed `reauth_required` error.

The bridge opens its own browser profile for Graph Explorer login/MFA and observes Graph requests via Chrome DevTools Protocol **pipe**, not a remote-debugging TCP port. It has a deliberately narrow token-acceptance policy; it cannot provide stronger guarantees about Microsoft-side revocation.

See [`docs/macos-sm15-acceptance.md`](./docs/macos-sm15-acceptance.md) for the non-secret macOS verification workflow.

## Permissions, capabilities, and limitations

[`PERMISSIONS.md`](./PERMISSIONS.md) lists the catalog's declared delegated scope hints. Actual availability depends on Microsoft Graph endpoint capabilities, your delegated consent and your tenant's access policy. Some selected endpoints are unavailable under certain permission sets. A Graph 401 or 403 is surfaced as a typed error; it is not a reason to broaden scopes automatically.

Read [`TOOL_SURFACE.md`](./TOOL_SURFACE.md) for the exact available operations. `graph-batch` permits only bounded GET subrequests that match implemented allowlisted endpoints; no arbitrary origins, bodies, dependency chains, or mutation methods.

## Contributing and security reports

Changes to the tool catalog, security boundaries, token handling, or supply chain require explicit review and regression coverage. Follow [`AGENTS.md`](./AGENTS.md) for contributor and coding-agent invariants. Report suspected security issues using the private channels in [SECURITY.md](./SECURITY.md), not public issues. See [CONTRIBUTING.md](./CONTRIBUTING.md) before proposing changes.

This source repository is distinct from the minimal npm distribution artifact. The `"private": true` npm setting intentionally blocks accidental npm registry publication; it does **not** describe GitHub repository visibility.

## Engineering traceability

Source comments, tests and contract documents may retain historical D# (design decision) and SM-# (development milestone) identifiers. They are engineering provenance labels, not public support cases, product features or links to private planning records. This release does not depend on access to internal decision or operator ledgers.

## Acknowledgements

The fixed capability catalog in endpoints.json is derived from the endpoint metadata of Softeria's ms-365-mcp-server (https://github.com/Softeria/ms-365-mcp-server), under its MIT licence (Copyright (c) 2025 Softeria). It selects Softeria entries, keeps a subset of their schema and retains some descriptive text. The runtime server, Graph client, handlers, token bridge and security controls are independently implemented here. The full upstream licence text and scope of derived material are recorded in THIRD_PARTY_NOTICES.md.

The minimal npm package ships both LICENSE and THIRD_PARTY_NOTICES.md.

## License

MIT — see [`LICENSE`](./LICENSE).
