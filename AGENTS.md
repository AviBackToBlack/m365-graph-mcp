# Contributor and coding-agent guide

This file documents **public product engineering constraints**, not a private agent orchestration protocol.

## Invariants

- Preserve the retrieval-only boundary: **no mutations, arbitrary Graph proxy, arbitrary URL/HTTP-method inputs, or widening of the fixed Microsoft Graph origin**.
- `endpoints.json` defines the exact allowlisted tool surface. The registered handlers, method counts, `TOOL_SURFACE.md` and schemas must stay in sync. Catalog changes require deliberate review and tests.
- Treat Microsoft 365 content and other MCP tool inputs as untrusted. Do not permit embedded content to influence execution, routing, hostnames, headers, or process execution.
- Authentication mode is explicit: `M365_GRAPH_TOKEN_SOURCE=manual|bridge`; no auto-detection or fallback. Keep the bridge IPC user-restricted and Unix-domain-only on macOS.
- Never output credentials, decoded JWT claims, raw access tokens, browser profile contents, tenant/client identifiers, or private Microsoft 365 data to logs or tests.
- Do not add hidden network requests, telemetry, dynamic evaluation, unrestricted shell/process execution, or unbounded result collection.
- Preserve token security and browser-profile retirement guidance documented in `ARCHITECTURE.md`, `THREAT_MODEL.md`, and `SECURITY_INVARIANTS.md`.

## Contribution workflow

Make bounded changes on a feature branch. Do not commit directly to the default branch. Run `npm ci`, `npm run check`, and `npm run build`; add negative and regression coverage for changed invariants. The package-release workflow additionally checks the minimal npm tarball, its exact file manifest, and install smoke tests.

Changes must pass relevant CI, dependency, secret and project-specific static-analysis gates. Pin GitHub Actions by complete SHA and containers by immutable digest. Do not commit credentials or machine-specific paths. Open a pull request and allow its normal review process; reviews and merge approval are separate decisions.

For prior art, consult the official Microsoft Graph documentation first and inspect external implementations only under their applicable licences. Preserve required attribution where code is reused; do not import third-party rules or snippets whose redistribution terms are incompatible with this MIT-licensed repository.

## Runtime

The supported production design targets macOS and Claude Desktop; keep production code portable and do not introduce mandatory Windows development utilities. The local development environment is not a production dependency.
