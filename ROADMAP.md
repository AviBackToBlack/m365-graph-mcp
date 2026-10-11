# Product roadmap

This roadmap describes public product capability and remaining verification work, rather than the historical engineering sequence.

## Implemented, with repository test coverage

- A fixed 76-tool retrieval-only Microsoft Graph catalog (71 GET, five retrieval-only POST), backed by exact surface tests.
- Typed and bounded Graph requests, pagination, batch subrequests, error handling, and Microsoft 365 content normalization.
- Manual delegated bearer mode and an explicit bridge mode with local IPC.
- Graph Explorer bridge using a dedicated browser profile and CDP pipe on macOS.
- Project-specific Semgrep rules, gitleaks history scanning, npm audit, OSV-Scanner, Dependabot, CycloneDX SBOM, SHA-256 artifact hashes, and an exact minimal npm package manifest.

## Security and release work

- Enable and verify native GitHub CodeQL/code scanning, secret scanning, Dependency Review and artifact attestations when the repository's publication settings support them. The independent scanners above are defense-in-depth, **not equivalents**.
- Verify all security workflows against the clean repository and review any findings.
- Document and test a reproducible release process using a precise source identity, minimal npm tarball, SBOM, SHA-256 hashes and provenance.
- Perform final clean-machine macOS/Claude Desktop acceptance against the release distribution, including safe bridge operation and explicit permission limitations.

## Non-goals

- A generic Microsoft Graph proxy or arbitrary Graph API calls.
- Mutation operations, automatic escalation of scopes, hidden bearer capture, cloud relay or TCP token broker.
- Silent publication to the npm registry.
