# Prior art and external attribution

The capability metadata is derived from the public `Softeria/ms-365-mcp-server` endpoint catalog, and the project also uses official Microsoft Graph documentation.

The audited prior-art reference was `Softeria/ms-365-mcp-server` at commit `3683f55f3917885d80f31cb6f5423ef6da4cee52`. This reference is **not a runtime dependency** or a reason to inherit broader capabilities. The full Softeria MIT notice and precise scope of reuse are included in `THIRD_PARTY_NOTICES.md`, including source metadata and some retained descriptive text.

Our normative source of truth for available capabilities is `endpoints.json`; security and architecture contracts are in `SECURITY_INVARIANTS.md` and `ARCHITECTURE.md`. Microsoft documentation governs Graph endpoint semantics and permissions.

When implementing or reviewing a capability, consult relevant prior art and its issue/test history to identify known edge cases. Reuse is not automatic: only adopt compatible code after checking the licence, notices and provenance, and include attribution where required. Document a decision to adopt, adapt, or independently implement material behavior.

Do **not** expand the product to include mutation tools, unrestricted Graph calls, additional authentication modes or third-party telemetry because those exist upstream. No upstream Semgrep rules are included: the project's own security rules are authored and maintained within this repository.
