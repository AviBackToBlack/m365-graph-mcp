# Security policy

## Reporting a vulnerability

Please report suspected vulnerabilities **privately**. After this repository is published, prefer GitHub's **Private vulnerability reporting** feature (Security tab → Report a vulnerability) when enabled.

If private vulnerability reporting is not yet enabled, request a private reporting channel from the maintainer **without including** any tokens, secrets, tenant identifiers, personal data, exploit payloads or corporate Microsoft 365 content. Do not file a public issue containing sensitive evidence.

A report should identify the affected source version or release, a short description of the trust boundary, safe reproduction steps using synthetic data, and the observed versus expected behaviour.

## Scope

Relevant areas include:

- The fixed Microsoft Graph origin and allowlisted tool surface;
- Token source selection, JWT-compatible acceptance checks and the user-restricted local IPC bridge;
- Browser profile, CDP pipe and Graph bearer handling;
- Graph query/path/header validation, batch containment and untrusted output processing;
- Build, npm release artifact, dependency security and supply-chain provenance.

This project is retrieval-only **by implementation**, not by token authority. Treat any Graph bearer and bridge browser profile as credentials, even if the server itself cannot mutate data. Read SECURITY_INVARIANTS.md and THREAT_MODEL.md for the detailed contract.

No response-time, maintenance-support period, or bounty commitment is promised by this document.
