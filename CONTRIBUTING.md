# Contributing

Contributions are welcome when they preserve the intentionally narrow, retrieval-only Microsoft Graph capability boundary.

Before changing code, read AGENTS.md, ARCHITECTURE.md, SECURITY_INVARIANTS.md, TOOL_SURFACE.md and the relevant scoped part of PERMISSIONS.md. The exact catalog lives in endpoints.json.

Use a feature branch and a focused pull request. Keep generated/documented tool surfaces, permission metadata and production schemas consistent. Prefer typed inputs, bounded reads and targeted negative regression tests. Do not add arbitrary Graph URLs, mutation methods, hidden telemetry, refresh-token handling or browser-profile access outside the dedicated bridge.

Local validation requires Node.js 22.12+:

    npm ci
    npm run check
    npm run build

The release workflow separately validates the minimal packed runtime artifact, SBOM and hashes. Do not change package allowlists to include private operator notes, generated source maps, secrets or test artifacts.

Do not commit credentials, personal host paths, real Microsoft 365 data or live token material. Follow SECURITY.md to report vulnerabilities privately. Pull-request reviews and merge approval remain distinct gates.
