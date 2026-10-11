## Scope and rationale

<!-- Summarize what changed, why it is needed, and any linked issue. -->

## Security and contract impact

<!-- Describe security, token-handling, Microsoft Graph tool-surface, schema, or dependency impact; write "None" when not applicable. -->

- [ ] No credentials, real tenant data, personal host paths, or other private material added.
- [ ] The retrieval-only boundary and fixed Graph origin remain intact.
- [ ] If the tool catalog changed, handlers, schemas, endpoints.json, catalog-summary.json, TOOL_SURFACE.md, PERMISSIONS.md, and tests are consistent; otherwise N/A.

## Validation

<!-- Record results; explain checks that are not applicable or cannot be run locally. -->

- [ ] npm ci
- [ ] npm run check
- [ ] npm run build
- [ ] Relevant CI, package, and security checks reviewed.

## Review and merge

<!-- Allow the normal automated review workflow. Resolve substantive findings; request re-review when material changes warrant it. Reviewer completion does not authorize merge. -->
