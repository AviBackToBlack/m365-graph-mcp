# macOS + Claude Desktop: first run from a verified CI artifact

**Status: supervised pilot, not a certified release.** Production acceptance on a clean Mac is still required. This repo has no published GitHub Release or npm registry distribution; the CI artifact expires after its retention period.

Requirements: macOS with Claude Desktop, Chrome/Edge, Node.js 22.12+, and an authenticated GitHub CLI (`gh`). Use FileVault, a dedicated bridge-owned browser profile, and your corporate tenant GUID. Do **not** record Graph access tokens, browser profile data, tenant/client identifiers or Graph responses in issues/logs.

## 1. Pin, download and verify

The following known-successful `main` Package workflow is the pilot source:

- Commit `eedf75b038be17420553773a42ef19f2f9aca45b`
- [Package run 38309631306](https://github.com/AviBackToBlack/m365-graph-mcp/actions/runs/38309631306)
- Artifact `package-sbom-sha256`

In a Mac Terminal:

```sh
set -euo pipefail
repo=AviBackToBlack/m365-graph-mcp
run=38309631306
sha=eedf75b038be17420553773a42ef19f2f9aca45b
bundle="$HOME/Downloads/m365-graph-mcp-pilot-$run"
test "$(gh run view "$run" -R "$repo" --json headSha --jq .headSha)" = "$sha"
test "$(gh run view "$run" -R "$repo" --json conclusion --jq .conclusion)" = success
mkdir -p "$bundle"
gh run download "$run" -R "$repo" -n package-sbom-sha256 -D "$bundle"
cd "$bundle"
shasum -a 256 -c SHA256SUMS
set -- ./*.tgz
test "$#" -eq 1 && test -f "$1"
gh attestation verify "$1" -R "$repo"
```

**Stop** if artifact retrieval, checksum or GitHub attestation fails. This is a GitHub-attested build from the pinned source SHA, not an unspecified latest npm package.

## 2. Install without registry publication

```sh
set -euo pipefail
install_root="$HOME/Library/Application Support/m365-graph-mcp-runtime"
mkdir -p "$install_root"
npm install --prefix "$install_root" --ignore-scripts --no-audit --no-fund "$bundle"/*.tgz
server_js="$install_root/node_modules/m365-graph-mcp/dist/index.js"
bridge_js="$install_root/node_modules/m365-graph-mcp/dist/bridge/main.js"
test -f "$server_js" && test -f "$bridge_js"
node_bin="$(command -v node)"
printf 'Node: %s\nServer: %s\nBridge: %s\n' "$node_bin" "$server_js" "$bridge_js"
```

Copy the *actual absolute paths* printed here into the next steps. Do not assume that the Claude Desktop GUI inherits Terminal's PATH.


## 3. Run the bridge in a separate Terminal

Set the corporate tenant GUID **only in the bridge process**, replacing the placeholder locally:

```sh
M365_GRAPH_TENANT_ID='<CORPORATE-TENANT-GUID>' "$node_bin" "$bridge_js"
```

The bridge starts a dedicated Chrome/Edge profile. Complete Graph Explorer SSO/MFA *inside that browser* and run a permitted Graph request there (for example GET /me) so the bridge can observe a Graph-bound bearer. **Never copy or display the bearer.** Keep the bridge Terminal running during the pilot. The broker uses an owner-restricted local Unix socket, never TCP.

An unavailable/stopped bridge causes a typed `reauth_required` on MCP calls, with **no fallback** to manual tokens. The browser profile remains credential-bearing after shutdown; consult [profile retirement](macos-sm15-acceptance.md).

## 4. Connect Claude Desktop (stdio MCP)

Open Claude Desktop **Settings → Developer → Edit Config**. Its macOS file is `~/Library/Application Support/Claude/claude_desktop_config.json`. Preserve all existing MCP entries; add the one named `m365-graph-mcp` below, changing both example paths to the *exact* Node/server paths printed in step 2:

```json
{
  "mcpServers": {
    "m365-graph-mcp": {
      "command": "/opt/homebrew/bin/node",
      "args": [
        "/Users/REPLACE/Library/Application Support/m365-graph-mcp-runtime/node_modules/m365-graph-mcp/dist/index.js"
      ],
      "env": {
        "M365_GRAPH_TOKEN_SOURCE": "bridge"
      }
    }
  }
}
```

Never put `M365_GRAPH_ACCESS_TOKEN`, `M365_GRAPH_TENANT_ID`, or `M365_GRAPH_CLIENT_ID_PIN` in the **server** configuration. The first is deliberately unused in bridge mode; the other settings, when applicable, belong to the separate bridge. Do not embed login credentials or launch a shell that prints environment variables.

Save valid JSON and **fully restart Claude Desktop**.

## 5. Local acceptance checklist

1. Claude Desktop connects to `m365-graph-mcp` and lists **76 retrieval-only tools**.
2. Stop the bridge and request `get-current-user`: expect typed `reauth_required`, not a fallback token.
3. Start the bridge and complete Graph Explorer interactive sign-in/MFA/request. Call `get-current-user` in Claude Desktop. A `401`/`403` may reflect tenant or delegated-scope policy, never a reason to request extra scopes automatically.
4. Stop/restart the bridge and verify the fail-closed transition and recovery; never paste private output into issue comments.
5. Complete the [macOS bridge acceptance runbook](macos-sm15-acceptance.md) on the real Mac before claiming production readiness.

The bridge diagnostic switch `M365_GRAPH_BRIDGE_ACCEPTANCE_DIAGNOSTICS=1` is a *temporary*, shape-only, bridge-only test setting. Unset it for normal operation. Do not publish even its category-only transcripts without privacy approval.

## Limitations and safe updates

- This is a supervised pilot. The separate bridge **does not** run automatically at login; no LaunchAgent or persistence is installed.
- A CI tarball is not a GitHub Release; verify a *new* pinned run/commit/attestation independently before upgrades. Do not auto-follow `main`.
- Do not use the separate private archive as a distribution source.
- No npm registry publication, mutation operations, scope escalation or hidden Graph proxy is part of this procedure.
