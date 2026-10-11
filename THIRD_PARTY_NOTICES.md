# Third-party notices

## Softeria ms-365-mcp-server

- Project: [`Softeria/ms-365-mcp-server`](https://github.com/Softeria/ms-365-mcp-server)
- License: MIT (reproduced in full below, from the project's `LICENSE` file)
- Reference commit: `3683f55f3917885d80f31cb6f5423ef6da4cee52` (the license text is identical at `a73c42797beaf52ff66a8fd29b4fd6fdede6cbb1`)

### What is derived from it

`endpoints.json`, the capability catalog shipped with this package, is derived from Softeria's endpoint metadata (`src/endpoints.json`):

- **Selection.** All 76 catalog entries are a selection of Softeria entries with the same `toolName` and `pathPattern`.
- **Schema.** The catalog keeps a subset of Softeria's per-entry fields: `toolName`, `method`, `pathPattern`, `scopes`, `workScopes`, `llmTip`, `descriptionOverride`, `presets`, `readOnly`, `requestBodySchema`, `apiVersion`, `skipEncoding`, `supportsTimezone`, `supportsExpandExtendedProperties`.
- **Text.** Some `llmTip` and `descriptionOverride` texts are retained from Softeria, and others have since been corrected or rewritten for this project.

`catalog-summary.json` and `TOOL_SURFACE.md` are generated from that catalog.

### What is not derived from it

The runtime implementation is not a copy or fork of Softeria. That covers:

- the MCP server;
- tool handlers and input validation;
- the Microsoft Graph client;
- the authentication and token bridge;
- the security controls.

Softeria was consulted as prior art for Microsoft Graph edge cases. This project's permissions and endpoint semantics are verified against current Microsoft documentation and are not inherited from Softeria.

### License

```text
MIT License

Copyright (c) 2025 Softeria

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
