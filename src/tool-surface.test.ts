import { readFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import type { Tool } from "@modelcontextprotocol/sdk/types.js";
import { beforeAll, describe, expect, it } from "vitest";
import type { TokenProvider } from "./auth/token-provider.js";
import { DEFAULT_CATALOG_PATH, loadEndpointCatalog, type EndpointCatalog } from "./catalog.js";
import { GraphClient } from "./graph/client.js";
import { batchTargets } from "./graph/batch.js";
import { refersToBatch } from "./graph/url.js";
import { createServer } from "./server.js";
import { createHandlers } from "./tools/index.js";

/**
 * SM-16 exact tool-surface acceptance. Pins, against the real `endpoints.json` and the production
 * wiring (`createHandlers` + `createServer`), that the advertised MCP surface is exactly the
 * catalog: no extra tool, no missing tool, no duplicate, every tool implemented, no generic
 * proxy-style input anywhere in any input schema, and `graph-batch` as the only multi-request tool.
 *
 * The 76 / 71 GET / 5 POST counts are re-pinned here on purpose (the `TOOL_SURFACE.md` header
 * check needs them), so a catalog change must touch both this file and `catalog.test.ts`. The
 * exact POST allowlist and "no mutation verbs" stay in `catalog.test.ts` only; since the registered
 * set is proven equal to the catalog here, they transitively cover the registered tools (a test
 * file is never imported).
 */

// Deliberate pin: changing endpoints.json (and TOOL_SURFACE.md / catalog-summary.json) must touch
// this number. 76 = after D21 (list-chat-message-replies), D32 (two site-items tools) and D36
// (sites delta) removals.
const PINNED_TOOL_COUNT = 76;
const PINNED_GET_COUNT = 71;
const PINNED_POST_COUNT = 5;

/** Production wiring with a token provider that can never be asked for a token. */
const forbiddenTokenProvider: TokenProvider = {
  getAccessToken: () => Promise.reject(new Error("tool-surface test must not request a token")),
};

let catalog: EndpointCatalog;
let handlerKeys: string[];
let batchTargetNames: string[];
let tools: Tool[];

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  const client = new GraphClient({
    tokenProvider: forbiddenTokenProvider,
    catalog,
    fetch: () => Promise.reject(new Error("tool-surface test must not reach the network")),
  });
  const handlers = createHandlers(client, catalog);
  handlerKeys = [...handlers.keys()];
  batchTargetNames = [...batchTargets(catalog, handlers).keys()].map((e) => e.toolName);
  const server = createServer(catalog, handlers);
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const mcp = new Client({ name: "tool-surface", version: "0.0.0" });
  await Promise.all([server.connect(serverTransport), mcp.connect(clientTransport)]);
  tools = (await mcp.listTools()).tools;
  await mcp.close();
});

describe("registered tool set equals the catalog", () => {
  it("has exactly the pinned number of tools, with no duplicates", () => {
    const names = tools.map((t) => t.name);
    expect(names).toHaveLength(PINNED_TOOL_COUNT);
    expect(new Set(names).size).toBe(PINNED_TOOL_COUNT);
    expect(catalog).toHaveLength(PINNED_TOOL_COUNT);
    expect(catalog.filter((e) => e.method === "get")).toHaveLength(PINNED_GET_COUNT);
    expect(catalog.filter((e) => e.method === "post")).toHaveLength(PINNED_POST_COUNT);
  });

  it("registers no extra and no missing tool relative to the catalog", () => {
    const registered = new Set(tools.map((t) => t.name));
    const cataloged = new Set(catalog.map((e) => e.toolName));
    expect([...registered].filter((n) => !cataloged.has(n))).toEqual([]);
    expect([...cataloged].filter((n) => !registered.has(n))).toEqual([]);
  });

  it("has a handler for every catalog tool and no handler outside the catalog", () => {
    // createServer ignores handler keys that are not in the catalog, so both directions are
    // asserted here: a stray handler would otherwise be dead code that silently drifts.
    const cataloged = catalog.map((e) => e.toolName).sort();
    expect([...handlerKeys].sort()).toEqual(cataloged);
    expect(new Set(handlerKeys).size).toBe(handlerKeys.length);
  });
});

describe("TOOL_SURFACE.md agrees with endpoints.json", () => {
  let doc: string;
  beforeAll(async () => {
    doc = await readFile(join(dirname(DEFAULT_CATALOG_PATH), "TOOL_SURFACE.md"), "utf8");
  });

  it("states the pinned counts in its header", () => {
    expect(doc).toContain(`Selected tools: **${PINNED_TOOL_COUNT}**`);
    expect(doc).toContain(`**${PINNED_GET_COUNT} GET**, **${PINNED_POST_COUNT} POST**`);
  });

  it("lists exactly the catalog's tools with the same method and path, in order", () => {
    // Every line of the tool table must be a well-formed row: a malformed or non-GET/POST row
    // (for example `DELETE`) fails here instead of being skipped by the row pattern.
    const lines = doc.split(/\r?\n/);
    const start = lines.findIndex((l) => l.startsWith("| # | Tool |"));
    expect(start).toBeGreaterThanOrEqual(0);
    const end = lines.findIndex((l, i) => i > start && !l.startsWith("|"));
    const body = lines.slice(start + 2, end === -1 ? undefined : end);
    const ROW = /^\| (\d+) \| `([^`]+)` \| `(GET|POST)` \| `([^`]+)` \|/;
    expect(body.filter((l) => !ROW.test(l))).toEqual([]);
    expect(lines.filter((l) => l.startsWith("|")).length).toBe(body.length + 2);
    const rows = body.map((l) => {
      const m = ROW.exec(l) as RegExpExecArray;
      return { n: Number(m[1]), toolName: m[2], method: m[3], pathPattern: m[4] };
    });
    expect(rows.map((r) => r.n)).toEqual(catalog.map((_, i) => i + 1));
    expect(rows).toEqual(
      catalog.map((e, i) => ({
        n: i + 1,
        toolName: e.toolName,
        method: e.method.toUpperCase(),
        pathPattern: e.pathPattern,
      })),
    );
  });

  it("lists exactly the catalog's POST tools in the explicit POST allowlist", () => {
    const section = doc.split("## Explicit POST allowlist")[1]?.split("### ")[0] ?? "";
    // Every bullet must be a well-formed POST entry; anything else fails instead of being skipped.
    const bullets = section.split(/\r?\n/).filter((l) => l.startsWith("- "));
    const BULLET = /^- `([^`]+)` → `POST ([^`]+)`/;
    expect(bullets.filter((l) => !BULLET.test(l))).toEqual([]);
    const listed = bullets.map((l) => {
      const m = BULLET.exec(l) as RegExpExecArray;
      return [m[1], m[2]];
    });
    const posts = catalog
      .filter((e) => e.method === "post")
      .map((e) => [e.toolName, e.pathPattern]);
    expect(listed).toEqual(posts);
  });
});

// Beyond the method counts pinned above, the exact POST allowlist and "no mutation verbs" are in
// catalog.test.ts, and readOnlyHint on every registered tool is in server.test.ts. The registered
// set equals the catalog (above), so those pins apply transitively to the registered tools.

/**
 * Every property name at any depth of a JSON Schema (`properties` keys), as `path -> name`.
 * Only schema-structural `properties` objects are read; a property that is itself called
 * "properties" is still walked as data.
 */
function propertyNames(schema: unknown, trail = ""): Array<{ at: string; name: string }> {
  const out: Array<{ at: string; name: string }> = [];
  const walk = (node: unknown, at: string): void => {
    if (Array.isArray(node)) {
      node.forEach((item, i) => walk(item, `${at}[${i}]`));
    } else if (typeof node === "object" && node !== null) {
      for (const [key, value] of Object.entries(node)) {
        if (key === "properties" && typeof value === "object" && value !== null) {
          for (const [name, sub] of Object.entries(value)) {
            out.push({ at, name });
            walk(sub, `${at}.${name}`);
          }
        } else {
          walk(value, `${at}.${key}`);
        }
      }
    }
  };
  walk(schema, trail);
  return out;
}

/** Keywords whose value maps names to subschemas. `properties` names are part of the path. */
const SUBSCHEMA_MAP_KEYWORDS: readonly string[] = [
  "properties",
  "patternProperties",
  "$defs",
  "definitions",
  "dependencies",
  "dependentSchemas",
];

/** Keywords whose value is a subschema or an array of subschemas. */
const SUBSCHEMA_KEYWORDS: readonly string[] = [
  "items",
  "prefixItems",
  "additionalItems",
  "anyOf",
  "oneOf",
  "allOf",
  "not",
  "contains",
  "if",
  "then",
  "else",
  "propertyNames",
  "unevaluatedProperties",
  "unevaluatedItems",
  "additionalProperties",
];

/**
 * The single enumerated open object: graph-batch's per-item `headers`, a string->string map that
 * validateExtraHeaders (src/graph/headers.ts) re-validates against the ConsistencyLevel/Prefer
 * allowlist before anything is sent.
 */
const OPEN_BY_DESIGN = "graph-batch.requests.items.headers";

/**
 * Every open object (not `additionalProperties:false`, or using `patternProperties`) and every
 * untyped value schema (`true` or `{}`) reachable through a subschema keyword. Only schema
 * keywords are walked, so data such as `default: {}` or `enum` values is never mistaken for a
 * schema.
 */
function openSchemaProblems(tool: string, schema: unknown): string[] {
  const problems: string[] = [];
  const walk = (node: unknown, at: string): void => {
    const here = `${tool}${at}`;
    if (Array.isArray(node)) {
      node.forEach((item, i) => walk(item, `${at}[${i}]`));
      return;
    }
    if (node === true) {
      problems.push(`${here}: untyped value schema`);
      return;
    }
    if (typeof node !== "object" || node === null) return;
    const s = node as Record<string, unknown>;
    if (Object.keys(s).length === 0) {
      problems.push(`${here}: untyped value schema`);
      return;
    }
    const type = s["type"];
    const objectLike =
      type === "object" ||
      (Array.isArray(type) && type.includes("object")) ||
      "properties" in s ||
      "patternProperties" in s;
    if (
      objectLike &&
      (s["additionalProperties"] !== false || "patternProperties" in s) &&
      here !== OPEN_BY_DESIGN
    ) {
      problems.push(`${here}: open object`);
    }
    for (const [key, value] of Object.entries(s)) {
      if (SUBSCHEMA_MAP_KEYWORDS.includes(key) && typeof value === "object" && value !== null) {
        const prefix = key === "properties" ? at : `${at}.${key}`;
        for (const [name, sub] of Object.entries(value)) walk(sub, `${prefix}.${name}`);
      } else if (key === "anyOf" || key === "oneOf" || key === "allOf") {
        walk(value, at);
      } else if (key === "additionalProperties") {
        if (value !== false && here !== OPEN_BY_DESIGN) walk(value, `${at}.additionalProperties`);
      } else if (SUBSCHEMA_KEYWORDS.includes(key)) {
        walk(value, `${at}.${key}`);
      }
    }
  };
  walk(schema, "");
  return problems;
}

/** Exact (case-insensitive) names that would read as a generic Graph passthrough parameter. */
const PASSTHROUGH_NAMES: readonly string[] = [
  "url",
  "uri",
  "method",
  "path",
  "endpoint",
  "headers",
  "body",
  "host",
];

/**
 * The ONLY places a passthrough-looking name may appear. `graph-batch`'s `requests[]` items carry
 * `method`, `url` and `headers` by design (Graph's batch wire shape), but `validateBatch`
 * (src/graph/batch.ts) accepts only method === "GET", a relative URL matching an implemented,
 * policy-bearing GET catalog path, and the ConsistencyLevel/Prefer header allowlist; there is no
 * `body` or `dependsOn`. Exact name per tool, at exact depth. Nothing else is allowed.
 */
const ALLOWED_PASSTHROUGH_NAMES: ReadonlyMap<string, readonly string[]> = new Map([
  ["graph-batch", ["method", "url", "headers"]],
]);

describe("no generic passthrough input", () => {
  it("exposes no passthrough-named property at any depth, except graph-batch's batch items", () => {
    const offenders: string[] = [];
    for (const tool of tools) {
      const allowed = ALLOWED_PASSTHROUGH_NAMES.get(tool.name) ?? [];
      for (const { at, name } of propertyNames(tool.inputSchema)) {
        if (!PASSTHROUGH_NAMES.includes(name.toLowerCase())) continue;
        // Allowed only as a field of graph-batch's requests[] items, never at the top level.
        const isBatchItemField = tool.name === "graph-batch" && /^\.requests\.items$/.test(at);
        if (!(isBatchItemField && allowed.includes(name)))
          offenders.push(`${tool.name}${at}.${name}`);
      }
    }
    expect(offenders).toEqual([]);
  });

  it("pins graph-batch's exact request-item fields (no body, no dependsOn)", () => {
    const batch = tools.find((t) => t.name === "graph-batch");
    expect(Object.keys(batch?.inputSchema.properties ?? {})).toEqual(["requests"]);
    const names = propertyNames(batch?.inputSchema).map((p) => p.name);
    expect(names.sort()).toEqual(["headers", "id", "method", "requests", "url"]);
  });

  it("offers no absolute-URL input or format anywhere in any input schema", () => {
    // The whole advertised schema, not only `properties`: root-level `format`, `anyOf`, `$defs`,
    // `enum` or `default` would otherwise escape. Only the root `$schema` meta-schema URI (present
    // on graph-batch's advertised schema) is excluded: it names the dialect, not an input.
    for (const tool of tools) {
      const schema = JSON.stringify({ ...tool.inputSchema, $schema: undefined });
      expect(schema, tool.name).not.toMatch(/"format":"(?:uri|url|uri-reference)"|https?:\/\//i);
    }
  });

  it("has no free-form object or untyped value schema anywhere", () => {
    // A record, an open object or an untyped `{}` would accept arbitrary nested keys such as "url"
    // or "body" without any named property for the scan above to see.
    expect(tools.flatMap((t) => openSchemaProblems(t.name, t.inputSchema))).toEqual([]);
  });

  // The real schemas use none of these keywords, so the test above alone cannot show that the
  // walker reaches them. Each case hides an open object under one subschema keyword.
  const OPEN = { type: "object" };
  const CLOSED = { type: "object", additionalProperties: false };
  it.each([
    ["properties", { ...CLOSED, properties: { x: OPEN } }],
    ["patternProperties (itself open)", { ...CLOSED, patternProperties: { "^x": CLOSED } }],
    ["$defs", { ...CLOSED, $defs: { d: OPEN } }],
    ["definitions", { ...CLOSED, definitions: { d: OPEN } }],
    ["dependentSchemas", { ...CLOSED, dependentSchemas: { d: OPEN } }],
    ["dependencies", { ...CLOSED, dependencies: { a: ["b"], d: OPEN } }],
    ["type array", { ...CLOSED, properties: { x: { type: ["object", "null"] } } }],
    ["items", { type: "array", items: OPEN }],
    ["items (tuple)", { type: "array", items: [OPEN] }],
    ["prefixItems", { type: "array", prefixItems: [OPEN] }],
    ["additionalItems", { type: "array", additionalItems: OPEN }],
    ["anyOf", { anyOf: [OPEN] }],
    ["oneOf", { oneOf: [OPEN] }],
    ["allOf", { allOf: [OPEN] }],
    ["not", { not: OPEN }],
    ["contains", { type: "array", contains: OPEN }],
    ["if", { if: OPEN }],
    ["then", { then: OPEN }],
    ["else", { else: OPEN }],
    ["propertyNames", { ...CLOSED, propertyNames: OPEN }],
    ["unevaluatedProperties", { ...CLOSED, unevaluatedProperties: OPEN }],
    ["unevaluatedItems", { type: "array", unevaluatedItems: OPEN }],
  ])("finds an open object reached through %s", (_keyword, schema) => {
    expect(openSchemaProblems("synthetic", schema)).not.toEqual([]);
  });

  it("allows the open headers map only at graph-batch's exact path", () => {
    const shape = {
      ...CLOSED,
      properties: {
        requests: {
          type: "array",
          items: {
            ...CLOSED,
            properties: { headers: { type: "object", additionalProperties: { type: "string" } } },
          },
        },
      },
    };
    expect(openSchemaProblems("graph-batch", shape)).toEqual([]);
    expect(openSchemaProblems("other-tool", shape)).toEqual([
      "other-tool.requests.items.headers: open object",
    ]);
  });

  it("rejects unknown top-level arguments for every tool (additionalProperties:false)", () => {
    for (const tool of tools) expect(tool.inputSchema.additionalProperties, tool.name).toBe(false);
  });
});

describe("graph-batch is the only multi-request tool", () => {
  // No schema property can express "multi-request"; this is approximated by three structural
  // facts plus the pinned batch-item shape above:
  //  1. /$batch is the only catalog path that refers to $batch;
  it("is the only catalog entry that refers to $batch", () => {
    const batchEntries = catalog.filter((e) => refersToBatch(e.pathPattern));
    expect(batchEntries.map((e) => e.toolName)).toEqual(["graph-batch"]);
    expect(batchEntries[0]?.method).toBe("post");
  });

  //  2. it is the only tool whose input schema carries an array of request-like items;
  it("is the only tool whose input schema has nested request items", () => {
    const withRequestItems = tools
      .filter((t) => propertyNames(t.inputSchema).some((p) => p.at !== "" && p.name === "url"))
      .map((t) => t.name);
    expect(withRequestItems).toEqual(["graph-batch"]);
  });

  //  3. its sub-requests can only target implemented GET catalog entries (never POST tools, never
  //     itself). GraphClient.request() refusing $batch is pinned in graph/client.test.ts.
  it("can only target implemented GET catalog entries that opted in with a batchPolicy", () => {
    const get = new Set(catalog.filter((e) => e.method === "get").map((e) => e.toolName));
    expect(batchTargetNames.length).toBeGreaterThan(0);
    for (const name of batchTargetNames) expect(get.has(name), name).toBe(true);
    expect(batchTargetNames).not.toContain("graph-batch");
  });
});
