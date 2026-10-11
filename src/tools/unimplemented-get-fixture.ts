import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import type { EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createBatchTool } from "./batch.js";
import { createHandlers } from "./index.js";

/**
 * TEST-ONLY fixture (imported by `*.test.ts`; never by production code).
 *
 * Several suites prove "a catalogued GET whose tool is not implemented is refused with zero Graph
 * requests (single tool call and graph-batch)". Until SM-14 part A that property was shown with
 * `get-user-presence`; once every catalogued GET has a handler, there is no real unimplemented GET
 * left, so the property is proven against a SYNTHETIC catalog entry appended to a COPY of the real
 * catalog (the production `endpoints.json` is never touched). The path has no placeholders, so
 * URL validation accepts it and "not implemented" is the only possible reason for a refusal; the
 * separate "not in the catalog" property stays asserted by each suite's own out-of-catalog cases.
 */
export const SYNTHETIC_GET_TOOL = "zz-synthetic-unimplemented-get";
export const SYNTHETIC_GET_PATH = "/me/zz-synthetic-unimplemented";

/** A copy of `catalog` plus one catalogued-but-unimplemented GET entry. */
export function withUnimplementedGet(catalog: EndpointCatalog): EndpointCatalog {
  return [
    ...catalog,
    {
      pathPattern: SYNTHETIC_GET_PATH,
      method: "get",
      toolName: SYNTHETIC_GET_TOOL,
      presets: ["work"],
      workScopes: ["User.Read"],
    },
  ];
}

type Env = { ok: boolean; error?: { code: string } };

/**
 * Connects a production-wired server (`createHandlers`) over the augmented catalog, to a
 * GraphClient built from the SAME augmented catalog, and probes the synthetic entry both as a tool
 * call and as a `graph-batch` subrequest. `requests` counts every fetch made; it must stay 0.
 */
export async function probeUnimplementedGet(base: EndpointCatalog): Promise<{
  catalogued: boolean;
  hasHandler: boolean;
  direct: Env;
  batch: Env;
  requests: number;
}> {
  const catalog = withUnimplementedGet(base);
  let requests = 0;
  const fetch: FetchLike = () => {
    requests += 1;
    return Promise.resolve(new Response("{}", { status: 200 }));
  };
  const graph = new GraphClient({
    tokenProvider: { getAccessToken: () => Promise.resolve("FAKE_TOKEN_fixture") },
    catalog,
    fetch,
    sleep: () => Promise.resolve(),
  });
  const handlers = createHandlers(graph, catalog);
  const server = createServer(catalog, handlers);
  const [ct, st] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "fixture", version: "0" });
  await Promise.all([server.connect(st), client.connect(ct)]);
  try {
    const direct = await client.callTool({ name: SYNTHETIC_GET_TOOL, arguments: {} });
    const batch = await client.callTool({
      name: "graph-batch",
      arguments: { requests: [{ id: "1", method: "GET", url: SYNTHETIC_GET_PATH }] },
    });
    return {
      catalogued: catalog.some((e) => e.toolName === SYNTHETIC_GET_TOOL && e.method === "get"),
      hasHandler: handlers.has(SYNTHETIC_GET_TOOL),
      direct: direct.structuredContent as Env,
      batch: batch.structuredContent as Env,
      requests,
    };
  } finally {
    await client.close();
  }
}

/**
 * Positive control: with the synthetic entry explicitly PERMITTED as a batch target (what a real
 * tool's `batchPolicy` would do), the same subrequest URL passes validation and reaches the
 * network. This proves that, in `probeUnimplementedGet`, "no handler" is the only reason for the
 * refusal. Returns whether the call succeeded and how many fetches were made.
 */
export async function probePermittedSyntheticBatch(
  base: EndpointCatalog,
): Promise<{ ok: boolean; requests: number }> {
  const catalog = withUnimplementedGet(base);
  const entry = catalog.find((e) => e.toolName === SYNTHETIC_GET_TOOL);
  if (entry === undefined) throw new Error("synthetic entry missing");
  let requests = 0;
  const fetch: FetchLike = () => {
    requests += 1;
    return Promise.resolve(
      new Response(JSON.stringify({ responses: [{ id: "1", status: 200, body: { v: 1 } }] }), {
        status: 200,
        headers: { "content-type": "application/json" },
      }),
    );
  };
  const graph = new GraphClient({
    tokenProvider: { getAccessToken: () => Promise.resolve("FAKE_TOKEN_fixture") },
    catalog,
    fetch,
    sleep: () => Promise.resolve(),
  });
  const server = createServer(
    catalog,
    new Map([
      [
        "graph-batch",
        createBatchTool(graph, catalog, new Map([[entry, { check: () => undefined }]])),
      ],
    ]),
  );
  const [ct, st] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "fixture", version: "0" });
  await Promise.all([server.connect(st), client.connect(ct)]);
  try {
    const r = await client.callTool({
      name: "graph-batch",
      arguments: { requests: [{ id: "1", method: "GET", url: SYNTHETIC_GET_PATH }] },
    });
    return { ok: (r.structuredContent as Env).ok, requests };
  } finally {
    await client.close();
  }
}
