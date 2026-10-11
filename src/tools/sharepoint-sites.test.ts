import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";
import { sanitizeSite } from "./sharepoint-sites.js";

/**
 * SM-13 part A: search-sharepoint-sites, get-sharepoint-site, get-sharepoint-site-by-path. The shared
 * 401/403/token-leak capability matrix for these three tools lives in sharepoint-capability.test.ts.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const SEARCH = "search-sharepoint-sites";
const GET = "get-sharepoint-site";
const BY_PATH = "get-sharepoint-site-by-path";
const HOST = "contoso.sharepoint.com";
const GUID_A = "f17e0000-0000-4000-8000-000000000004";
const GUID_B = "f17e0000-0000-4000-8000-000000000005";
const COMPOSITE = `${HOST},${GUID_A},${GUID_B}`;
const enc = encodeURIComponent;

interface Call {
  url: string;
  method: string;
  headers: Record<string, string>;
}

let catalog: EndpointCatalog;
let mcp: Client;
let calls: Call[];
let responses: Array<() => Response>;

const json = (body: unknown, status = 200): Response =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });

type Env = {
  ok: boolean;
  data?: unknown;
  meta?: Record<string, unknown>;
  error?: { code: string; message: string; details?: Record<string, unknown> };
};
const call = async (name: string, args: Record<string, unknown> = {}) => {
  const r = await mcp.callTool({ name, arguments: args });
  return { r, env: r.structuredContent as Env, text: JSON.stringify(r) };
};
const parsed = (i = 0) => new URL((calls[i] as Call).url);

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  const fetch: FetchLike = (url, init) => {
    calls.push({
      url,
      method: String(init.method),
      headers: { ...(init.headers as Record<string, string>) },
    });
    const next = responses.shift();
    return Promise.resolve(next ? next() : json({ value: [] }));
  };
  const graph = new GraphClient({
    tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
    catalog,
    fetch,
    sleep: () => Promise.resolve(),
  });
  const server = createServer(catalog, createHandlers(graph, catalog));
  const [ct, st] = InMemoryTransport.createLinkedPair();
  mcp = new Client({ name: "t", version: "0" });
  await Promise.all([server.connect(st), mcp.connect(ct)]);
});
afterAll(async () => {
  await mcp.close();
});
beforeEach(() => {
  calls = [];
  responses = [];
});

/** A site as Graph returns it, with extras that must be dropped. */
const rawSite = (over: Record<string, unknown> = {}) => ({
  "@odata.context": "ctx-secret",
  "@odata.etag": "etag-secret",
  id: COMPOSITE,
  name: "1drvteam",
  displayName: "OneDrive Team Site",
  description: "Team site",
  webUrl: `https://${HOST}/teams/1drvteam`,
  createdDateTime: "2017-05-09T20:56:00Z",
  lastModifiedDateTime: "2017-05-09T20:56:01Z",
  isPersonalSite: false,
  eTag: "eTag-secret",
  root: {},
  sharepointIds: { siteUrl: "siteurl-secret", tenantId: "tenant-secret" },
  siteCollection: { hostname: HOST, dataLocationCode: "EUR", root: {}, extra: "sc-secret" },
  lists: [{ id: "list-secret" }],
  drives: [{ id: "drive-secret" }],
  ...over,
});
const projectedSite = {
  id: COMPOSITE,
  name: "1drvteam",
  displayName: "OneDrive Team Site",
  description: "Team site",
  webUrl: `https://${HOST}/teams/1drvteam`,
  createdDateTime: "2017-05-09T20:56:00Z",
  lastModifiedDateTime: "2017-05-09T20:56:01Z",
  isPersonalSite: false,
  siteCollection: { hostname: HOST, dataLocationCode: "EUR" },
};
const SECRETS = /secret|sharepointIds|eTag|etag|@odata|"lists"|"drives"|"root"/;

describe("catalog and surface", () => {
  it("matches the catalog entries (v1.0 GET)", () => {
    for (const [name, pattern] of [
      [SEARCH, "/sites"],
      [GET, "/sites/{site-id}"],
      [BY_PATH, "/sites/{site-id}:/{path}"],
    ] as const) {
      const e = catalog.find((x) => x.toolName === name);
      expect(e?.method, name).toBe("get");
      expect(e?.pathPattern, name).toBe(pattern);
      expect(e?.["apiVersion"], name).not.toBe("beta");
    }
    expect(catalog.find((x) => x.toolName === BY_PATH)?.["skipEncoding"]).toEqual(["path"]);
  });

  it("exposes only the typed inputs (relative-path, never a bare path)", async () => {
    const tools = (await mcp.listTools()).tools;
    const schema = (n: string) => tools.find((t) => t.name === n)?.inputSchema;
    const props = (n: string) =>
      Object.keys((schema(n)?.properties ?? {}) as Record<string, unknown>).sort();
    expect(props(SEARCH)).toEqual(["search", "top"]);
    expect(schema(SEARCH)?.required).toEqual(["search"]);
    expect(props(GET)).toEqual(["site-id"]);
    expect(props(BY_PATH)).toEqual(["relative-path", "site-id"]);
    expect([...(schema(BY_PATH)?.required ?? [])].sort()).toEqual(["relative-path", "site-id"]);
  });

  it("describes the typed contracts, not raw Graph syntax (D30, D31)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    const tip = (n: string) => String(catalog.find((x) => x.toolName === n)?.["llmTip"]);
    for (const n of [SEARCH, GET, BY_PATH]) {
      expect(desc(n)).toMatch(/work\/school accounts only/);
      expect(desc(n)).toMatch(
        /id, name, displayName, description, webUrl, createdDateTime, lastModifiedDateTime, isPersonalSite and siteCollection hostname and dataLocationCode; everything else is dropped/,
      );
      expect(desc(n)).toMatch(/No raw OData options or caller-supplied headers/);
      expect(desc(n)).toMatch(/normal error mapping/);
    }
    expect(desc(SEARCH)).toMatch(/search is required free-text input \(max 200 characters/);
    expect(desc(SEARCH)).toMatch(/wildcard-only input such as \* are refused/);
    expect(desc(SEARCH)).toMatch(/never enumerates every site/);
    expect(desc(SEARCH)).toMatch(/bare search=<text> parameter, not \$search/);
    expect(desc(SEARCH)).toMatch(
      /top is a local result limit \(1-200, default 25\): no Graph \$top/,
    );
    expect(desc(SEARCH)).toMatch(
      /Requires Sites\.Read\.All \(delegated\); Sites\.Selected is not supported/,
    );
    for (const n of [GET, BY_PATH]) {
      expect(desc(n)).toMatch(/Takes no query options/);
      expect(desc(n)).toMatch(/Requires Sites\.Read\.All or Sites\.Selected \(delegated\)/);
      expect(desc(n)).toMatch(/grant for this app on the target site/);
    }
    expect(desc(GET)).toMatch(/Path addressing belongs to get-sharepoint-site-by-path/);
    expect(desc(BY_PATH)).toMatch(/site-id must be a plain SharePoint hostname/);
    expect(desc(BY_PATH)).toMatch(/without a leading or trailing slash/);
    expect(desc(BY_PATH) + tip(BY_PATH)).not.toMatch(
      /(?<!-)\bpath is\b|list-sharepoint-site-drives/,
    );
    expect(tip(BY_PATH)).toMatch(/relative-path is the site path/);
  });
});

describe(SEARCH, () => {
  it("sends exactly GET /v1.0/sites?search=<text> (bare name, no quotes, no other option)", async () => {
    responses.push(() => json({ "@odata.context": "ctx-secret", value: [rawSite()] }));
    const { env, text } = await call(SEARCH, { search: "hr team" });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(calls[0]?.url).toBe("https://graph.microsoft.com/v1.0/sites?search=hr%20team");
    expect(parsed().searchParams.get("$search")).toBeNull();
    expect(parsed().searchParams.get("$top")).toBeNull();
    expect(env.data).toEqual([projectedSite]);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
    expect(text).not.toMatch(SECRETS);
  });

  it("treats the text as data: reserved URL characters stay inside the one search value", async () => {
    responses.push(() => json({ value: [] }));
    await call(SEARCH, { search: "a&$top=1&search=b#c ?" });
    expect(calls[0]?.url).toBe(
      "https://graph.microsoft.com/v1.0/sites?search=a%26%24top%3D1%26search%3Db%23c%20%3F",
    );
    expect([...parsed().searchParams.keys()]).toEqual(["search"]);
  });

  it("requires search: a bare GET /sites (enumeration) is never sent", async () => {
    for (const args of [{}, { top: 5 }, { search: undefined }]) {
      calls = [];
      const { env } = await call(SEARCH, args);
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(env.error?.code).toBe("invalid_input");
      expect(calls, JSON.stringify(args)).toHaveLength(0);
    }
  });

  it("rejects an empty, wildcard-only, control-character or over-long search text", async () => {
    for (const search of [
      "",
      " ",
      "*",
      "**",
      "* *",
      "?",
      "-",
      "a\nb",
      "a\u0000b",
      "x".repeat(201),
    ]) {
      calls = [];
      const { env } = await call(SEARCH, { search });
      expect(env.ok, JSON.stringify(search)).toBe(false);
      expect(env.error?.code, JSON.stringify(search)).toBe("invalid_input");
      expect(calls, JSON.stringify(search)).toHaveLength(0);
    }
    responses.push(() => json({ value: [] }));
    expect((await call(SEARCH, { search: "x".repeat(200) })).env.ok).toBe(true);
    responses.push(() => json({ value: [] }));
    expect((await call(SEARCH, { search: "hr*" })).env.ok).toBe(true);
  });

  it("rejects options outside the typed contract with zero requests", async () => {
    for (const extra of [
      { top: 0 },
      { top: 201 },
      { top: 1.5 },
      { select: ["id"] },
      { filter: "name eq 'x'" },
      { orderby: [{ property: "createdDateTime" }] },
      { skip: 1 },
      { expand: ["drives"] },
      { $search: "x" },
      { $top: 5 },
      { $select: "id" },
      { "site-id": COMPOSITE },
      { path: "sites/x" },
      { headers: { ConsistencyLevel: "eventual" } },
    ]) {
      calls = [];
      const { env } = await call(SEARCH, { search: "x", ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
    }
  });

  it("top is a local limit: no $top on the wire, extra items clipped, hasMore reported", async () => {
    responses.push(() =>
      json({
        value: [rawSite({ id: "a" }), rawSite({ id: "b" }), rawSite({ id: "c" })],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites?search=x&$skiptoken=abc",
      }),
    );
    const { env } = await call(SEARCH, { search: "x", top: 2 });
    expect(parsed().searchParams.get("$top")).toBeNull();
    expect((env.data as Array<{ id: string }>).map((s) => s.id)).toEqual(["a", "b"]);
    expect(env.meta).toMatchObject({ count: 2, hasMore: true });
    expect(calls).toHaveLength(1);
  });

  it("follows nextLink pages on the same /sites endpoint and keeps the search", async () => {
    responses.push(() =>
      json({
        value: [rawSite({ id: "a" })],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites?search=hr&$skiptoken=abc",
      }),
    );
    responses.push(() => json({ value: [rawSite({ id: "b" })] }));
    const { env } = await call(SEARCH, { search: "hr", top: 5 });
    expect((env.data as Array<{ id: string }>).map((s) => s.id)).toEqual(["a", "b"]);
    expect(calls).toHaveLength(2);
    expect(parsed(1).pathname).toBe("/v1.0/sites");
    expect(parsed(1).searchParams.get("search")).toBe("hr");
  });

  it("refuses a nextLink that leaves /sites (route switch)", async () => {
    responses.push(() =>
      json({
        value: [rawSite({ id: "a" })],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/sites/${COMPOSITE}/drives?$skiptoken=abc`,
      }),
    );
    const { env } = await call(SEARCH, { search: "hr", top: 5 });
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(1);
  });

  it("an empty result is a successful empty list; a missing value array is a failure", async () => {
    responses.push(() => json({ value: [] }));
    const empty = await call(SEARCH, { search: "nothing" });
    expect(empty.env.ok).toBe(true);
    expect(empty.env.data).toEqual([]);
    responses.push(() => json({}));
    const missing = await call(SEARCH, { search: "nothing" });
    expect(missing.env.ok).toBe(false);
    expect(missing.env.error?.code).toBe("graph_error");
  });
});

describe(GET, () => {
  it.each([
    ["root", "root"],
    [HOST, HOST],
    [COMPOSITE, enc(COMPOSITE)],
    [`${GUID_A},${GUID_B}`, enc(`${GUID_A},${GUID_B}`)],
    [GUID_A, GUID_A],
  ])("requests /sites/%s with no query and projects the allowlist", async (id, wire) => {
    responses.push(() => json(rawSite()));
    const { env, text } = await call(GET, { "site-id": id });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe(`/v1.0/sites/${wire}`);
    expect(parsed().search).toBe("");
    expect(env.data).toEqual(projectedSite);
    expect(text).not.toMatch(SECRETS);
  });

  it("rejects a site-id that could smuggle a colon path, sub-resource or route (zero requests)", async () => {
    for (const id of [
      "",
      `${HOST}:/sites/hr`,
      `${HOST}:`,
      `root:/drive`,
      `${COMPOSITE}:/drives`,
      "a:b",
      "a%3Ab",
      "a%253Ab",
      "../me",
      "..",
      ".",
      "a/b",
      `${COMPOSITE}/drives`,
      "a%2Fb",
      "a%2e%2e",
      "x?y=1",
      "x#y",
      "a\\b",
      "a\nb",
      "a\u0000b",
      "a b",
      "delta",
      "delta()",
      "$batch",
      "$value",
      "microsoft.graph.x",
      "host",
      `${HOST},${GUID_A}`,
      `${HOST},${GUID_A},${GUID_B},x`,
      `${HOST},not-a-guid,${GUID_B}`,
    ]) {
      calls = [];
      const { env } = await call(GET, { "site-id": id });
      expect(env.ok, JSON.stringify(id)).toBe(false);
      expect(env.error?.code, JSON.stringify(id)).toBe("invalid_input");
      expect(calls, JSON.stringify(id)).toHaveLength(0);
    }
  });

  it("offers no query options (none documented): extras fail with zero requests", async () => {
    for (const extra of [
      { select: ["id"] },
      { expand: ["drives"] },
      { top: 1 },
      { search: "x" },
      { $select: "id" },
      { "relative-path": "sites/x" },
      { path: "sites/x" },
    ]) {
      calls = [];
      const { env } = await call(GET, { "site-id": "root", ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
    }
  });

  it("fails closed on a non-object 200", async () => {
    for (const body of [[], "x", null]) {
      responses.push(() => json(body));
      const { env } = await call(GET, { "site-id": "root" });
      expect(env.ok).toBe(false);
      expect(env.error?.code).toBe("graph_error");
    }
  });

  it("sanitizeSite keeps scalars only and reduces siteCollection", () => {
    expect(sanitizeSite("x")).toBeUndefined();
    expect(
      sanitizeSite({ id: { a: 1 }, displayName: ["x"], webUrl: "u", siteCollection: "s" }),
    ).toEqual({ webUrl: "u" });
    expect(sanitizeSite({ id: "i", siteCollection: { hostname: "h", root: {}, x: "y" } })).toEqual({
      id: "i",
      siteCollection: { hostname: "h" },
    });
  });
});

describe(BY_PATH, () => {
  it("requests /sites/{host}:/{path} with the separators literal and each segment encoded", async () => {
    responses.push(() => json(rawSite()));
    const { env, text } = await call(BY_PATH, {
      "site-id": HOST,
      "relative-path": "sites/hr team",
    });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(calls[0]?.url).toBe(`https://graph.microsoft.com/v1.0/sites/${HOST}:/sites/hr%20team`);
    expect(env.data).toEqual(projectedSite);
    expect(text).not.toMatch(SECRETS);
  });

  it("handles nested paths and the teams/ prefix", async () => {
    responses.push(() => json(rawSite()));
    await call(BY_PATH, { "site-id": HOST, "relative-path": "teams/hr/benefits" });
    expect(calls[0]?.url).toBe(`https://graph.microsoft.com/v1.0/sites/${HOST}:/teams/hr/benefits`);
  });

  it("never ends in a ':/resource' suffix: the URL stops at the path", async () => {
    responses.push(() => json(rawSite()));
    await call(BY_PATH, { "site-id": HOST, "relative-path": "sites/hr" });
    const path = parsed().pathname;
    expect(path.match(/:/g)).toHaveLength(1);
    expect(path.endsWith("/sites/hr")).toBe(true);
    expect(parsed().search).toBe("");
  });

  it("rejects an invalid relative-path with zero requests", async () => {
    for (const bad of [
      "",
      "/sites/hr",
      "sites/hr/",
      "sites//hr",
      "..",
      "sites/..",
      "../me",
      "sites/../../me/messages",
      "sites/%2e%2e/x",
      "sites/%252e%252e/x",
      "sites/hr:/drive",
      "sites/hr:",
      "sites/hr%3A/drive",
      "sites/hr%253A",
      "sites/hr?$select=id",
      "sites/hr%3Fx",
      "sites/hr#frag",
      "sites/hr%23frag",
      "sites/hr%2Fdrive",
      "sites/hr%252Fdrive",
      "sites\\hr",
      "sites/hr%5Cx",
      "sites/hr\nx",
      "sites/hr%00",
      "sites/$batch",
      "sites/$value",
      "sites/microsoft.graph.x",
      "sites/a(1)",
      "sites/%zz",
      "sites/hr%2F..%2Fx",
    ]) {
      calls = [];
      const { env } = await call(BY_PATH, { "site-id": HOST, "relative-path": bad });
      expect(env.ok, JSON.stringify(bad)).toBe(false);
      expect(env.error?.code, JSON.stringify(bad)).toBe("invalid_input");
      expect(calls, JSON.stringify(bad)).toHaveLength(0);
    }
  });

  it("names the exposed relative-path input, never the catalog placeholder, in errors", async () => {
    for (const bad of ["sites/hr:x", "sites//hr", "sites/.."]) {
      const { env } = await call(BY_PATH, { "site-id": HOST, "relative-path": bad });
      expect(env.error?.message, bad).toMatch(/Path parameter relative-path /);
      expect(env.error?.message, bad).not.toMatch(/parameter path\b/);
    }
  });

  it("requires a plain hostname as site-id (not root, GUID or composite id, nor a path)", async () => {
    for (const id of [
      "root",
      GUID_A,
      `${GUID_A},${GUID_B}`,
      COMPOSITE,
      "host",
      `${HOST}:/x`,
      `${HOST}/sites`,
      "..",
      "a b.example",
    ]) {
      calls = [];
      const { env } = await call(BY_PATH, { "site-id": id, "relative-path": "sites/hr" });
      expect(env.ok, id).toBe(false);
      expect(env.error?.code, id).toBe("invalid_input");
      expect(calls, id).toHaveLength(0);
    }
  });

  it("requires relative-path and rejects the bare catalog name `path` and extra options", async () => {
    for (const args of [
      { "site-id": HOST },
      { "site-id": HOST, path: "sites/hr" },
      { "site-id": HOST, "relative-path": "sites/hr", path: "sites/other" },
      { "site-id": HOST, "relative-path": "sites/hr", select: ["id"] },
      { "site-id": HOST, "relative-path": "sites/hr", $select: "id" },
    ]) {
      calls = [];
      const { env } = await call(BY_PATH, args);
      expect(env.ok, JSON.stringify(args)).toBe(false);
      expect(calls, JSON.stringify(args)).toHaveLength(0);
    }
  });
});

describe("graph-batch reachability and projection parity", () => {
  const batch = (url: string, headers?: Record<string, string>) =>
    call("graph-batch", {
      requests: [{ id: "1", method: "GET", url, ...(headers ? { headers } : {}) }],
    });
  const ok = (body: unknown) => () => json({ responses: [{ id: "1", status: 200, body }] });
  const bodyOf = (r: { env: Env }) =>
    (r.env.data as { responses: Array<{ body?: Record<string, unknown> }> }).responses[0]?.body;

  it("search: batch sends the bare search param and projects like the single path", async () => {
    responses.push(() => json({ value: [rawSite()] }));
    const single = (await call(SEARCH, { search: "hr team" })).env.data;
    calls = [];
    responses.push(ok({ "@odata.nextLink": "x-secret", value: [rawSite()] }));
    const b = await batch("/sites?search=hr%20team");
    expect(b.env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(bodyOf(b)?.["value"]).toEqual(single);
    expect(b.text).not.toMatch(SECRETS);
  });

  it("get: batch projects exactly like the single path", async () => {
    responses.push(() => json(rawSite()));
    const single = (await call(GET, { "site-id": COMPOSITE })).env.data;
    responses.push(ok(rawSite()));
    const b = await batch(`/sites/${enc(COMPOSITE)}`);
    expect(b.env.ok).toBe(true);
    expect(bodyOf(b)).toEqual(single);
    expect(bodyOf(b)).toEqual(projectedSite);
    expect(b.text).not.toMatch(SECRETS);
  });

  it("by-path: batch projects exactly like the single path", async () => {
    responses.push(() => json(rawSite()));
    const single = (await call(BY_PATH, { "site-id": HOST, "relative-path": "sites/hr team" })).env
      .data;
    calls = [];
    responses.push(ok(rawSite()));
    const b = await batch(`/sites/${HOST}:/sites/hr%20team`);
    expect(b.env.ok).toBe(true);
    expect(bodyOf(b)).toEqual(single);
    expect(b.text).not.toMatch(SECRETS);
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const bad: Array<[string, Record<string, string>?]> = [
      // search: required, bare name only, no other option
      ["/sites"],
      ["/sites?search="],
      ["/sites?search=*"],
      ["/sites?search=%20"],
      ["/sites?$search=hr"],
      [`/sites?$search=${enc('"hr"')}`],
      ["/sites?search=hr&search=hr"],
      ["/sites?search=hr&$top=5"],
      ["/sites?search=hr&$select=id"],
      ["/sites?search=hr&$filter=x"],
      ["/sites?search=hr&$orderby=createdDateTime"],
      ["/sites?search=hr&$expand=drives"],
      ["/sites?search=hr&$count=true"],
      ["/sites?search=hr+team"],
      [`/sites?search=${"x".repeat(201)}`],
      ["/sites?search=hr", { ConsistencyLevel: "eventual" }],
      // get: no options
      ["/sites/root?$select=id"],
      ["/sites/root?$expand=drives"],
      ["/sites/root", { ConsistencyLevel: "eventual" }],
      // site-id shape
      [`/sites/${enc(`${HOST}:/sites/hr`)}`],
      ["/sites/a%3Ab"],
      ["/sites/host"],
      // D36: the application-only sites delta is not in the catalog.
      ["/sites/delta"],
      ["/sites/delta()"],
      ["/sites/delta?token=latest"],
      ["/sites/delta()?$deltatoken=abc"],
      [`/sites/${COMPOSITE}/drive`],
      // by-path
      [`/sites/${HOST}:`],
      [`/sites/${HOST}:/`],
      [`/sites/${HOST}:/sites/hr:/drive/root/children`],
      [`/sites/${HOST}:/sites/hr%3A/drive`],
      [`/sites/${HOST}:/sites/../x`],
      [`/sites/${HOST}:/sites/hr?$select=id`],
      [`/sites/${HOST}:/sites/hr`, { ConsistencyLevel: "eventual" }],
      [`/sites/root:/sites/hr`],
      [`/sites/${COMPOSITE}:/sites/hr`],
      [`/beta/sites/root`],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });

  it("refuses the removed get-sharepoint-sites-delta tool by name (D36)", async () => {
    const names = (await mcp.listTools()).tools.map((t) => t.name);
    expect(names).not.toContain("get-sharepoint-sites-delta");
    calls = [];
    const refused = await mcp
      .callTool({ name: "get-sharepoint-sites-delta", arguments: {} })
      .then((r) => r.isError === true)
      .catch(() => true);
    expect(refused).toBe(true);
    expect(calls).toHaveLength(0);
  });
});
