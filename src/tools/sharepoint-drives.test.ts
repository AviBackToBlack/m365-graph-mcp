import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";
import { sanitizeDrive } from "./sharepoint-drives.js";

/**
 * SM-13 part B: list-sharepoint-site-drives and get-sharepoint-site-drive-by-id. The shared
 * 401/403/token-leak capability matrix lives in sharepoint-capability.test.ts. The two site `items`
 * tools were removed from the catalog (D32; Microsoft: the collection "can't be enumerated"; no API
 * page), and the tests below assert their routes are refused.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const LIST = "list-sharepoint-site-drives";
const GET = "get-sharepoint-site-drive-by-id";
const HOST = "contoso.sharepoint.com";
const GUID_A = "f17e0000-0000-4000-8000-000000000004";
const GUID_B = "f17e0000-0000-4000-8000-000000000005";
const COMPOSITE = `${HOST},${GUID_A},${GUID_B}`;
const DRIVE = "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd";
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

/** A drive as Graph returns it, with extras that must be dropped. */
const rawDrive = (over: Record<string, unknown> = {}) => ({
  "@odata.context": "ctx-secret",
  "@odata.etag": "etag-secret",
  id: DRIVE,
  name: "Documents",
  description: "Team documents",
  driveType: "documentLibrary",
  webUrl: `https://${HOST}/teams/hr/Shared%20Documents`,
  createdDateTime: "2017-07-27T02:41:36Z",
  lastModifiedDateTime: "2018-03-27T07:34:38Z",
  eTag: "eTag-secret",
  createdBy: { user: { displayName: "creator-secret", email: "c-secret@contoso.com" } },
  lastModifiedBy: { user: { displayName: "modifier-secret", id: "uid-secret" } },
  owner: { user: { displayName: "owner-secret", id: "oid-secret" } },
  quota: { total: 1099511627776, used: 294606476, state: "normal" },
  root: { id: "root-secret" },
  items: [{ id: "item-secret" }],
  list: { template: "documentLibrary" },
  sharepointIds: { siteUrl: "siteurl-secret" },
  system: {},
  "@microsoft.graph.downloadUrl": "https://dl-secret",
  ...over,
});
const projectedDrive = {
  id: DRIVE,
  name: "Documents",
  description: "Team documents",
  driveType: "documentLibrary",
  webUrl: `https://${HOST}/teams/hr/Shared%20Documents`,
  createdDateTime: "2017-07-27T02:41:36Z",
  lastModifiedDateTime: "2018-03-27T07:34:38Z",
};
const SECRETS =
  /secret|owner|createdBy|lastModifiedBy|quota|sharepointIds|eTag|etag|@odata|downloadUrl|"root"|"items"|"system"|"list"/;

describe("catalog and surface", () => {
  it("matches the catalog entries (v1.0 GET) and exposes only typed inputs", async () => {
    for (const [name, pattern] of [
      [LIST, "/sites/{site-id}/drives"],
      [GET, "/sites/{site-id}/drives/{drive-id}"],
    ] as const) {
      const e = catalog.find((x) => x.toolName === name);
      expect(e?.method, name).toBe("get");
      expect(e?.pathPattern, name).toBe(pattern);
      expect(e?.["apiVersion"], name).not.toBe("beta");
    }
    const tools = (await mcp.listTools()).tools;
    const schema = (n: string) => tools.find((t) => t.name === n)?.inputSchema;
    const props = (n: string) =>
      Object.keys((schema(n)?.properties ?? {}) as Record<string, unknown>).sort();
    expect(props(LIST)).toEqual(["site-id", "top"]);
    expect(props(GET)).toEqual(["drive-id", "site-id"]);
    expect([...(schema(GET)?.required ?? [])].sort()).toEqual(["drive-id", "site-id"]);
  });

  it("describes the typed contracts, not raw Graph syntax (D33, D34)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    const tip = (n: string) => String(catalog.find((x) => x.toolName === n)?.["llmTip"]);
    for (const n of [LIST, GET]) {
      expect(desc(n)).toMatch(
        /Exposes only id, name, description, driveType, webUrl, createdDateTime and lastModifiedDateTime/,
      );
      expect(desc(n)).toMatch(/owner, creator and modifier identities, quota, root/);
      expect(desc(n)).toMatch(/It does not list drive items or files/);
      expect(desc(n)).toMatch(/file contents and downloads are outside this capability/);
      expect(desc(n)).toMatch(/work\/school accounts only/);
      expect(desc(n)).toMatch(/No raw OData options or caller-supplied headers/);
      expect(desc(n)).toMatch(/normal error mapping/);
      expect(desc(n)).toMatch(/Requires Files\.Read\.All, Sites\.Read\.All or Sites\.Selected/);
      expect(desc(n) + tip(n)).not.toMatch(
        /downloadUrl|list-folder-files|\bFiles\.Read\b(?!\.All)/,
      );
    }
    expect(desc(LIST)).toMatch(
      /top is a typed result limit \(1-200, default 25\) sent as the documented Graph \$top/,
    );
    expect(desc(LIST)).toMatch(/at most 100 drives/);
    expect(desc(LIST)).toMatch(/hidden system drives are not included/);
    expect(desc(LIST)).toMatch(/do not construct or parse composite ids/);
    expect(desc(GET)).toMatch(/drive-id is the opaque id returned by list-sharepoint-site-drives/);
    expect(desc(GET)).toMatch(/never construct, parse or append route segments/);
    expect(desc(GET)).toMatch(/Takes no query options/);
  });

  it("keeps the site items routes out of the catalog and refuses them (D32)", async () => {
    // Microsoft: the site `items` collection "can't be enumerated"; `baseItem` is abstract and
    // `/sites/{id}/items/{id}` has no API page, so neither is a public tool.
    expect(catalog.some((e) => /^\/sites\/\{site-id\}\/items(\/|$)/.test(e.pathPattern))).toBe(
      false,
    );
    const names = (await mcp.listTools()).tools.map((t) => t.name);
    for (const n of ["list-sharepoint-site-items", "get-sharepoint-site-item"]) {
      expect(
        catalog.some((e) => e.toolName === n),
        n,
      ).toBe(false);
      expect(names, n).not.toContain(n);
      calls = [];
      // An unknown tool is an MCP error, never a forwarded request.
      const refused = await mcp
        .callTool({ name: n, arguments: { "site-id": "root", "baseItem-id": "01ABC" } })
        .then((r) => r.isError === true)
        .catch(() => true);
      expect(refused, n).toBe(true);
      expect(calls, n).toHaveLength(0);
    }
  });
});

describe(LIST, () => {
  it("sends exactly GET /v1.0/sites/{id}/drives?$top=25 and projects the allowlist", async () => {
    responses.push(() => json({ "@odata.context": "ctx-secret", value: [rawDrive()] }));
    const { env, text } = await call(LIST, { "site-id": COMPOSITE });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe(`/v1.0/sites/${enc(COMPOSITE)}/drives`);
    expect([...parsed().searchParams.keys()]).toEqual(["$top"]);
    expect(parsed().searchParams.get("$top")).toBe("25");
    expect(env.data).toEqual([projectedDrive]);
    expect(text).not.toMatch(SECRETS);
  });

  it("accepts root, a hostname, and a GUID site-id", async () => {
    for (const id of ["root", HOST, GUID_A]) {
      calls = [];
      responses.push(() => json({ value: [] }));
      const { env } = await call(LIST, { "site-id": id });
      expect(env.ok, id).toBe(true);
      expect(parsed().pathname).toBe(`/v1.0/sites/${enc(id)}/drives`);
    }
  });

  it("rejects an invalid site-id with zero requests (no smuggled colon path, sub-resource or route)", async () => {
    for (const id of [
      "",
      `${HOST}:/sites/hr`,
      `${COMPOSITE}/drive`,
      "a:b",
      "a%3Ab",
      "..",
      "a/b",
      "a%2Fb",
      "x?y=1",
      "x#y",
      "a\\b",
      "a\u0000b",
      "delta",
      "$batch",
      "microsoft.graph.x",
      "host",
    ]) {
      calls = [];
      const { env } = await call(LIST, { "site-id": id });
      expect(env.ok, JSON.stringify(id)).toBe(false);
      expect(env.error?.code, JSON.stringify(id)).toBe("invalid_input");
      expect(calls, JSON.stringify(id)).toHaveLength(0);
    }
  });

  it("offers only top: select/expand/orderby/filter/search and raw $ options fail with zero requests", async () => {
    for (const extra of [
      { select: ["id"] },
      { expand: ["root"] },
      { orderby: ["name"] },
      { filter: "name eq 'x'" },
      { search: "x" },
      { $select: "id,system" },
      { $expand: "root" },
      { "drive-id": DRIVE },
      { top: 0 },
      { top: 201 },
    ]) {
      calls = [];
      const { env } = await call(LIST, { "site-id": "root", ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
    }
  });

  it("top is sent as the documented $top (capped per page) and clips the result", async () => {
    responses.push(() =>
      json({
        value: [rawDrive({ id: "b!a" }), rawDrive({ id: "b!b" }), rawDrive({ id: "b!c" })],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/sites/root/drives?$top=2&$skiptoken=abc`,
      }),
    );
    const { env } = await call(LIST, { "site-id": "root", top: 2 });
    expect(parsed().searchParams.get("$top")).toBe("2");
    expect((env.data as Array<{ id: string }>).map((s) => s.id)).toEqual(["b!a", "b!b"]);
    expect(env.meta).toMatchObject({ count: 2, hasMore: true });
    expect(calls).toHaveLength(1);
    calls = [];
    responses.push(() => json({ value: [] }));
    await call(LIST, { "site-id": "root", top: 150 });
    expect(parsed().searchParams.get("$top")).toBe("100");
  });

  it("follows nextLink pages on the same drives collection", async () => {
    responses.push(() =>
      json({
        value: [rawDrive({ id: "b!a" })],
        "@odata.nextLink": `https://graph.microsoft.com/v1.0/sites/${enc(COMPOSITE)}/drives?$top=5&$skiptoken=abc`,
      }),
    );
    responses.push(() => json({ value: [rawDrive({ id: "b!b" })] }));
    const { env, text } = await call(LIST, { "site-id": COMPOSITE, top: 5 });
    expect((env.data as Array<{ id: string }>).map((s) => s.id)).toEqual(["b!a", "b!b"]);
    expect(calls).toHaveLength(2);
    expect(parsed(1).pathname).toBe(`/v1.0/sites/${enc(COMPOSITE)}/drives`);
    expect(text).not.toMatch(SECRETS);
  });

  it("refuses a nextLink that leaves the drives collection (route switch)", async () => {
    for (const next of [
      `https://graph.microsoft.com/v1.0/sites/${enc(COMPOSITE)}/items?$skiptoken=abc`,
      `https://graph.microsoft.com/v1.0/sites/${enc(COMPOSITE)}/lists?$skiptoken=abc`,
      `https://graph.microsoft.com/v1.0/sites/${enc(COMPOSITE)}/drives/${DRIVE}/root/children`,
      `https://graph.microsoft.com/v1.0/sites/root/drives?$skiptoken=abc`, // other site
      `https://evil.example/v1.0/sites/${enc(COMPOSITE)}/drives?$skiptoken=abc`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [rawDrive({ id: "b!a" })], "@odata.nextLink": next }));
      const { env } = await call(LIST, { "site-id": COMPOSITE, top: 5 });
      expect(env.ok, next).toBe(false);
      expect(env.error?.code, next).toBe("invalid_input");
      expect(calls, next).toHaveLength(1);
    }
  });

  it("an empty result is a successful empty list; a missing value array is a failure", async () => {
    responses.push(() => json({ value: [] }));
    const empty = await call(LIST, { "site-id": "root" });
    expect(empty.env.ok).toBe(true);
    expect(empty.env.data).toEqual([]);
    responses.push(() => json({}));
    const missing = await call(LIST, { "site-id": "root" });
    expect(missing.env.ok).toBe(false);
    expect(missing.env.error?.code).toBe("graph_error");
  });
});

describe(GET, () => {
  it("requests /sites/{id}/drives/{drive-id} with no query and projects the allowlist", async () => {
    responses.push(() => json(rawDrive()));
    const { env, text } = await call(GET, { "site-id": COMPOSITE, "drive-id": DRIVE });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(parsed().pathname).toBe(`/v1.0/sites/${enc(COMPOSITE)}/drives/${DRIVE}`);
    expect(parsed().search).toBe("");
    expect(env.data).toEqual(projectedDrive);
    expect(text).not.toMatch(SECRETS);
  });

  it("rejects an invalid drive-id or site-id with zero requests (no route, content or OData smuggling)", async () => {
    const badDrive = [
      "",
      "abc", // purely alphabetic: could be a bare route name
      "b!a/root/children",
      "b!a/items/x",
      "b!a%2Froot",
      "b!a%252Froot",
      "b!a:/x",
      "b!a%3Ax",
      "b!a?$select=id",
      "b!a#x",
      "b!a\\b",
      "b!a\nb",
      "b!a\u0000b",
      "b!a b",
      "..",
      ".",
      "%2e%2e",
      "$value",
      "$batch",
      "delta",
      "count",
      "content",
      "microsoft.graph.x",
      "root()",
      "b!a(1)",
      "b!".repeat(300),
    ];
    for (const id of badDrive) {
      calls = [];
      const { env } = await call(GET, { "site-id": "root", "drive-id": id });
      expect(env.ok, JSON.stringify(id)).toBe(false);
      expect(env.error?.code, JSON.stringify(id)).toBe("invalid_input");
      expect(calls, JSON.stringify(id)).toHaveLength(0);
    }
    for (const id of ["", "a:b", "..", `${HOST}:/x`, `${COMPOSITE}/drives`, "delta"]) {
      calls = [];
      const { env } = await call(GET, { "site-id": id, "drive-id": DRIVE });
      expect(env.ok, JSON.stringify(id)).toBe(false);
      expect(calls, JSON.stringify(id)).toHaveLength(0);
    }
  });

  it("names the exposed input in a drive-id error", async () => {
    const { env } = await call(GET, { "site-id": "root", "drive-id": "b!a/x" });
    expect(env.error?.message).toMatch(/drive-id/);
  });

  it("offers no query options (select is documented but not offered): extras fail with zero requests", async () => {
    for (const extra of [
      { select: ["id"] },
      { expand: ["root"] },
      { top: 1 },
      { $select: "id,system" },
      { "baseItem-id": "01ABC" },
    ]) {
      calls = [];
      const { env } = await call(GET, { "site-id": "root", "drive-id": DRIVE, ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
    }
  });

  it("fails closed on a non-object 200", async () => {
    for (const body of [[], "x", null]) {
      responses.push(() => json(body));
      const { env } = await call(GET, { "site-id": "root", "drive-id": DRIVE });
      expect(env.ok).toBe(false);
      expect(env.error?.code).toBe("graph_error");
    }
  });

  it("sanitizeDrive keeps scalars only and drops identities, quota, root and downloads", () => {
    expect(sanitizeDrive("x")).toBeUndefined();
    expect(sanitizeDrive({ id: { a: 1 }, name: ["x"], webUrl: "u", owner: { user: {} } })).toEqual({
      webUrl: "u",
    });
    expect(sanitizeDrive(rawDrive())).toEqual(projectedDrive);
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

  it("list: batch projects exactly like the single path and drops the nextLink", async () => {
    responses.push(() => json({ value: [rawDrive()] }));
    const single = (await call(LIST, { "site-id": COMPOSITE })).env.data;
    calls = [];
    responses.push(ok({ "@odata.nextLink": "x-secret", value: [rawDrive()] }));
    const b = await batch(`/sites/${enc(COMPOSITE)}/drives?$top=25`);
    expect(b.env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(bodyOf(b)?.["value"]).toEqual(single);
    expect(b.text).not.toMatch(SECRETS);
  });

  it("get: batch projects exactly like the single path", async () => {
    responses.push(() => json(rawDrive()));
    const single = (await call(GET, { "site-id": COMPOSITE, "drive-id": DRIVE })).env.data;
    responses.push(ok(rawDrive()));
    const b = await batch(`/sites/${enc(COMPOSITE)}/drives/${DRIVE}`);
    expect(b.env.ok).toBe(true);
    expect(bodyOf(b)).toEqual(single);
    expect(bodyOf(b)).toEqual(projectedDrive);
    expect(b.text).not.toMatch(SECRETS);
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const bad: Array<[string, Record<string, string>?]> = [
      ["/sites/root/drives?$select=id,system"],
      ["/sites/root/drives?$expand=root"],
      ["/sites/root/drives?$orderby=name"],
      ["/sites/root/drives?$filter=x"],
      ["/sites/root/drives?$search=x"],
      ["/sites/root/drives?$count=true"],
      ["/sites/root/drives?$top=0"],
      ["/sites/root/drives?$top=201"],
      ["/sites/root/drives?$top=5&$top=6"],
      ["/sites/root/drives", { ConsistencyLevel: "eventual" }],
      [`/sites/root/drives/${DRIVE}?$select=id`],
      [`/sites/root/drives/${DRIVE}?$expand=root`],
      [`/sites/root/drives/${DRIVE}`, { ConsistencyLevel: "eventual" }],
      [`/sites/root/drives/${DRIVE}/root`],
      [`/sites/root/drives/${DRIVE}/root/children`],
      [`/sites/root/drives/${DRIVE}/items/01ABC`],
      [`/sites/root/drives/${DRIVE}/list`],
      [`/sites/root/drives/${DRIVE}/root/content`],
      [`/sites/root/drives/${DRIVE}/special/documents`],
      ["/sites/root/drives/abc"],
      ["/sites/root/drives/..%2Froot"],
      ["/sites/root/drives/%24value"],
      ["/sites/root/drives/delta"],
      ["/sites/a%3Ab/drives"],
      ["/sites/host/drives"],
      ["/sites/delta/drives"],
      [`/sites/${HOST}:/sites/hr:/drives`],
      ["/drives/b!abc"],
      ["/beta/sites/root/drives"],
      ["/sites/root/items"],
      ["/sites/root/items/01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K"],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });
});
