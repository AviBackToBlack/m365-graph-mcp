import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { sanitizeFilesFolder, sanitizeTab } from "./channel-tabs.js";
import { createHandlers } from "./index.js";

/**
 * SM-12 part C: list-channel-tabs and get-channel-files-folder. The shared 401/403/token-leak
 * capability cases for these two tools live in teams-capability.test.ts.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const TABS = "list-channel-tabs";
const FOLDER = "get-channel-files-folder";
const TID = "f17e0000-0000-4000-8000-000000000020";
const OTHER_TID = "f17e0000-0000-4000-8000-000000000009";
const CID = "19:4a95f7d8db4c4e7fae857bcebe0623e6@thread.tacv2";
const OTHER_CID = "19:982abbfca323a582f0a6d00ae2deca@thread.tacv2";
const enc = encodeURIComponent;
const ECID = enc(CID);
const IDS = { "team-id": TID, "channel-id": CID };

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

/** A tab as Graph returns it, configuration and extras included (must be dropped). */
const rawTab = (over: Record<string, unknown> = {}) => ({
  "@odata.etag": "etag-secret",
  id: "f17e0000-0000-4000-8000-000000000008",
  displayName: "My Contoso Tab",
  webUrl: "https://teams.microsoft.com/l/channel/19%3Aabc%40thread.tacv2/tab%3A%3Ax?label=t",
  sortOrderIndex: "20",
  configuration: {
    entityId: "entity-secret",
    contentUrl: "https://www.contoso.com/Orders/tabView?token=content-secret",
    websiteUrl: "https://www.contoso.com/Orders?token=website-secret",
    removeUrl: "https://www.contoso.com/uninstallTab?token=remove-secret",
    custom: "custom-secret",
  },
  teamsApp: {
    id: "f17e0000-0000-4000-8000-000000000001",
    displayName: "Contoso",
    distributionMethod: "store",
    appDefinitions: [{ secret: "appdef-secret" }],
    externalId: "ext-secret",
  },
  ...over,
});
const projectedTab = (withApp: boolean) => ({
  id: "f17e0000-0000-4000-8000-000000000008",
  displayName: "My Contoso Tab",
  webUrl: "https://teams.microsoft.com/l/channel/19%3Aabc%40thread.tacv2/tab%3A%3Ax?label=t",
  sortOrderIndex: "20",
  ...(withApp
    ? {
        teamsApp: {
          id: "f17e0000-0000-4000-8000-000000000001",
          displayName: "Contoso",
          distributionMethod: "store",
        },
      }
    : {}),
});

const rawFolder = (over: Record<string, unknown> = {}) => ({
  "@odata.context": "ctx-secret",
  "@microsoft.graph.downloadUrl": "https://download.example/dl-secret",
  id: "01H7CFEKENJSSIUHGADZBKODARINQC5JMD",
  createdDateTime: "0001-01-01T00:00:00Z",
  lastModifiedDateTime: "2020-01-23T18:47:13Z",
  name: "Documentation Planning",
  webUrl: "https://contoso.sharepoint.com/teams/x/Shared%20Documents/Documentation%20Planning",
  size: 2374080,
  eTag: "etag-secret",
  cTag: "ctag-secret",
  createdBy: { user: { displayName: "creator-secret", id: "u1" } },
  lastModifiedBy: { user: { displayName: "modifier-secret", id: "u2" } },
  shared: { scope: "users", owner: { user: { displayName: "owner-secret" } } },
  permissions: [{ link: { webUrl: "https://share.example/s/link-secret" } }],
  content: "content-secret",
  fileSystemInfo: { createdDateTime: "2020-01-23T18:47:12Z", lastModifiedDateTime: "x" },
  parentReference: {
    driveId: "b!2SInBlQrN0K8",
    driveType: "documentLibrary",
    id: "01PARENT",
    path: "/drives/x/root:/path-secret",
    siteId: "site-secret",
    sharepointIds: { listItemUniqueId: "sp-secret" },
  },
  folder: { childCount: 7, view: { viewType: "view-secret" } },
  ...over,
});
const projectedFolder = {
  id: "01H7CFEKENJSSIUHGADZBKODARINQC5JMD",
  createdDateTime: "0001-01-01T00:00:00Z",
  lastModifiedDateTime: "2020-01-23T18:47:13Z",
  name: "Documentation Planning",
  webUrl: "https://contoso.sharepoint.com/teams/x/Shared%20Documents/Documentation%20Planning",
  size: 2374080,
  parentReference: { driveId: "b!2SInBlQrN0K8", driveType: "documentLibrary", id: "01PARENT" },
  folder: { childCount: 7 },
};

describe("catalog and surface", () => {
  it("matches the catalog entries (v1.0 GET, team + channel placeholders only)", () => {
    for (const [name, tail, scope] of [
      [TABS, "tabs", "TeamsTab.Read.All"],
      // D28: the documented Files permission, not ChannelSettings.Read.All.
      [FOLDER, "filesFolder", "Files.Read.All"],
    ] as const) {
      const e = catalog.find((x) => x.toolName === name);
      expect(e?.method, name).toBe("get");
      expect(e?.pathPattern, name).toBe(`/teams/{team-id}/channels/{channel-id}/${tail}`);
      expect(e?.["workScopes"], name).toEqual([scope]);
    }
  });

  it("describes the typed contracts, not raw Graph syntax (D29)", async () => {
    const tools = (await mcp.listTools()).tools;
    const desc = (n: string) => tools.find((t) => t.name === n)?.description ?? "";
    const tip = (n: string) => String(catalog.find((x) => x.toolName === n)?.["llmTip"]);
    for (const n of [TABS, FOLDER]) {
      expect(desc(n)).toMatch(
        /team-id is a GUID; channel-id is an opaque Teams channel\/thread identifier/,
      );
      expect(desc(n)).toMatch(/work\/school accounts only/);
      expect(desc(n) + tip(n)).not.toMatch(
        /list-folder-files|downloadUrl|contentUrl|websiteUrl|removeUrl|entityId|ChannelSettings/,
      );
    }
    expect(desc(TABS)).toMatch(
      /top is a local result limit \(1-200, default 50\): no Graph \$top is sent/,
    );
    expect(desc(TABS)).toMatch(/only typed expansion is teamsApp/);
    expect(desc(TABS)).toMatch(/configuration object .* is deliberately not exposed/);
    expect(desc(TABS)).toMatch(/does not return the native Files tab/);
    expect(desc(TABS)).toMatch(/Requires TeamsTab\.Read\.All/);
    expect(desc(FOLDER)).toMatch(/sends no OData parameters/);
    expect(desc(FOLDER)).toMatch(/does not list the files in the folder/);
    expect(desc(FOLDER)).toMatch(/file contents and downloads are outside this capability/);
    expect(desc(FOLDER)).toMatch(/migrated private channels/);
    expect(desc(FOLDER)).toMatch(/Requires Files\.Read\.All/);
  });

  it("exposes only the typed options", async () => {
    const tools = (await mcp.listTools()).tools;
    const props = (n: string) =>
      Object.keys(
        (tools.find((t) => t.name === n)?.inputSchema.properties ?? {}) as Record<string, unknown>,
      ).sort();
    expect(props(TABS)).toEqual(["channel-id", "expand", "team-id", "top"]);
    expect(props(FOLDER)).toEqual(["channel-id", "team-id"]);
  });
});

describe(TABS, () => {
  it("requests the tabs path with no query by default and projects to the allowlist", async () => {
    responses.push(() =>
      json({ "@odata.context": "ctx-secret", value: [rawTab({ teamsApp: undefined })] }),
    );
    const { env, text } = await call(TABS, IDS);
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(calls[0]?.method).toBe("GET");
    expect(decodeURIComponent(parsed().pathname)).toBe(`/v1.0/teams/${TID}/channels/${CID}/tabs`);
    expect(parsed().search).toBe(""); // no $top, $select or $expand sent
    expect(env.data).toEqual([projectedTab(false)]);
    expect(env.meta).toMatchObject({ count: 1, hasMore: false });
    expect(text).not.toMatch(/secret|contoso\.com|configuration/);
  });

  it("sends $expand=teamsApp only when asked and projects teamsApp to three scalars", async () => {
    responses.push(() => json({ value: [rawTab()] }));
    const { env, text } = await call(TABS, { ...IDS, expand: ["teamsApp"] });
    expect(parsed().searchParams.get("$expand")).toBe("teamsApp");
    expect(parsed().searchParams.get("$select")).toBeNull();
    expect(env.data).toEqual([projectedTab(true)]);
    expect(text).not.toMatch(/secret|appDefinitions|configuration/);
  });

  it("drops configuration and an unrequested teamsApp even when Graph over-returns them", async () => {
    responses.push(() =>
      json({
        value: [rawTab(), rawTab({ id: "b", configuration: null, teamsApp: null })],
      }),
    );
    const { env, text } = await call(TABS, IDS);
    expect((env.data as unknown[])[0]).toEqual(projectedTab(false));
    expect((env.data as Array<Record<string, unknown>>)[1]).not.toHaveProperty("teamsApp");
    expect(text).not.toMatch(/secret|configuration|removeUrl|contentUrl|websiteUrl|Contoso"/);
  });

  it("keeps teamsApp only when expanded, and tolerates a null teamsApp/configuration", async () => {
    responses.push(() =>
      json({
        value: [rawTab(), rawTab({ id: "b", configuration: null, teamsApp: null })],
      }),
    );
    const { env, text } = await call(TABS, { ...IDS, expand: ["teamsApp"] });
    expect((env.data as unknown[])[0]).toEqual(projectedTab(true));
    expect((env.data as Array<Record<string, unknown>>)[1]).toMatchObject({
      id: "b",
      teamsApp: null,
    });
    expect(text).not.toMatch(/secret|configuration|removeUrl|contentUrl|websiteUrl/);
  });

  it("sanitizeTab drops objects where scalars are expected and non-objects", () => {
    expect(sanitizeTab("x")).toBeUndefined();
    expect(sanitizeTab({ id: { a: 1 }, displayName: ["x"], webUrl: "u" })).toEqual({ webUrl: "u" });
  });

  it("an empty tab list is a successful empty list", async () => {
    responses.push(() => json({ value: [] }));
    const { env } = await call(TABS, IDS);
    expect(env.ok).toBe(true);
    expect(env.data).toEqual([]);
    expect(env.meta).toMatchObject({ count: 0, hasMore: false });
  });

  it("top is a local limit: no $top on the wire, extra items clipped", async () => {
    responses.push(() =>
      json({ value: [rawTab({ id: "a" }), rawTab({ id: "b" }), rawTab({ id: "c" })] }),
    );
    const { env } = await call(TABS, { ...IDS, top: 2 });
    expect(parsed().searchParams.get("$top")).toBeNull();
    expect((env.data as unknown[]).length).toBe(2);
  });

  it("rejects options outside the typed contract with zero requests", async () => {
    for (const extra of [
      { top: 0 },
      { top: 201 },
      { top: 1.5 },
      { expand: [] },
      { expand: ["configuration"] },
      { expand: ["teamsApp", "teamsApp.appDefinitions"] },
      { select: ["configuration"] },
      { filter: "teamsApp/id eq 'x'" },
      { orderby: "displayName" },
      { skip: 1 },
      { search: "x" },
      { $expand: "teamsApp" },
      { headers: { ConsistencyLevel: "eventual" } },
    ]) {
      calls = [];
      const { env } = await call(TABS, { ...IDS, ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
    }
  });
});

describe(FOLDER, () => {
  it("requests the filesFolder path with no query and projects folder metadata only", async () => {
    responses.push(() => json(rawFolder()));
    const { env, text } = await call(FOLDER, IDS);
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    expect(decodeURIComponent(parsed().pathname)).toBe(
      `/v1.0/teams/${TID}/channels/${CID}/filesFolder`,
    );
    expect(parsed().search).toBe("");
    expect(env.data).toEqual(projectedFolder);
    expect(text).not.toMatch(
      /secret|downloadUrl|createdBy|lastModifiedBy|fileSystemInfo|eTag|cTag|sharepointIds|siteId|viewType/,
    );
  });

  it("keeps a minimal driveItem (no folder facet, no parentReference.id)", async () => {
    responses.push(() =>
      json({ id: "x", name: "n", parentReference: { driveId: "d" }, folder: null }),
    );
    const { env } = await call(FOLDER, IDS);
    expect(env.data).toEqual({ id: "x", name: "n", parentReference: { driveId: "d" } });
  });

  it("fails closed on a non-object 200 body", async () => {
    for (const body of [[], "text", null, 5]) {
      responses.push(() => json(body));
      const { env } = await call(FOLDER, IDS);
      expect(env.ok, JSON.stringify(body)).toBe(false);
      expect(env.error?.code).toBe("graph_error");
    }
  });

  it("sanitizeFilesFolder drops objects where scalars are expected", () => {
    expect(sanitizeFilesFolder(null)).toBeUndefined();
    expect(
      sanitizeFilesFolder({ id: "x", name: { a: 1 }, size: "7", folder: "x", parentReference: 1 }),
    ).toEqual({ id: "x", size: "7" });
  });

  it("rejects any option with zero requests (the page documents no OData)", async () => {
    for (const extra of [
      { top: 1 },
      { select: ["id"] },
      { expand: ["children"] },
      { filter: "x" },
    ]) {
      calls = [];
      const { env } = await call(FOLDER, { ...IDS, ...extra });
      expect(env.ok, JSON.stringify(extra)).toBe(false);
      expect(calls, JSON.stringify(extra)).toHaveLength(0);
    }
  });
});

describe("path id injection (zero network calls)", () => {
  const badTeam = [
    "",
    "not-a-guid",
    `${TID}/channels`,
    `${TID}%2Fchannels`,
    "..",
    "%2e%2e",
    `${TID}?$top=1`,
    `${TID}#x`,
    ` ${TID}`,
    `${TID}\n`,
    "me",
  ];
  const badChannel = [
    "",
    `${CID}/tabs`,
    `${CID}/filesFolder`,
    "19%3Aabc%40thread.tacv2%2Ftabs",
    `${CID}%2F..%2F..`,
    "..",
    "$ref",
    "getAllMessages",
    `${CID}?$expand=teamsApp`,
    `${CID}#frag`,
    ` ${CID}`,
    `19:${"a".repeat(600)}@thread.tacv2`,
    TID,
  ];
  it.each([TABS, FOLDER])("%s rejects malformed team and channel ids", async (name) => {
    for (const t of badTeam) {
      calls = [];
      const { env } = await call(name, { "team-id": t, "channel-id": CID });
      expect(env.ok, `team ${JSON.stringify(t)}`).toBe(false);
      expect(calls, `team ${JSON.stringify(t)}`).toHaveLength(0);
    }
    for (const c of badChannel) {
      calls = [];
      const { env } = await call(name, { "team-id": TID, "channel-id": c });
      expect(env.ok, `channel ${JSON.stringify(c)}`).toBe(false);
      expect(calls, `channel ${JSON.stringify(c)}`).toHaveLength(0);
    }
  });
});

describe("tabs pagination fails closed (same team, channel and collection)", () => {
  const own = `/teams/${TID}/channels/${CID}/tabs`;
  const next = (path: string) => `https://graph.microsoft.com/v1.0${path}?$skiptoken=1`;

  it("refuses a nextLink for another team, channel, collection or version", async () => {
    const cidPart = (cid: string) => `/teams/${TID}/channels/${cid}`;
    const bad = [
      own.replace(TID, OTHER_TID),
      own.replace(CID, OTHER_CID),
      `${cidPart(CID)}/filesFolder`,
      `${cidPart(CID)}/messages`,
      `${cidPart(CID)}/tabs/${TID}`,
      `${cidPart(CID)}`,
      `/teams/${TID}/channels`,
      `/chats/${CID}/tabs`,
      `${own}/$value`,
      `${own}%2f..%2f${enc(OTHER_CID)}`,
    ];
    for (const path of bad) {
      calls = [];
      responses.push(() => json({ value: [rawTab()], "@odata.nextLink": next(path) }));
      const { env } = await call(TABS, { ...IDS, top: 5 });
      expect(env.ok, path).toBe(false);
      expect(env.error?.code, path).toBe("invalid_input");
      expect(calls, path).toHaveLength(1);
    }
    for (const link of [
      `https://evil.example/v1.0${own}?$skiptoken=1`,
      `https://graph.microsoft.com.evil.example/v1.0${own}`,
      `https://graph.microsoft.com/beta${own}?$skiptoken=1`,
      `http://graph.microsoft.com/v1.0${own}?$skiptoken=1`,
      `https://graph.microsoft.com:444/v1.0${own}?$skiptoken=1`,
    ]) {
      calls = [];
      responses.push(() => json({ value: [rawTab()], "@odata.nextLink": link }));
      const { env } = await call(TABS, { ...IDS, top: 5 });
      expect(env.error?.code, link).toBe("invalid_input");
      expect(calls, link).toHaveLength(1);
    }
  });

  it("follows the documented same-scope nextLink forms and projects every page", async () => {
    for (const link of [next(own), next(own.replace(CID, ECID))]) {
      calls = [];
      responses.push(() => json({ value: [rawTab()], "@odata.nextLink": link }));
      responses.push(() => json({ value: [rawTab({ id: "p2" })] }));
      const { env, text } = await call(TABS, { ...IDS, top: 5 });
      expect(env.ok, link).toBe(true);
      expect(calls).toHaveLength(2);
      expect((env.data as unknown[]).length).toBe(2);
      expect(text).not.toMatch(/secret|configuration/);
    }
  });

  it("fails closed on a 200 page without a value array", async () => {
    for (const body of [{ error: "backend failure" }, { value: null }, { value: {} }]) {
      responses.push(() => json(body));
      expect((await call(TABS, IDS)).env.error?.code, JSON.stringify(body)).toBe("graph_error");
    }
  });
});

describe("errors", () => {
  it.each([
    [TABS, IDS],
    [FOLDER, IDS],
  ] as const)("%s: 400 and 404 are graph_error, never empty success", async (name, args) => {
    for (const status of [400, 404]) {
      calls = [];
      responses.push(() => json({ error: { code: "NotFound", message: "m" } }, status));
      const { env } = await call(name, args);
      expect(env.ok, String(status)).toBe(false);
      expect(env.data).toBeUndefined();
      expect(env.error?.code, String(status)).toBe("graph_error");
      expect(calls).toHaveLength(1);
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

  it("tabs: batch projects exactly like the single path, with and without expand", async () => {
    responses.push(() => json({ value: [rawTab()] }));
    const single = (await call(TABS, { ...IDS, expand: ["teamsApp"] })).env.data;
    responses.push(ok({ value: [rawTab()] }));
    const b = await batch(`/teams/${TID}/channels/${ECID}/tabs?$expand=teamsApp`);
    expect(b.env.ok).toBe(true);
    expect(bodyOf(b)?.["value"]).toEqual(single);
    expect(b.text).not.toMatch(/secret|configuration/);

    responses.push(ok({ value: [rawTab()] }));
    const plain = await batch(`/teams/${TID}/channels/${CID}/tabs`);
    // An over-returned teamsApp is dropped in batch too when $expand did not ask for it.
    expect(bodyOf(plain)?.["value"]).toEqual([projectedTab(false)]);
    expect(plain.text).not.toMatch(/secret|configuration/);
  });

  it("files folder: batch projects exactly like the single path", async () => {
    responses.push(() => json(rawFolder()));
    const single = (await call(FOLDER, IDS)).env.data;
    responses.push(ok(rawFolder()));
    const b = await batch(`/teams/${TID}/channels/${ECID}/filesFolder`);
    expect(b.env.ok).toBe(true);
    expect(bodyOf(b)).toEqual(single);
    expect(bodyOf(b)).toEqual(projectedFolder);
    expect(b.text).not.toMatch(/secret|downloadUrl|createdBy/);
  });

  it("rejects out-of-policy subrequests with zero network calls", async () => {
    const c = `/teams/${TID}/channels/${ECID}`;
    const bad: Array<[string, Record<string, string>?]> = [
      [`${c}/tabs?$top=5`],
      [`${c}/tabs?$select=configuration`],
      [`${c}/tabs?$expand=configuration`],
      [`${c}/tabs?$expand=teamsApp,teamsApp`],
      [`${c}/tabs?$expand=teamsApp&$expand=teamsApp`],
      [`${c}/tabs?$filter=${enc("teamsApp/id eq 'x'")}`],
      [`${c}/tabs?$count=true`],
      [`${c}/tabs?$skip=1`],
      [`${c}/tabs/abc`],
      [`${c}/tabs/abc?$expand=teamsApp`],
      [`${c}/tabs`, { ConsistencyLevel: "eventual" }],
      [`${c}/tabs`, { Prefer: 'outlook.body-content-type="html"' }],
      [`${c}/filesFolder?$select=id`],
      [`${c}/filesFolder?$expand=children`],
      [`${c}/filesFolder/children`],
      [`${c}/filesFolder/content`],
      [`${c}/filesFolder`, { ConsistencyLevel: "eventual" }],
      [`/teams/${TID}/channels/${enc(`${CID}/x`)}/tabs`],
      [`/teams/not-a-guid/channels/${ECID}/tabs`],
      [`/groups/${TID}/team/channels/${ECID}/tabs`],
      [`/beta/teams/${TID}/channels/${ECID}/tabs`],
      [`/beta/teams/${TID}/channels/${ECID}/filesFolder`],
    ];
    for (const [url, headers] of bad) {
      calls = [];
      const { env } = await batch(url, headers);
      expect(env.ok, url).toBe(false);
      expect(calls, url).toHaveLength(0);
    }
  });
});
