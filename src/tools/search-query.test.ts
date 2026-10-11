import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { GraphClient, type FetchLike } from "../graph/client.js";
import { createServer } from "../server.js";
import { createHandlers } from "./index.js";
import {
  MAX_FROM,
  MAX_QUERY_STRING,
  MAX_SUMMARY,
  entityTypeCombinationIssue,
  maxSizeFor,
  projectSummary,
} from "./search-query.js";

/**
 * SM-14 part B behavior tests for `search-query` (D38): exact request body, local validation
 * (combination, size per family, strict unknown properties), response projection (resource
 * allowlist, summary markup/bound), fail-closed bodies, and batch unreachability.
 * The permission/capability matrix lives in search-query-capability.test.ts.
 */

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";

interface Call {
  url: string;
  method: string;
  headers: Record<string, string>;
  body: unknown;
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
  data?: { searchTerms: string[]; hitsContainers: Array<Record<string, unknown>> };
  meta?: Record<string, unknown>;
  error?: { code: string; message: string; details?: Record<string, unknown> };
};
const call = async (name: string, args: Record<string, unknown> = {}) => {
  const r = await mcp.callTool({ name, arguments: args });
  return { r, env: r.structuredContent as Env, text: JSON.stringify(r) };
};
const search = (args: Record<string, unknown>) => call("search-query", args);

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  const fetch: FetchLike = (url, init) => {
    calls.push({
      url,
      method: String(init.method),
      headers: { ...(init.headers as Record<string, string>) },
      body: typeof init.body === "string" ? JSON.parse(init.body) : init.body,
    });
    const next = responses.shift();
    return Promise.resolve(
      next ? next() : json({ value: [{ searchTerms: [], hitsContainers: [] }] }),
    );
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

const hit = (resource: unknown, extra: Record<string, unknown> = {}) => ({
  hitId: "h1",
  rank: 1,
  summary: "s",
  resource,
  ...extra,
});
const reply = (hits: unknown[], extra: Record<string, unknown> = {}) =>
  json({
    value: [
      {
        searchTerms: ["contoso"],
        hitsContainers: [{ total: hits.length, moreResultsAvailable: false, hits, ...extra }],
      },
    ],
  });
const firstHit = (env: Env) =>
  (env.data?.hitsContainers[0]?.["hits"] as Array<Record<string, unknown>>)[0];

describe("search-query request construction", () => {
  it("sends exactly POST /v1.0/search/query with the constructed single-request body", async () => {
    responses.push(() => reply([]));
    const { env } = await search({ entityTypes: ["message"], queryString: "contoso" });
    expect(env.ok).toBe(true);
    expect(calls).toHaveLength(1);
    const c = calls[0] as Call;
    expect(c.method).toBe("POST");
    const url = new URL(c.url);
    expect(url.origin + url.pathname).toBe("https://graph.microsoft.com/v1.0/search/query");
    expect(url.search).toBe("");
    expect(c.body).toStrictEqual({
      requests: [
        { entityTypes: ["message"], query: { queryString: "contoso" }, from: 0, size: 25 },
      ],
    });
    expect(c.headers["Prefer"]).toBeUndefined();
    expect(c.headers["ConsistencyLevel"]).toBeUndefined();
  });

  it("passes typed from/size and a SharePoint combination through unchanged", async () => {
    responses.push(() => reply([]));
    await search({
      entityTypes: ["driveItem", "listItem", "site"],
      queryString: "a b filetype:docx",
      from: 50,
      size: 50,
    });
    expect(calls[0]?.body).toStrictEqual({
      requests: [
        {
          entityTypes: ["driveItem", "listItem", "site"],
          query: { queryString: "a b filetype:docx" },
          from: 50,
          size: 50,
        },
      ],
    });
  });
});

describe("search-query local validation (zero Graph requests)", () => {
  it.each([
    ["message", "event"],
    ["message", "chatMessage"],
    ["event", "chatMessage"],
    ["message", "driveItem"],
    ["chatMessage", "site"],
    ["event", "list", "listItem"],
  ])("combination: %j is rejected", async (...types) => {
    const { env } = await search({ entityTypes: types, queryString: "x" });
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("combination: every non-empty SharePoint/OneDrive subset is allowed, each solo type alone", () => {
    const sp = ["drive", "driveItem", "list", "listItem", "site"] as const;
    for (let mask = 1; mask < 1 << sp.length; mask++) {
      const subset = sp.filter((_, i) => (mask & (1 << i)) !== 0);
      expect(entityTypeCombinationIssue(subset), subset.join()).toBeUndefined();
    }
    for (const solo of ["message", "event", "chatMessage"] as const) {
      expect(entityTypeCombinationIssue([solo])).toBeUndefined();
      expect(entityTypeCombinationIssue([solo, "site"])).toMatch(/alone/);
    }
  });

  it("entityTypes: empty, duplicate, unknown (person, externalItem) and non-array are rejected", async () => {
    for (const entityTypes of [
      [],
      ["site", "site"],
      ["person"],
      ["externalItem"],
      ["acronym"],
      "message",
    ]) {
      const { env } = await search({ entityTypes, queryString: "x" });
      expect(env.error?.code, JSON.stringify(entityTypes)).toBe("invalid_input");
    }
    expect((await search({ queryString: "x" })).env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("size per family: 25 cap for message/event, 50 for the rest, 1 minimum", async () => {
    expect(maxSizeFor(["message"])).toBe(25);
    expect(maxSizeFor(["event"])).toBe(25);
    expect(maxSizeFor(["chatMessage"])).toBe(50);
    expect(maxSizeFor(["drive", "site"])).toBe(50);
    for (const entityTypes of [["message"], ["event"]]) {
      expect((await search({ entityTypes, queryString: "x", size: 26 })).env.error?.code).toBe(
        "invalid_input",
      );
    }
    for (const size of [0, -1, 51, 1.5, 500]) {
      expect(
        (await search({ entityTypes: ["site"], queryString: "x", size })).env.error?.code,
        String(size),
      ).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
    responses.push(() => reply([]));
    expect(
      (await search({ entityTypes: ["chatMessage"], queryString: "x", size: 50 })).env.ok,
    ).toBe(true);
    responses.push(() => reply([]));
    expect((await search({ entityTypes: ["message"], queryString: "x", size: 25 })).env.ok).toBe(
      true,
    );
  });

  it("queryString: blank, missing, over the project limit, non-string; limit pinned at 1000", async () => {
    expect(MAX_QUERY_STRING).toBe(1000);
    for (const queryString of ["", "   ", "\n\t", "x".repeat(1001), 5, null]) {
      const { env } = await search({ entityTypes: ["message"], queryString });
      expect(env.error?.code, String(queryString).slice(0, 8)).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
    responses.push(() => reply([]));
    expect((await search({ entityTypes: ["message"], queryString: "x".repeat(1000) })).env.ok).toBe(
      true,
    );
  });

  it("from: negative, fractional and over the project maximum (pinned at 1000) are rejected", async () => {
    expect(MAX_FROM).toBe(1000);
    for (const from of [-1, 1.5, 1001, "0"]) {
      const { env } = await search({ entityTypes: ["site"], queryString: "x", from });
      expect(env.error?.code, String(from)).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it.each([
    "fields",
    "aggregations",
    "aggregationFilters",
    "collapseProperties",
    "sortProperties",
    "queryTemplate",
    "contentSources",
    "region",
    "enableTopResults",
    "queryAlterationOptions",
    "resultTemplateOptions",
    "sharePointOneDriveOptions",
    "requests",
    "query",
    "headers",
    "body",
    "url",
    "version",
  ])("unknown input property %s is rejected before any request", async (key) => {
    const { env, text } = await search({
      entityTypes: ["driveItem"],
      queryString: "x",
      [key]: "x",
    });
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
    expect(text).not.toContain(TOKEN);
  });
});

describe("search-query summary", () => {
  it("strips Search highlight markup and keeps ordinary text", () => {
    expect(projectSummary("<c0>Contoso</c0> Detailed Design <ddd/>")).toBe(
      "Contoso Detailed Design ",
    );
    expect(projectSummary("a <c12>b</c12> <b>kept</b>")).toBe("a b <b>kept</b>");
    expect(projectSummary(null)).toBeUndefined();
  });

  it("bounds the summary length (pinned at 500) and marks truncation", () => {
    expect(MAX_SUMMARY).toBe(500);
    const out = projectSummary("x".repeat(5000)) as string;
    expect(out).toBe(`${"x".repeat(499)}…`);
    expect(out.length).toBe(500);
    expect(projectSummary("z".repeat(501))).toBe(`${"z".repeat(499)}…`);
    expect(projectSummary("y".repeat(500))).toBe("y".repeat(500));
  });

  it("a non-string summary fails the call closed", async () => {
    responses.push(() => reply([hit(undefined, { summary: { evil: true } })]));
    const { env } = await search({ entityTypes: ["site"], queryString: "x" });
    expect(env.ok).toBe(false);
    expect(env.error?.code).toBe("graph_error");
  });
});

const FULL: Record<string, Record<string, unknown>> = {
  message: {
    "@odata.type": "#microsoft.graph.message",
    id: "m1",
    subject: "Weekly digest",
    bodyPreview: "BODY_PREVIEW_SECRET",
    body: { content: "BODY_SECRET" },
    receivedDateTime: "2019-10-07T10:00:09Z",
    sentDateTime: "2019-10-07T09:59:52Z",
    hasAttachments: false,
    webLink: "https://outlook.office365.com/owa/?ItemID=x",
    toRecipients: [{ emailAddress: { address: "TO_SECRET@contoso.com" } }],
    from: { emailAddress: { name: "Office365", address: "gt@contoso.com", extra: "EXTRA" } },
    sender: { emailAddress: { address: "SENDER_SECRET@contoso.com" } },
  },
  event: {
    "@odata.type": "#microsoft.graph.event",
    id: "e1",
    subject: "Sync",
    start: { dateTime: "2020-06-16T03:30:00Z", timeZone: "UTC" },
    end: { dateTime: "2020-06-16T04:15:00Z", timeZone: "UTC" },
    isAllDay: false,
    iCalUId: "ICAL_SECRET",
    location: { displayName: "LOCATION_SECRET" },
    organizer: { emailAddress: { name: "Org", address: "ORG_ADDR_SECRET@contoso.com" } },
    attendees: [{ emailAddress: { address: "ATT_SECRET@contoso.com" } }],
    body: { content: "BODY_SECRET" },
  },
  chatMessage: {
    "@odata.type": "microsoft.graph.chatMessage",
    id: "1657782060227",
    createdDateTime: "2022-07-14T07:01:01Z",
    lastModifiedDateTime: "2022-07-14T07:01:03Z",
    subject: "SUBJECT_SECRET",
    importance: "normal",
    etag: "ETAG_SECRET",
    webLink: "https://outlook.office365.com/owa/?ItemID=y",
    chatId: "19:abc@thread.v2",
    channelIdentity: { teamId: "t1", channelId: "c1", other: "OTHER_SECRET" },
    from: {
      emailAddress: { name: "Bob", address: "bob@contoso.com" },
      user: { displayName: "Bob U", id: "UID_SECRET" },
    },
  },
  driveItem: {
    "@odata.type": "#microsoft.graph.driveItem",
    id: "d1",
    name: "design.docx",
    webUrl: "https://contoso.sharepoint.com/sites/x/design.docx",
    size: 1234,
    createdDateTime: "2019-06-10T06:37:43Z",
    lastModifiedDateTime: "2019-06-11T06:37:43Z",
    createdBy: { user: { displayName: "CREATED_BY_SECRET" } },
    "@microsoft.graph.downloadUrl": "https://download.example/DOWNLOAD_SECRET",
    parentReference: {
      driveId: "drv1",
      siteId: "site1",
      path: "PATH_SECRET",
      sharepointIds: { listId: "LIST_SECRET" },
    },
    listItem: { fields: { title: "FIELDS_SECRET" } },
  },
  listItem: {
    "@odata.type": "#microsoft.graph.listItem",
    id: "li1",
    webUrl: "https://contoso.sharepoint.com/Lists/a/DispForm.aspx?ID=1",
    createdDateTime: "2019-06-10T06:37:43Z",
    lastModifiedDateTime: "2019-06-11T06:37:43Z",
    fields: { title: "FIELDS_SECRET" },
    parentReference: { siteId: "site1", path: "PATH_SECRET" },
  },
  site: {
    "@odata.type": "#microsoft.graph.site",
    id: "contoso.sharepoint.com,a,b",
    displayName: "Contoso",
    name: "contoso",
    webUrl: "https://contoso.sharepoint.com/sites/contoso",
    description: "DESCRIPTION_SECRET",
  },
  drive: {
    "@odata.type": "#microsoft.graph.drive",
    id: "drv1",
    name: "Documents",
    webUrl: "https://contoso.sharepoint.com/Documents",
    driveType: "documentLibrary",
    owner: { user: { displayName: "OWNER_SECRET" } },
  },
  list: {
    "@odata.type": "#microsoft.graph.list",
    id: "l1",
    displayName: "Tasks",
    name: "Tasks",
    webUrl: "https://contoso.sharepoint.com/Lists/Tasks",
    description: "DESCRIPTION_SECRET",
  },
};
const EXPECTED: Record<string, Record<string, unknown>> = {
  message: {
    id: "m1",
    subject: "Weekly digest",
    receivedDateTime: "2019-10-07T10:00:09Z",
    sentDateTime: "2019-10-07T09:59:52Z",
    hasAttachments: false,
    webLink: "https://outlook.office365.com/owa/?ItemID=x",
    from: { emailAddress: { name: "Office365", address: "gt@contoso.com" } },
  },
  event: {
    id: "e1",
    subject: "Sync",
    isAllDay: false,
    start: { dateTime: "2020-06-16T03:30:00Z", timeZone: "UTC" },
    end: { dateTime: "2020-06-16T04:15:00Z", timeZone: "UTC" },
    organizer: { emailAddress: { name: "Org" } },
  },
  chatMessage: {
    id: "1657782060227",
    createdDateTime: "2022-07-14T07:01:01Z",
    lastModifiedDateTime: "2022-07-14T07:01:03Z",
    chatId: "19:abc@thread.v2",
    webLink: "https://outlook.office365.com/owa/?ItemID=y",
    from: {
      emailAddress: { name: "Bob", address: "bob@contoso.com" },
      user: { displayName: "Bob U" },
    },
    channelIdentity: { teamId: "t1", channelId: "c1" },
  },
  driveItem: {
    id: "d1",
    name: "design.docx",
    webUrl: "https://contoso.sharepoint.com/sites/x/design.docx",
    size: 1234,
    createdDateTime: "2019-06-10T06:37:43Z",
    lastModifiedDateTime: "2019-06-11T06:37:43Z",
    parentReference: { driveId: "drv1", siteId: "site1" },
  },
  listItem: {
    id: "li1",
    webUrl: "https://contoso.sharepoint.com/Lists/a/DispForm.aspx?ID=1",
    createdDateTime: "2019-06-10T06:37:43Z",
    lastModifiedDateTime: "2019-06-11T06:37:43Z",
    parentReference: { siteId: "site1" },
  },
  site: {
    id: "contoso.sharepoint.com,a,b",
    displayName: "Contoso",
    name: "contoso",
    webUrl: "https://contoso.sharepoint.com/sites/contoso",
  },
  drive: {
    id: "drv1",
    name: "Documents",
    webUrl: "https://contoso.sharepoint.com/Documents",
    driveType: "documentLibrary",
  },
  list: {
    id: "l1",
    displayName: "Tasks",
    name: "Tasks",
    webUrl: "https://contoso.sharepoint.com/Lists/Tasks",
  },
};

describe("search-query resource allowlist projection", () => {
  it.each(Object.keys(FULL))(
    "%s: only allowlisted fields survive, nothing else leaks",
    async (type) => {
      responses.push(() =>
        reply([
          hit(FULL[type], {
            summary: "<c0>hi</c0> there <ddd/>",
            extraHitProp: "HIT_EXTRA_SECRET",
          }),
        ]),
      );
      const { env, text } = await search({ entityTypes: [type], queryString: "x" });
      expect(env.ok).toBe(true);
      expect(firstHit(env)).toStrictEqual({
        hitId: "h1",
        rank: 1,
        summary: "hi there ",
        resourceType: type,
        resource: EXPECTED[type],
      });
      expect(text).not.toMatch(/SECRET/);
      expect(text).not.toContain("@odata.type");
      expect(text).not.toContain(TOKEN);
    },
  );

  it("drops resources of unknown, malformed or non-requested types but keeps the envelope", async () => {
    const hits = [
      hit({ "@odata.type": "#microsoft.graph.person", id: "p", displayName: "PERSON_SECRET" }),
      hit({
        "@odata.type": "#microsoft.graph.externalItem",
        id: "x",
        properties: { a: "EXT_SECRET" },
      }),
      hit({ id: "no-type", subject: "NOTYPE_SECRET" }),
      hit("a string resource"),
      hit(null),
      hit(undefined),
      hit(FULL["message"]), // not requested: a driveItem search must not surface a message
      hit({ "@odata.type": "#microsoft.graph.driveItemVersion", id: "v", name: "VERSION_SECRET" }),
      hit({ "@odata.type": "evil.graph.driveItem", id: "e", name: "EVIL_SECRET" }),
    ];
    responses.push(() => reply(hits));
    const { env, text } = await search({ entityTypes: ["driveItem"], queryString: "x", size: 50 });
    expect(env.ok).toBe(true);
    const out = env.data?.hitsContainers[0]?.["hits"] as Array<Record<string, unknown>>;
    expect(out).toHaveLength(hits.length);
    for (const h of out) {
      expect(h).toStrictEqual({ hitId: "h1", rank: 1, summary: "s" });
    }
    expect(text).not.toMatch(/SECRET/);
  });

  it("a SharePoint-family request accepts any SharePoint-family resource (files are also listItems)", async () => {
    responses.push(() => reply([hit(FULL["listItem"]), hit(FULL["driveItem"])]));
    const { env } = await search({ entityTypes: ["listItem"], queryString: "x" });
    const out = env.data?.hitsContainers[0]?.["hits"] as Array<Record<string, unknown>>;
    expect(out.map((h) => h["resourceType"])).toEqual(["listItem", "driveItem"]);
  });

  it("projects only searchTerms and per-container total/moreResultsAvailable/hits", async () => {
    responses.push(() =>
      json({
        "@odata.context": "CONTEXT_SECRET",
        value: [
          {
            searchTerms: ["contoso"],
            queryAlterationResponse: { originalQueryString: "ALTER_SECRET" },
            resultTemplates: { a: "TEMPLATE_SECRET" },
            hitsContainers: [
              {
                total: 7,
                moreResultsAvailable: true,
                aggregations: [{ field: "AGG_SECRET" }],
                hits: [
                  hit(FULL["site"], { resultTemplateId: "RT_SECRET", contentSource: "CS_SECRET" }),
                ],
              },
            ],
          },
        ],
      }),
    );
    const { env, text } = await search({ entityTypes: ["site"], queryString: "contoso" });
    expect(env.data).toStrictEqual({
      searchTerms: ["contoso"],
      hitsContainers: [
        {
          total: 7,
          moreResultsAvailable: true,
          hits: [
            {
              hitId: "h1",
              rank: 1,
              summary: "s",
              resourceType: "site",
              resource: EXPECTED["site"],
            },
          ],
        },
      ],
    });
    expect(text).not.toMatch(/SECRET/);
  });

  it("a present-but-empty hitsContainers list is accepted (deliberate; live checklist item 8)", async () => {
    responses.push(() => json({ value: [{ searchTerms: [], hitsContainers: [] }] }));
    const { env } = await search({ entityTypes: ["driveItem"], queryString: "nothing" });
    expect(env.ok).toBe(true);
    expect(env.data?.hitsContainers).toStrictEqual([]);
  });

  it("a genuine empty result (one container, empty hits) is a successful empty projection", async () => {
    responses.push(() =>
      json({
        value: [
          {
            searchTerms: [],
            hitsContainers: [{ total: 0, moreResultsAvailable: false, hits: [] }],
          },
        ],
      }),
    );
    const { env } = await search({ entityTypes: ["message"], queryString: "nothing" });
    expect(env.ok).toBe(true);
    expect(env.data?.hitsContainers[0]).toStrictEqual({
      total: 0,
      moreResultsAvailable: false,
      hits: [],
    });
  });

  it("an explicit empty container without hits (total 0, no more results) is an empty result", async () => {
    responses.push(() =>
      json({
        value: [
          { searchTerms: ["x"], hitsContainers: [{ total: 0, moreResultsAvailable: false }] },
        ],
      }),
    );
    const { env } = await search({ entityTypes: ["driveItem"], queryString: "nothing" });
    expect(env.ok).toBe(true);
    expect(env.data?.hitsContainers[0]).toStrictEqual({
      total: 0,
      moreResultsAvailable: false,
      hits: [],
    });
  });
});

describe("search-query fails closed on malformed bodies", () => {
  const bad: Array<[string, unknown]> = [
    ["array body", []],
    ["no value", {}],
    ["value not an array", { value: "x" }],
    ["empty value", { value: [] }],
    ["two search responses", { value: [{ hitsContainers: [] }, { hitsContainers: [] }] }],
    ["value item not an object", { value: [1] }],
    ["searchTerms not an array", { value: [{ searchTerms: "x", hitsContainers: [] }] }],
    ["searchTerms with a non-string", { value: [{ searchTerms: [1], hitsContainers: [] }] }],
    ["too many searchTerms", { value: [{ searchTerms: Array(51).fill("t"), hitsContainers: [] }] }],
    ["hitsContainers not an array", { value: [{ hitsContainers: {} }] }],
    ["hitsContainers absent", { value: [{ searchTerms: ["t"] }] }],
    ["empty searchResponse", { value: [{}] }],
    ["container without hits or total", { value: [{ hitsContainers: [{}] }] }],
    ["container without hits but total > 0", { value: [{ hitsContainers: [{ total: 3 }] }] }],
    [
      "container without hits but more results",
      { value: [{ hitsContainers: [{ total: 0, moreResultsAvailable: true }] }] },
    ],
    ["container with hits: null", { value: [{ hitsContainers: [{ total: 0, hits: null }] }] }],
    ["too many containers", { value: [{ hitsContainers: Array(6).fill({ hits: [] }) }] }],
    ["container not an object", { value: [{ hitsContainers: ["x"] }] }],
    ["hits not an array", { value: [{ hitsContainers: [{ hits: {} }] }] }],
    ["hit not an object", { value: [{ hitsContainers: [{ hits: ["x"] }] }] }],
    ["hit without hitId", { value: [{ hitsContainers: [{ hits: [{ rank: 1 }] }] }] }],
    ["hitId not a string", { value: [{ hitsContainers: [{ hits: [{ hitId: 5 }] }] }] }],
    ["rank not a number", { value: [{ hitsContainers: [{ hits: [{ hitId: "a", rank: "1" }] }] }] }],
    ["total not a number", { value: [{ hitsContainers: [{ total: "7", hits: [] }] }] }],
    [
      "moreResultsAvailable not a boolean",
      { value: [{ hitsContainers: [{ moreResultsAvailable: "no", hits: [] }] }] },
    ],
    [
      "more hits than the requested size",
      { value: [{ hitsContainers: [{ hits: [hit(null), hit(null)] }] }] },
    ],
    [
      "oversized resource text",
      {
        value: [
          {
            hitsContainers: [
              { hits: [hit({ "@odata.type": "#microsoft.graph.site", id: "x".repeat(5000) })] },
            ],
          },
        ],
      },
    ],
  ];
  it.each(bad)("%s", async (_name, body) => {
    responses.push(() => json(body));
    const { env, text } = await search({ entityTypes: ["site"], queryString: "x", size: 1 });
    expect(env.ok).toBe(false);
    expect(env.data).toBeUndefined();
    expect(env.error?.code).toBe("graph_error");
    expect(calls).toHaveLength(1);
    expect(text).not.toContain(TOKEN);
  });
});

describe("search-query is not reachable through graph-batch (batch is GET-only)", () => {
  it("rejects a POST and a GET subrequest to /search/query with zero Graph requests", async () => {
    for (const method of ["POST", "GET"]) {
      const { env, text } = await call("graph-batch", {
        requests: [{ id: "1", method, url: "/search/query" }],
      });
      expect(env.ok, method).toBe(false);
      expect(calls, method).toHaveLength(0);
      expect(text).not.toContain(TOKEN);
    }
  });

  it("the handler declares no batchPolicy and the tool is a POST catalog entry", () => {
    const graph = new GraphClient({
      tokenProvider: { getAccessToken: () => Promise.resolve(TOKEN) },
      catalog,
    });
    const handlers = createHandlers(graph, catalog);
    expect(handlers.get("search-query")?.batchPolicy).toBeUndefined();
    expect(catalog.find((e) => e.toolName === "search-query")?.method).toBe("post");
  });
});
