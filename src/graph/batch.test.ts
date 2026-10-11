import { beforeAll, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog, type EndpointDefinition } from "../catalog.js";
import type { TokenProvider } from "../auth/token-provider.js";
import {
  batchTargets,
  encodeBatchBody,
  isValidatedBatch,
  normalizeBatchResponse,
  validateBatch,
  type BatchTargets,
} from "./batch.js";
import { GraphClient, type FetchLike, type GraphClientOptions } from "./client.js";
import { GraphClientError } from "./errors.js";
import { validateRelativeGraphUrl } from "./url.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const G = "https://graph.microsoft.com";

let catalog: EndpointCatalog;
let permitted: EndpointDefinition[];
let targets: BatchTargets;
/** Every entry batchable with a permissive query policy (validator rules under test). */
const targetsOf = (entries: readonly EndpointDefinition[]): BatchTargets =>
  new Map(entries.map((e) => [e, { check: () => undefined }]));
beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  // Validator rules are exercised against every catalog GET entry.
  permitted = catalog.filter((e) => e.method === "get");
  targets = targetsOf(permitted);
});

describe("malformed escapes", () => {
  it("are reported as malformed percent-encoding, not as nested $batch", () => {
    expect(() => validateRelativeGraphUrl("/me/%zz", permitted, "v1.0")).toThrow(
      /malformed percent-encoding/,
    );
  });
});

describe("batchTargets", () => {
  it("admits only implemented GET tools that declare a batchPolicy", () => {
    const policy = { check: () => undefined };
    const handler = { handler: () => ({ content: [] }) };
    expect(batchTargets(catalog, new Map()).size).toBe(0);
    const got = batchTargets(
      catalog,
      new Map<string, { handler: unknown; batchPolicy?: typeof policy }>([
        ["list-mail-messages", { ...handler, batchPolicy: policy }],
        ["list-mail-folders", handler], // implemented, but no query contract: not batchable
        ["get-schedule", { ...handler, batchPolicy: policy }], // POST: never batchable
        ["graph-batch", { ...handler, batchPolicy: policy }],
      ]),
    );
    expect([...got.keys()].map((e) => e.toolName)).toEqual(["list-mail-messages"]);
    expect([...got.values()]).toEqual([policy]);
  });
});

describe("per-tool batch query policy", () => {
  const messages = () => permitted.filter((e) => e.pathPattern === "/me/messages");
  it("receives the decoded query pairs and its rejection fails validation", () => {
    const seen: Array<ReadonlyArray<readonly [string, string]>> = [];
    const strict: BatchTargets = new Map(
      messages().map((e) => [
        e,
        {
          check: (q: ReadonlyArray<readonly [string, string]>) => {
            seen.push(q);
            for (const [name, value] of q) {
              if (name !== "$top" || Number(value) > 10) {
                throw new GraphClientError("invalid_input", "query not allowed");
              }
            }
          },
        },
      ]),
    );
    expect(() =>
      validateBatch({ requests: [req("/me/messages?$top=5")] }, strict, "v1.0"),
    ).not.toThrow();
    expect(seen[0]).toEqual([["$top", "5"]]);
    for (const url of ["/me/messages?$top=999", "/me/messages?$expand=attachments"]) {
      expect(() => validateBatch({ requests: [req(url)] }, strict, "v1.0"), url).toThrow(
        /rejected by endpoint policy/,
      );
    }
  });
});

describe("batch policy headers and projector", () => {
  it("passes canonical subrequest headers to check and freezes the projector onto the request", () => {
    const seen: Array<Readonly<Record<string, string>>> = [];
    const project = (b: unknown) => ({ projected: b });
    const t: BatchTargets = new Map(
      permitted
        .filter((e) => e.pathPattern === "/me/messages")
        .map((e) => [e, { check: (_q, h) => void seen.push(h), project }]),
    );
    const b = validateBatch(
      { requests: [req("/me/messages", { headers: { prefer: "x" } })] },
      t,
      "v1.0",
    );
    expect(seen).toEqual([{ Prefer: "x" }]);
    expect(typeof b.requests[0]?.project).toBe("function");
    expect(
      normalizeBatchResponse({ responses: [{ id: "1", status: 200, body: { v: 1 } }] }, b, "t")
        .responses[0]?.body,
    ).toEqual({ projected: { v: 1 } });
  });
});

describe("validated batch immutability", () => {
  it("is deeply frozen, so validated urls/headers cannot be swapped before sending", () => {
    const b = validateBatch(
      { requests: [req("/me/messages", { headers: { Prefer: 'outlook.timezone="UTC"' } })] },
      targets,
      "v1.0",
    );
    const sub = b.requests[0] as unknown as Record<string, unknown>;
    expect(Object.isFrozen(b)).toBe(true);
    expect(Object.isFrozen(b.requests)).toBe(true);
    expect(Object.isFrozen(sub)).toBe(true);
    expect(Object.isFrozen(sub["headers"])).toBe(true);
    expect(() => {
      sub["url"] = "/users/someone-else/messages";
    }).toThrow(TypeError);
    expect(() => {
      (b.requests as unknown as unknown[]).push({ id: "x", url: "/x", headers: {} });
    }).toThrow(TypeError);
    expect(() => {
      (sub["headers"] as Record<string, string>)["Authorization"] = "Bearer x";
    }).toThrow(TypeError);
    expect(encodeBatchBody(b)).not.toMatch(/someone-else|Authorization/);
  });
});

const req = (url: string, extra: Record<string, unknown> = {}) => ({
  id: "1",
  method: "GET",
  url,
  ...extra,
});

function rejects(input: unknown): string {
  try {
    validateBatch(input, targets, "v1.0");
  } catch (e) {
    expect(e).toBeInstanceOf(GraphClientError);
    expect((e as GraphClientError).code).toBe("invalid_input");
    return (e as GraphClientError).message;
  }
  throw new Error("expected rejection");
}

function urlRejects(url: unknown): string {
  return rejects({ requests: [req(url as string)] });
}

describe("batch validator: request shape", () => {
  it("accepts a valid batch and produces an opaque validated value", () => {
    const b = validateBatch(
      {
        requests: [
          req("/me/messages?$top=5&$select=id,subject"),
          { id: "b-2", method: "GET", url: "/me/mailFolders('inbox')/messages/delta()" },
        ],
      },
      targets,
      "v1.0",
    );
    expect(isValidatedBatch(b)).toBe(true);
    expect(b.requests.map((r) => r.url)).toEqual([
      "/me/messages?$top=5&$select=id%2Csubject",
      "/me/mailFolders('inbox')/messages/delta()",
    ]);
    // Structurally identical forgeries are not accepted.
    expect(isValidatedBatch({ version: "v1.0", requests: b.requests })).toBe(false);
  });

  it("enforces 1..20 requests", () => {
    const mk = (n: number) => ({
      requests: Array.from({ length: n }, (_v, i) => ({ ...req("/me/messages"), id: `r${i}` })),
    });
    expect(() => validateBatch(mk(20), targets, "v1.0")).not.toThrow();
    rejects(mk(21));
    rejects(mk(0));
    rejects({ requests: "x" });
    rejects({});
    rejects(null);
    rejects([]);
  });

  it("rejects unknown top-level and subrequest fields, body, and dependsOn", () => {
    rejects({ requests: [req("/me/messages")], extra: 1 });
    rejects({ requests: [req("/me/messages", { body: {} })] });
    rejects({ requests: [req("/me/messages", { body: null })] });
    rejects({ requests: [req("/me/messages", { dependsOn: [] })] });
    rejects({ requests: [req("/me/messages", { atomicityGroup: "g" })] });
    rejects({ requests: ["/me/messages"] });
  });

  it("requires unique (case-insensitive), bounded, charset-checked string ids", () => {
    const withIds = (...ids: unknown[]) => ({
      requests: ids.map((id) => ({ ...req("/me/messages"), id })),
    });
    rejects(withIds("a", "a"));
    rejects(withIds("a", "A"));
    for (const bad of ["", "x".repeat(65), "a b", "a/b", "../x", 1, null, undefined, "é", "а1"]) {
      rejects(withIds(bad));
    }
    expect(() => validateBatch(withIds("x".repeat(64), "A_b-9"), targets, "v1.0")).not.toThrow();
  });

  it("limits headers to ConsistencyLevel/Prefer with printable-ASCII values", () => {
    const ok = validateBatch(
      { requests: [req("/me/messages", { headers: { consistencylevel: "eventual" } })] },
      targets,
      "v1.0",
    );
    expect(ok.requests[0]?.headers).toEqual({ ConsistencyLevel: "eventual" });
    for (const headers of [
      { Authorization: "Bearer x" },
      { authorization: "x" },
      { Host: "evil.example" },
      { "Content-Type": "application/json" },
      { "X-Anything": "1" },
      { Prefer: "a\r\nX: y" },
      { Prefer: "Ā" },
      { Prefer: 5 },
      { Prefer: "x", prefer: "y" },
      ["Prefer"],
      "Prefer",
      null,
    ]) {
      rejects({ requests: [req("/me/messages", { headers })] });
    }
  });
});

describe("batch validator: forbidden verb classes", () => {
  const verbs: unknown[] = [
    "POST",
    "PATCH",
    "PUT",
    "DELETE",
    "HEAD",
    "OPTIONS",
    "TRACE",
    "CONNECT",
    "get",
    "Get",
    "gET",
    "GET ",
    " GET",
    "GET\n",
    "\tGET",
    "",
    " ",
    undefined,
    null,
    0,
    true,
    ["GET"],
    { toString: () => "GET" },
    "G\u0415T", // Cyrillic Ie
    "GE\u0422", // Cyrillic Te
    "\uff27\uff25\uff34", // fullwidth
    "G\u200bET", // zero-width space
    "\u212aET", // Kelvin sign
  ];
  it("rejects every non-exact-GET method", () => {
    for (const method of verbs) {
      rejects({ requests: [{ id: "1", method, url: "/me/messages" }] });
    }
    rejects({ requests: [{ id: "1", url: "/me/messages" }] });
  });
});

describe("batch validator: relative URL (hostile input)", () => {
  const hostile = [
    "https://graph.microsoft.com/v1.0/me/messages",
    "http://graph.microsoft.com/me/messages",
    "//graph.microsoft.com/me/messages",
    "//evil.example/me/messages",
    "\\\\evil.example\\me",
    "/\\evil.example/me/messages",
    "/me\\messages",
    "javascript:alert(1)",
    "mailto:a@b.c",
    "file:///etc/passwd",
    "ftp://x/me/messages",
    "me/messages", // leading slash is required
    "?$top=5",
    "",
    " /me/messages",
    "/me/messages ",
    "/me/ messages",
    "/me/messages\n",
    "/me/messages\u0000",
    "/me/messages#frag",
    "/me/messages?$top=5#frag",
    "/me/%2e%2e/messages",
    "/me/%2E%2E/messages",
    "/me/./messages",
    "/me/../me/messages",
    "/me/messages/..",
    "/me/%2fmessages",
    "/me%2Fmessages",
    "/me/%5cmessages",
    "/me/%00messages",
    "/me/%252e%252e/messages",
    "/me/%252fmessages",
    "/me/%25252e/messages",
    "/me/%zzmessages",
    "/me/%",
    "/me//messages",
    "/me/messages/é",
    "/v1.0/me/messages",
    "/V1.0/me/messages",
    "/beta/me/messages",
    "/Beta/me/messages",
    "/me/messages?$top",
    "/me/messages?=1",
    "/me/messages?$top=1&&$skip=1",
    "/me/messages?$top=1&",
    "/me/messages?%24top=1",
    "/me/messages?a b=1",
    "/me/messages?$filter=a+b",
    "/me/messages?$filter=%zz",
    "/me/messages?$filter=a\\b",
    "/me/messages?$filter=a%0Ab\u0001",
    `/me/messages?${"a=1&".repeat(60)}b=2`,
    `/me/messages/${"a".repeat(9000)}`,
    "/me/unlisted-endpoint",
    "/me/messages/abc/extra/segments/here",
    "/me/sendMail",
    "/search/query",
    "/copilot/retrieval",
    "/me/profile", // beta-only catalog entry; not reachable through a v1.0 batch
  ];
  it("rejects absolute/protocol-relative/scheme/backslash/encoded/dot/version/unlisted URLs", () => {
    for (const url of hostile) {
      expect(() => urlRejects(url), JSON.stringify(url.slice(0, 80))).not.toThrow();
    }
    for (const nonString of [undefined, null, 1, {}, ["/me/messages"], true]) {
      urlRejects(nonString);
    }
  });

  it("rejects nested $batch in every spelling", () => {
    for (const url of [
      "/$batch",
      "/$Batch",
      "/$BATCH",
      "/%24batch",
      "/%24BATCH",
      "/%2524batch",
      "/$batch/",
      "/$batch?x=1",
      "/me/$batch",
      "/me/messages/$batch",
      "/$batch;x",
      "/%zz$batch",
      "/me/messages?$expand=/$batch",
      "/me/messages?x=%2F%24batch",
      "/v1.0/$batch",
      "/beta/$batch",
      "/me/../$batch",
      "/%2e%2e/$batch",
    ]) {
      expect(() => urlRejects(url), url).not.toThrow();
    }
  });

  it("matches only catalog GET entries and never a POST-only path", () => {
    expect(permitted.every((e) => e.method === "get")).toBe(true);
    urlRejects("/search/query");
    urlRejects("/me/calendar/getSchedule");
    urlRejects("/me/findMeetingTimes");
  });

  it("only matches entries of the batch version", () => {
    const beta = permitted.find((e) => e["apiVersion"] === "beta");
    expect(beta).toBeDefined();
    const sample = (beta as EndpointDefinition).pathPattern.replace(/\{[^}]+\}/g, "x1");
    expect(() => validateRelativeGraphUrl(sample, permitted, "v1.0")).toThrow(GraphClientError);
    expect(validateRelativeGraphUrl(sample, permitted, "beta").entry).toBe(beta);
    // A beta-labelled entry never satisfies a v1.0 batch even when passed explicitly.
    expect(() => validateRelativeGraphUrl(sample, [beta as EndpointDefinition], "v1.0")).toThrow();
  });

  it("takes the permitted set as a parameter (tightening is wiring only)", () => {
    const only = targetsOf(permitted.filter((e) => e.pathPattern === "/me/messages"));
    expect(() => validateBatch({ requests: [req("/me/messages")] }, only, "v1.0")).not.toThrow();
    expect(() => validateBatch({ requests: [req("/me/mailFolders")] }, only, "v1.0")).toThrow();
  });
});

describe("batch validator: {path} / {url} skipEncoding entries", () => {
  const site = "/sites/contoso.sharepoint.com:";
  it("accepts a normal site-relative path", () => {
    const r = validateRelativeGraphUrl(`${site}/sites/marketing`, permitted, "v1.0");
    expect(r.entry.toolName).toBe("get-sharepoint-site-by-path");
    expect(r.url).toBe(`${site}/sites/marketing`);
    expect(() =>
      validateRelativeGraphUrl("/sites/contoso.sharepoint.com:/sites/a%20b", permitted, "v1.0"),
    ).not.toThrow();
  });

  it("requires a plain hostname before ':/' (no root, no composite or GUID site id)", () => {
    const guid = "f17e0000-0000-4000-8000-000000000004";
    const guid2 = "f17e0000-0000-4000-8000-000000000005";
    for (const id of [
      "root",
      guid,
      `${guid},${guid2}`,
      `contoso.sharepoint.com,${guid},${guid2}`,
    ]) {
      expect(
        () => validateRelativeGraphUrl(`/sites/${id}:/sites/marketing`, permitted, "v1.0"),
        id,
      ).toThrow(GraphClientError);
    }
  });

  it("rejects traversal, encoded separators, nested batch, scheme tricks inside {path}", () => {
    for (const tail of [
      "/../../me/messages",
      "/sites/../../me",
      "/sites/%2e%2e/x",
      "/sites/%2E%2e/x",
      "/sites%2Fa",
      "/sites/a%5Cb",
      "/sites/a%00",
      "//evil.example/x",
      "/sites//x",
      "/$batch",
      "/sites/$batch",
      "/sites/%24batch",
      "/https://evil.example/x",
      "/sites/a\\b",
      "/sites/a#b",
      "/sites/a%252e%252e/b",
      "/sites/marketing:/drive/root/children",
      "/sites/marketing:",
      "/sites/marketing%3A/drive/root/children",
      "/sites/marketing%3a/onenote/notebooks",
    ]) {
      expect(() => validateRelativeGraphUrl(`${site}${tail}`, permitted, "v1.0"), tail).toThrow(
        GraphClientError,
      );
    }
  });

  it("rejects hostile {site-id} and a missing {path}", () => {
    for (const url of [
      "/sites/../me:/x",
      "/sites/%2e%2e:/x",
      "/sites/a/b/c/d/e:/x/y/z/q/w/e/r",
      "/sites/x:",
      "/sites/x:/",
    ]) {
      // Either unmatched or matched to a different catalog entry; never accepted as traversal.
      try {
        const r = validateRelativeGraphUrl(url, permitted, "v1.0");
        expect(r.url).toBe(url);
        expect(url).not.toMatch(/\.\.|%2e/i);
      } catch (e) {
        expect(e).toBeInstanceOf(GraphClientError);
      }
    }
  });

  it("every catalog {path}/{url}-style skipEncoding entry has hostile coverage", () => {
    const skip = permitted.filter((e) => Array.isArray(e["skipEncoding"]));
    expect(skip.map((e) => e.pathPattern)).toEqual(["/sites/{site-id}:/{path}"]);
  });
});

describe("GraphClient.batch", () => {
  const MAX = 5 * 1024 * 1024;
  void MAX;
  function setup(
    responder: (url: string, init: RequestInit, n: number) => Response,
    extra: Partial<GraphClientOptions> = {},
    token = TOKEN,
  ) {
    const calls: { url: string; init: RequestInit }[] = [];
    const tokenProvider: TokenProvider = { getAccessToken: () => Promise.resolve(token) };
    const fetch: FetchLike = (url, init) => {
      calls.push({ url, init });
      return Promise.resolve(responder(url, init, calls.length));
    };
    const client = new GraphClient({
      tokenProvider,
      catalog,
      fetch,
      sleep: () => Promise.resolve(),
      ...extra,
    });
    return { client, calls };
  }
  const json = (body: unknown, init: ResponseInit = {}): Response =>
    new Response(JSON.stringify(body), {
      status: 200,
      headers: { "content-type": "application/json" },
      ...init,
    });
  const two = () =>
    validateBatch(
      { requests: [req("/me/messages?$top=1"), { ...req("/me/mailFolders"), id: "two" }] },
      targets,
      "v1.0",
    );
  const sub = (id: string, status: number, body?: unknown, headers: object = {}) => ({
    id,
    status,
    headers,
    ...(body === undefined ? {} : { body }),
  });

  async function fail(p: Promise<unknown>): Promise<GraphClientError> {
    try {
      await p;
    } catch (e) {
      expect(e).toBeInstanceOf(GraphClientError);
      return e as GraphClientError;
    }
    throw new Error("expected failure");
  }

  it("sends one POST to the fixed $batch URL with bearer and a GET-only wire body", async () => {
    const { client, calls } = setup(() =>
      json({ responses: [sub("two", 200, { value: [2] }), sub("1", 200, { value: [1] })] }),
    );
    const out = await client.batch(two());
    expect(calls).toHaveLength(1);
    expect(calls[0]?.url).toBe(`${G}/v1.0/$batch`);
    expect(calls[0]?.init.method).toBe("POST");
    expect(calls[0]?.init.redirect).toBe("manual");
    const h = calls[0]?.init.headers as Record<string, string>;
    expect(h["Authorization"]).toBe(`Bearer ${TOKEN}`);
    expect(h["Content-Type"]).toBe("application/json");
    const wire = JSON.parse(String(calls[0]?.init.body)) as { requests: { method: string }[] };
    expect(wire.requests.map((r) => r.method)).toEqual(["GET", "GET"]);
    expect(String(calls[0]?.init.body)).toBe(encodeBatchBody(two()));
    // Results are returned in request order regardless of Graph's order.
    expect(out.responses.map((r) => r.id)).toEqual(["1", "two"]);
    expect(out.responses[0]?.body).toEqual({ value: [1] });
  });

  it("refuses anything that did not come from the validator, with zero activity", async () => {
    const { client, calls } = setup(() => json({}));
    const forged = {
      version: "v1.0",
      requests: [{ id: "1", url: "/me", headers: {}, entry: catalog[0] }],
    };
    const err = await fail(client.batch(forged as never));
    expect(err.code).toBe("invalid_input");
    await fail(client.batch(null as never));
    expect(calls).toHaveLength(0);
  });

  it("generic request() still refuses $batch", async () => {
    const { client, calls } = setup(() => json({}));
    const err = await fail(
      client.request({
        pathPattern: "/$batch",
        method: "post",
        body: { requests: [{ id: "1", method: "GET", url: "/me" }] },
      }),
    );
    expect(err.code).toBe("invalid_input");
    expect(err.message).toMatch(/batch/);
    expect(calls).toHaveLength(0);
  });

  it("reuses send(): retries a throttled batch, honoring Retry-After", async () => {
    const sleeps: number[] = [];
    const { client, calls } = setup(
      (_u, _i, n) =>
        n === 1
          ? new Response("{}", { status: 429, headers: { "retry-after": "2" } })
          : json({ responses: [sub("1", 200, {}), sub("two", 200, {})] }),
      {
        sleep: (ms) => {
          sleeps.push(ms);
          return Promise.resolve();
        },
      },
    );
    await client.batch(two());
    expect(calls).toHaveLength(2);
    expect(sleeps).toEqual([2000]);
  });

  it("fails closed on redirect, oversize, and non-JSON top-level responses", async () => {
    const redirect = setup(
      () => new Response(null, { status: 302, headers: { location: "https://evil.example" } }),
    );
    expect((await fail(redirect.client.batch(two()))).code).toBe("graph_error");
    const big = setup(() => json({ responses: [] }), { maxResponseBytes: 10 });
    expect((await fail(big.client.batch(two()))).message).toMatch(/size limit/);
    const text = setup(() => new Response("<html>", { status: 200 }));
    expect((await fail(text.client.batch(two()))).code).toBe("graph_error");
  });

  it("oversize request body is refused before any network call", async () => {
    const { client, calls } = setup(() => json({}), { maxRequestBytes: 20 });
    expect((await fail(client.batch(two()))).code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("top-level 401/403 map to reauth_required/capability_gap without the token", async () => {
    const a = setup(() =>
      json({ error: { code: "x", message: `echo ${TOKEN}` } }, { status: 401 }),
    );
    const e1 = await fail(a.client.batch(two()));
    expect(e1.code).toBe("reauth_required");
    expect(JSON.stringify(e1.toJSON())).not.toContain(TOKEN);
    const b = setup(() => json({ error: { code: "x" } }, { status: 403 }));
    expect((await fail(b.client.batch(two()))).code).toBe("capability_gap");
  });

  describe("untrusted subresponses", () => {
    const cases: [string, unknown][] = [
      ["unknown id", { responses: [sub("1", 200, {}), sub("zzz", 200, {})] }],
      ["duplicate id", { responses: [sub("1", 200, {}), sub("1", 200, {})] }],
      ["duplicate id (case)", { responses: [sub("two", 200, {}), sub("TWO", 200, {})] }],
      ["missing id", { responses: [sub("1", 200, {})] }],
      ["too many", { responses: [sub("1", 200), sub("two", 200), sub("1", 200)] }],
      ["no responses", {}],
      ["responses not array", { responses: {} }],
      ["non-object item", { responses: ["1", 5] }],
      ["bad status", { responses: [sub("1", 99), sub("two", 200)] }],
      ["string status", { responses: [{ id: "1", status: "200" }, sub("two", 200)] }],
      ["fractional status", { responses: [sub("1", 200.5), sub("two", 200)] }],
      ["missing id field", { responses: [{ status: 200 }, sub("two", 200)] }],
      ["top-level array", [sub("1", 200), sub("two", 200)]],
    ];
    for (const [name, body] of cases) {
      it(`fails closed: ${name}`, async () => {
        const { client } = setup(() => json(body));
        const err = await fail(client.batch(two()));
        expect(err.code).toBe("graph_error");
        expect(err.message).toMatch(/invalid batch response/);
      });
    }

    it("sanitizes non-2xx subresponses and never passes raw bodies or Location through", async () => {
      const jwt = "eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJ4In0.c2ln";
      const { client } = setup(() =>
        json({
          responses: [
            sub(
              "1",
              403,
              {
                error: {
                  code: "Forbidden\u0007",
                  message: `Denied. Bearer abc.def ${jwt}\u0000 ${"x".repeat(500)}`,
                  innerError: { "request-id": "req-123", secret: "leak" },
                },
                extra: "raw",
              },
              { Location: "https://evil.example/x", "Retry-After": "7", "X-Other": "y" },
            ),
            sub("two", 429, "not-json", { "retry-after": "Wed, 21 Oct 2026 07:28:00 GMT\r\n" }),
          ],
        }),
      );
      const out = await client.batch(two());
      const text = JSON.stringify(out);
      expect(text).not.toMatch(/evil\.example|Location|leak|raw|X-Other|eyJ|abc\.def/);
      const first = out.responses[0];
      expect(first?.error?.code).toBe("capability_gap");
      expect(first?.error?.graphMessage?.length).toBeLessThanOrEqual(203);
      expect(first?.error?.requestId).toBe("req-123");
      expect(first?.error?.retryAfter).toBe("7");
      expect(first?.body).toBeUndefined();
      expect(out.responses[1]?.error?.code).toBe("rate_limited");
      expect(out.responses[1]?.error?.retryAfter).not.toMatch(/[\r\n]/);
    });

    it("keeps only Content-Type from success headers", async () => {
      const { client } = setup(() =>
        json({
          responses: [
            sub(
              "1",
              200,
              { a: 1 },
              { "Content-Type": "application/json", Location: "https://e.example" },
            ),
            sub("two", 204),
          ],
        }),
      );
      const out = await client.batch(two());
      expect(out.responses[0]).toEqual({
        id: "1",
        status: 200,
        body: { a: 1 },
        contentType: "application/json",
      });
      expect(out.responses[1]).toEqual({ id: "two", status: 204 });
    });

    it("omits base64 (non-JSON) bodies, reporting only an approximate size", async () => {
      const b64 = Buffer.from("A".repeat(3000)).toString("base64");
      const { client } = setup(() =>
        json({
          responses: [sub("1", 200, b64, { "content-type": "image/png" }), sub("two", 200, {})],
        }),
      );
      const out = await client.batch(two());
      expect(out.responses[0]?.bodyOmitted).toEqual({ reason: "non_json", approxBytes: 3000 });
      expect(out.responses[0]?.body).toBeUndefined();
      expect(JSON.stringify(out)).not.toContain(b64.slice(0, 40));
    });

    it("maps 410 on delta subrequests to restart_required", async () => {
      const b = validateBatch(
        { requests: [req("/me/mailFolders('inbox')/messages/delta()")] },
        targets,
        "v1.0",
      );
      const { client } = setup(() =>
        json({ responses: [sub("1", 410, { error: { code: "syncStateNotFound" } })] }),
      );
      expect((await client.batch(b)).responses[0]?.error?.code).toBe("restart_required");
    });
  });

  describe("token-echo protection", () => {
    it("refuses a 2xx batch response echoing the token in a subresponse body", async () => {
      const { client } = setup(() =>
        json({ responses: [sub("1", 200, { v: `x${TOKEN}y` }), sub("two", 200, {})] }),
      );
      const err = await fail(client.batch(two()));
      expect(err.message).toMatch(/credential material/);
      expect(JSON.stringify(err.toJSON())).not.toContain(TOKEN);
    });

    it("refuses a token echoed in a subresponse error message or header value", async () => {
      for (const responses of [
        [sub("1", 500, { error: { message: `t ${TOKEN}` } }), sub("two", 200, {})],
        [sub("1", 200, {}, { "Content-Type": TOKEN }), sub("two", 200, {})],
        [sub("1", 200, { [TOKEN]: 1 }), sub("two", 200, {})],
      ]) {
        const { client } = setup(() => json({ responses }));
        const err = await fail(client.batch(two()));
        expect(JSON.stringify(err.toJSON())).not.toContain(TOKEN);
      }
    });
  });
});
