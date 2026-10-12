import { beforeAll, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "../catalog.js";
import { TokenError, type TokenProvider } from "../auth/token-provider.js";
import {
  GraphClient,
  type FetchLike,
  type GraphClientOptions,
  type GraphRequest,
} from "./client.js";
import { GraphClientError } from "./errors.js";

const TOKEN = "FAKE_TOKEN_aB3dE5gH7jK9mN1pQ3sT5vW7yZ_SECRET";
const G = "https://graph.microsoft.com";

let catalog: EndpointCatalog;
beforeAll(async () => {
  catalog = await loadEndpointCatalog();
});

interface Call {
  url: string;
  init: RequestInit;
}

function json(body: unknown, init: ResponseInit = {}): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { "content-type": "application/json" },
    ...init,
  });
}

function setup(
  responder: (call: Call, n: number) => Response | Promise<Response>,
  extra: Partial<GraphClientOptions> = {},
) {
  const calls: Call[] = [];
  const sleeps: number[] = [];
  let tokenCalls = 0;
  const tokenProvider: TokenProvider = {
    getAccessToken: () => {
      tokenCalls += 1;
      return Promise.resolve(TOKEN);
    },
  };
  const fetch: FetchLike = (url, init) => {
    calls.push({ url, init });
    return Promise.resolve(responder({ url, init }, calls.length));
  };
  const client = new GraphClient({
    tokenProvider,
    catalog,
    fetch,
    sleep: (ms) => {
      sleeps.push(ms);
      return Promise.resolve();
    },
    now: () => Date.UTC(2030, 0, 1, 0, 0, 0),
    ...extra,
  });
  return { client, calls, sleeps, tokenCalls: () => tokenCalls };
}

async function fail(p: Promise<unknown>): Promise<GraphClientError> {
  try {
    await p;
  } catch (e) {
    expect(e).toBeInstanceOf(GraphClientError);
    return e as GraphClientError;
  }
  throw new Error("expected failure");
}

const MSGS: GraphRequest = { pathPattern: "/me/messages", query: { $top: 2 } };
const authOf = (c: Call): string =>
  (c.init.headers as Record<string, string>)["Authorization"] ?? "";

describe("routing and allowlist", () => {
  it("builds the URL on the fixed origin and sends bearer + safe options", async () => {
    const { client, calls } = setup(() => json({ value: [] }));
    await client.request({
      pathPattern: "/me/messages/{message-id}",
      pathParams: { "message-id": "AAMkAGI+TG93-_AAA=" },
      query: { $select: "subject" },
    });
    expect(calls).toHaveLength(1);
    expect(calls[0]?.url).toBe(`${G}/v1.0/me/messages/AAMkAGI%2BTG93-_AAA=?$select=subject`);
    expect(calls[0]?.init.redirect).toBe("manual");
    expect(calls[0]?.init.method).toBe("GET");
    expect(authOf(calls[0] as Call)).toBe(`Bearer ${TOKEN}`);
  });

  it("rejects runtime-invalid query values before acquiring a bearer or making a request", async () => {
    const { client, calls, tokenCalls } = setup(() => json({ value: [] }));
    for (const value of [["id", "subject"], null, { id: "subject" }]) {
      const query = { $select: value } as unknown as NonNullable<GraphRequest["query"]>;
      const err = await fail(client.request({ pathPattern: "/me/messages", query }));
      expect(err.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
    expect(tokenCalls()).toBe(0);
  });

  it("uses beta only for catalog beta entries and refuses mismatches", async () => {
    const { client, calls } = setup(() => json({}));
    await client.request({ pathPattern: "/me/profile" });
    expect(calls[0]?.url).toBe(`${G}/beta/me/profile`);
    expect((await fail(client.request({ pathPattern: "/me/profile", version: "v1.0" }))).code).toBe(
      "invalid_input",
    );
    expect((await fail(client.request({ pathPattern: "/me", version: "beta" }))).code).toBe(
      "invalid_input",
    );
    expect(calls).toHaveLength(1);
  });

  it("refuses non-allowlisted method/path pairs with zero network or token activity", async () => {
    const { client, calls, tokenCalls } = setup(() => json({}));
    const refused: GraphRequest[] = [
      { pathPattern: "/me/messages", method: "post", body: {} },
      { pathPattern: "/me/sendMail", method: "post", body: {} },
      {
        pathPattern: "/me/messages/{message-id}/forward",
        method: "post",
        body: {},
        pathParams: { "message-id": "x" },
      },
      { pathPattern: "/me/other" },
      { pathPattern: "/me/messages/../users" },
      { pathPattern: "https://evil.example/x" },
      { pathPattern: "/ME/messages" },
      { pathPattern: "/search/query" }, // catalog POST, wrong method
      { pathPattern: "/me/messages", method: "delete" as never },
      { pathPattern: "/me/messages", method: "patch" as never },
      { pathPattern: "/me/messages", method: "put" as never },
    ];
    for (const r of refused) expect((await fail(client.request(r))).code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
    expect(tokenCalls()).toBe(0);
  });

  it("refuses $batch through the generic path in every spelling, with zero activity", async () => {
    const { client, calls, tokenCalls } = setup(() => json({}));
    const spellings = [
      "/$batch",
      "/$Batch",
      "/$BATCH",
      "/%24batch",
      "/%24BATCH",
      "/%2524batch",
      "/$batch/",
      "//$batch",
      "/$batch?x=1",
      "/$batch#frag",
      "/beta/$batch",
      "/v1.0/$batch",
      "/me/../$batch",
      "/./$batch",
      "/%2e%2e/$batch",
      "\\$batch",
      "/$batch;x",
      "/%zz$batch",
    ];
    for (const pathPattern of spellings) {
      for (const method of ["post", "get"] as const) {
        const body =
          method === "post" ? { requests: [{ id: "1", method: "DELETE", url: "/me" }] } : undefined;
        const req: GraphRequest = { pathPattern, method };
        if (body !== undefined) req.body = body;
        expect((await fail(client.request(req))).code).toBe("invalid_input");
      }
    }
    for (const version of ["v1.0", "beta"] as const) {
      const err = await fail(
        client.request({ pathPattern: "/$batch", method: "post", body: {}, version }),
      );
      expect(err.message).toMatch(/batch/);
    }
    expect(
      (await fail(client.fetchLink({ pathPattern: "/$batch" }, `${G}/v1.0/$batch`))).code,
    ).toBe("invalid_input");
    expect(calls).toHaveLength(0);
    expect(tokenCalls()).toBe(0);
  });

  it("refuses a path parameter value that would become a $batch segment", async () => {
    const { client, calls } = setup(() => json({}));
    for (const value of ["$batch", "%24batch", "$BATCH"]) {
      const err = await fail(
        client.request({
          pathPattern: "/users/{user-id}/messages",
          pathParams: { "user-id": value },
        }),
      );
      expect(err.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("ties POST to catalog (method, path) pairs; PATCH/DELETE/PUT never pass", async () => {
    const { client, calls } = setup(() => json({}));
    const getOnly = catalog.filter((e) => e.method === "get").slice(0, 10);
    for (const e of getOnly) {
      const pathParams = Object.fromEntries(
        [...e.pathPattern.matchAll(/{([^{}/]+)}/g)].map((m) => [m[1] as string, "x"]),
      );
      expect(
        (
          await fail(
            client.request({ pathPattern: e.pathPattern, method: "post", body: {}, pathParams }),
          )
        ).code,
      ).toBe("invalid_input");
      for (const method of ["patch", "delete", "put"]) {
        expect(
          (
            await fail(
              client.request({ pathPattern: e.pathPattern, method: method as never, pathParams }),
            )
          ).code,
        ).toBe("invalid_input");
      }
    }
    expect(calls).toHaveLength(0);
  });

  it("allows exactly the catalog POST entries except $batch (71 GET + 5 POST)", async () => {
    const posts = catalog.filter((e) => e.method === "post").map((e) => e.pathPattern);
    expect(posts).toHaveLength(5);
    expect(catalog.filter((e) => e.method === "get")).toHaveLength(71);
    for (const pathPattern of posts.filter((p) => p !== "/$batch")) {
      const { client, calls } = setup(() => json({ ok: 1 }));
      await client.request({ pathPattern, method: "post", body: { a: 1 } });
      expect(calls[0]?.init.method).toBe("POST");
      expect(calls[0]?.url).toBe(`${G}/v1.0${pathPattern}`);
      expect(calls[0]?.init.body).toBe('{"a":1}');
    }
  });

  it("validates bodies: GET with body, POST without, oversize, unserializable", async () => {
    const { client, calls } = setup(() => json({}), { maxRequestBytes: 50 });
    const post = { pathPattern: "/search/query", method: "post" } as const;
    expect((await fail(client.request({ ...MSGS, body: {} }))).code).toBe("invalid_input");
    expect((await fail(client.request(post))).code).toBe("invalid_input");
    expect((await fail(client.request({ ...post, body: { q: "x".repeat(100) } }))).code).toBe(
      "invalid_input",
    );
    expect((await fail(client.request({ ...post, body: { n: 10n } }))).code).toBe("invalid_input");
    expect(calls).toHaveLength(0);
  });

  it("only passes allowlisted extra headers and never lets callers set Authorization", async () => {
    const { client, calls } = setup(() => json({}));
    await client.request({ ...MSGS, headers: { consistencylevel: "eventual" } });
    expect((calls[0]?.init.headers as Record<string, string>)["ConsistencyLevel"]).toBe("eventual");
    for (const headers of [
      { Authorization: "Bearer x" },
      { Host: "evil" },
      { Prefer: "a\r\nX: y" },
      { Prefer: 'outlook.timezone="Ā"' },
    ]) {
      expect((await fail(client.request({ ...MSGS, headers }))).code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(1);
  });

  it("fetches a fresh token for every request", async () => {
    const { client, tokenCalls } = setup(() => json({}));
    await client.request(MSGS);
    await client.request(MSGS);
    expect(tokenCalls()).toBe(2);
  });
});

describe("redirects, timeouts, bounded bodies", () => {
  it("never follows redirects", async () => {
    const { client, calls } = setup(
      () => new Response(null, { status: 302, headers: { location: "https://evil.example/x" } }),
    );
    const err = await fail(client.request(MSGS));
    expect(err.code).toBe("graph_error");
    expect(JSON.stringify(err)).not.toContain("evil.example");
    expect(calls).toHaveLength(1);
  });

  it("times out via AbortController", async () => {
    const { client } = setup(
      ({ init }) =>
        new Promise<Response>((_res, rej) => {
          init.signal?.addEventListener("abort", () =>
            rej(new DOMException("aborted", "AbortError")),
          );
        }) as unknown as Response,
      { timeoutMs: 20 },
    );
    const err = await fail(client.request(MSGS));
    expect(err.code).toBe("graph_error");
    expect(err.message).toMatch(/timed out/);
  });

  it("aborts a streamed body that exceeds the cap", async () => {
    let cancelled = false;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        controller.enqueue(new Uint8Array(400));
      },
      cancel() {
        cancelled = true;
      },
    });
    const { client } = setup(() => new Response(stream, { status: 200 }), {
      maxResponseBytes: 1000,
    });
    const err = await fail(client.request(MSGS));
    expect(err.code).toBe("graph_error");
    expect(err.message).toMatch(/size limit/);
    expect(cancelled).toBe(true);
  });

  it("rejects an oversized declared content-length without reading", async () => {
    const { client } = setup(
      () => new Response("{}", { status: 200, headers: { "content-length": "999999999" } }),
      { maxResponseBytes: 1000 },
    );
    expect((await fail(client.request(MSGS))).code).toBe("graph_error");
  });

  it("maps invalid JSON and network failures to graph_error", async () => {
    const bad = setup(() => new Response("<html>", { status: 200 }));
    expect((await fail(bad.client.request(MSGS))).code).toBe("graph_error");
    const net = setup(() => {
      throw Object.assign(new TypeError(`fetch failed ${TOKEN}`), {
        cause: { code: "ECONNRESET" },
      });
    });
    const err = await fail(net.client.request(MSGS));
    expect(err.code).toBe("graph_error");
    expect(err.details?.["networkCode"]).toBe("ECONNRESET");
    expect(JSON.stringify(err)).not.toContain(TOKEN);
  });

  it("handles 204/empty successes", async () => {
    const { client } = setup(() => new Response(null, { status: 204 }));
    const res = await client.request(MSGS);
    expect(res.status).toBe(204);
    expect(res.data).toBeUndefined();
  });
});

describe("response body parsing", () => {
  it("tolerates a UTF-8 BOM before the JSON body", async () => {
    const { client } = setup(
      () => new Response(`\uFEFF${JSON.stringify({ value: [1] })}`, { status: 200 }),
    );
    expect((await client.request(MSGS)).data).toEqual({ value: [1] });
  });
});

describe("429 / Retry-After", () => {
  const throttled = (retryAfter?: string): Response =>
    json(
      { error: { code: "TooManyRequests", message: "slow down" } },
      { status: 429, headers: retryAfter === undefined ? {} : { "retry-after": retryAfter } },
    );

  it("retries honoring Retry-After seconds", async () => {
    const { client, calls, sleeps } = setup((_c, n) =>
      n < 3 ? throttled("2") : json({ ok: true }),
    );
    const res = await client.request(MSGS);
    expect(res.status).toBe(200);
    expect(calls).toHaveLength(3);
    expect(sleeps).toEqual([2000, 2000]);
  });

  it("honors an HTTP-date Retry-After relative to the injected clock", async () => {
    const date = new Date(Date.UTC(2030, 0, 1, 0, 0, 7)).toUTCString();
    const { client, sleeps } = setup((_c, n) => (n < 2 ? throttled(date) : json({})));
    await client.request(MSGS);
    expect(sleeps).toEqual([7000]);
  });

  it("uses bounded exponential backoff without or with garbage Retry-After", async () => {
    const { client, sleeps } = setup((_c, n) => (n < 3 ? throttled("garbage") : json({})));
    await client.request(MSGS);
    expect(sleeps).toEqual([1000, 2000]);
  });

  it("treats numeric-looking non-integer Retry-After as garbage, not a past date", async () => {
    // Date.parse("5.5") / ("-1") / ("1,5") yield past dates, which would mean a zero wait.
    for (const bad of ["5.5", "-1", "1,5", "10 20"]) {
      const { client, sleeps } = setup((_c, n) => (n < 2 ? throttled(bad) : json({})));
      await client.request(MSGS);
      expect(sleeps).toEqual([1000]);
    }
  });

  it("surfaces rate_limited after maxRetries", async () => {
    const { client, calls } = setup(() => throttled("1"), { maxRetries: 2 });
    const err = await fail(client.request(MSGS));
    expect(err.code).toBe("rate_limited");
    expect(calls).toHaveLength(3);
    expect(err.details?.["retryAfter"]).toBe("1");
  });

  it("does not wait for absurd Retry-After values", async () => {
    const { client, calls, sleeps } = setup(() => throttled("86400"));
    expect((await fail(client.request(MSGS))).code).toBe("rate_limited");
    expect(calls).toHaveLength(1);
    expect(sleeps).toEqual([]);
    const past = setup(() => throttled("99999999999999999999"));
    expect((await fail(past.client.request(MSGS))).code).toBe("rate_limited");
    expect(past.sleeps).toEqual([]);
  });

  it("enforces the total wait budget", async () => {
    const { client, sleeps } = setup(() => throttled("20"), {
      maxRetries: 10,
      maxTotalWaitMs: 50_000,
    });
    expect((await fail(client.request(MSGS))).code).toBe("rate_limited");
    expect(sleeps).toEqual([20_000, 20_000]);
  });

  it("clamps a past HTTP-date to zero wait", async () => {
    const past = new Date(Date.UTC(2020, 0, 1)).toUTCString();
    const { client, sleeps } = setup((_c, n) => (n < 2 ? throttled(past) : json({})));
    await client.request(MSGS);
    expect(sleeps).toEqual([0]);
  });

  it("retries 503/504 and reports graph_error when exhausted", async () => {
    const { client, calls } = setup(() => new Response("x", { status: 503 }), { maxRetries: 1 });
    const err = await fail(client.request(MSGS));
    expect(err.code).toBe("graph_error");
    expect(calls).toHaveLength(2);
  });
});

describe("pagination and nextLink", () => {
  const page = (items: unknown[], next?: string): Response =>
    json(next === undefined ? { value: items } : { value: items, "@odata.nextLink": next });

  it("follows validated nextLinks, rebuilding against the fixed origin", async () => {
    const { client, calls } = setup((_c, n) =>
      n === 1 ? page([1, 2], `${G}/v1.0/me/messages?$top=2&$skip=2`) : page([3]),
    );
    const out = await client.paginate(MSGS);
    expect(out).toEqual({ items: [1, 2, 3], pages: 2, truncated: false });
    expect(calls[1]?.url).toBe(`${G}/v1.0/me/messages?$top=2&$skip=2`);
  });

  it("enforces page and item ceilings", async () => {
    const link = (n: number): string => `${G}/v1.0/me/messages?$skip=${n}`;
    const byPages = setup((_c, n) => page([1, 2], link(n)));
    const a = await byPages.client.paginate(MSGS, { maxPages: 3 });
    expect(a).toMatchObject({ pages: 3, truncated: true, nextLink: link(3) });
    expect(a.items).toHaveLength(6);
    const byItems = setup((_c, n) => page([1, 2, 3], link(n)));
    const b = await byItems.client.paginate(MSGS, { maxItems: 4 });
    expect(b.items).toEqual([1, 2, 3, 1]);
    expect(b.truncated).toBe(true);
    expect(b.nextLink).toBeUndefined();
    const exact = setup((_c, n) => page([1, 2], link(n)));
    const c = await exact.client.paginate(MSGS, { maxItems: 2 });
    expect(c).toMatchObject({ items: [1, 2], truncated: true, nextLink: link(1) });
  });

  it("fails closed on a successful page without a value array (first or later page)", async () => {
    const first = setup(() => json({ error: "backend failure" }));
    expect((await fail(first.client.paginate(MSGS))).code).toBe("graph_error");
    const later = setup((_c, n) =>
      n === 1 ? page([1], `${G}/v1.0/me/messages?$skip=1`) : json({ value: "x" }),
    );
    expect((await fail(later.client.paginate(MSGS))).code).toBe("graph_error");
    const empty = setup(() => page([]));
    expect(await empty.client.paginate(MSGS)).toEqual({ items: [], pages: 1, truncated: false });
  });

  it("detects nextLink loops", async () => {
    const { client } = setup(() => page([1], `${G}/v1.0/me/messages?$skip=1`));
    expect((await fail(client.paginate(MSGS, { maxPages: 50 }))).code).toBe("graph_error");
  });

  const evil: [string, string][] = [
    ["other host", "https://evil.example/v1.0/me/messages?$skip=1"],
    ["lookalike", "https://graph.microsoft.com.evil.com/v1.0/me/messages"],
    ["http", "http://graph.microsoft.com/v1.0/me/messages"],
    ["userinfo", "https://graph.microsoft.com@evil.example/v1.0/me/messages"],
    ["port", "https://graph.microsoft.com:444/v1.0/me/messages"],
    ["different version", "https://graph.microsoft.com/beta/me/messages"],
    ["different endpoint", "https://graph.microsoft.com/v1.0/me/mailFolders"],
    ["protocol-relative", "//evil.example/v1.0/me/messages"],
    ["javascript", "javascript:alert(1)"],
    ["encoded trick", "https://graph.microsoft.com/v1.0/me/%2e%2e/users"],
    ["backslash", "https://evil.example\\@graph.microsoft.com/v1.0/me/messages"],
  ];
  it.each(evil)("malicious nextLink (%s): error, bearer never sent", async (_n, link) => {
    const { client, calls, tokenCalls } = setup((_c, n) => (n === 1 ? page([1], link) : page([2])));
    const err = await fail(client.paginate(MSGS));
    expect(err.code).toBe("invalid_input");
    expect(calls).toHaveLength(1); // only the original, legitimate request
    expect(calls.every((c) => c.url.startsWith(`${G}/`))).toBe(true);
    expect(tokenCalls()).toBe(1);
    expect(JSON.stringify(err)).not.toContain(TOKEN);
  });

  it("fetchLink refuses links on POST endpoints and non-catalog bases", async () => {
    const { client, calls } = setup(() => json({}));
    const post: GraphRequest = { pathPattern: "/search/query", method: "post", body: {} };
    expect((await fail(client.fetchLink(post, `${G}/v1.0/search/query`))).code).toBe(
      "invalid_input",
    );
    expect((await fail(client.fetchLink({ pathPattern: "/nope" }, `${G}/v1.0/nope`))).code).toBe(
      "invalid_input",
    );
    expect(calls).toHaveLength(0);
  });
});

describe("pagination hardening", () => {
  const page = (items: unknown[], next?: string): Response =>
    json(next === undefined ? { value: items } : { value: items, "@odata.nextLink": next });
  const headersOf = (c: Call | undefined): Record<string, string> =>
    (c?.init.headers ?? {}) as Record<string, string>;

  it("forwards allowlisted headers (ConsistencyLevel, Prefer) to follow-up pages", async () => {
    const { client, calls } = setup((_c, n) =>
      n === 1 ? page([1], `${G}/v1.0/me/messages?$skip=1`) : page([2]),
    );
    const base: GraphRequest = {
      ...MSGS,
      headers: { consistencylevel: "eventual", Prefer: 'outlook.timezone="UTC"' },
    };
    await client.paginate(base);
    expect(calls).toHaveLength(2);
    for (const c of calls) {
      expect(headersOf(c)["ConsistencyLevel"]).toBe("eventual");
      expect(headersOf(c)["Prefer"]).toBe('outlook.timezone="UTC"');
      expect(authOf(c)).toBe(`Bearer ${TOKEN}`);
    }
    await client.fetchLink(base, `${G}/v1.0/me/messages?$skip=2`);
    expect(headersOf(calls[2])["ConsistencyLevel"]).toBe("eventual");
  });

  it("rejects invalid fetchLink base queries before any bearer or network I/O", async () => {
    const { client, calls, tokenCalls } = setup(() => json({ value: [] }));
    const link = `${G}/v1.0/me/messages?$skip=10`;
    for (const value of [[], { id: "subject" }, null, 9n, Symbol("bad")]) {
      const query = { $select: value } as unknown as NonNullable<GraphRequest["query"]>;
      const err = await fail(client.fetchLink({ pathPattern: "/me/messages", query }, link));
      expect(err.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
    expect(tokenCalls()).toBe(0);
  });

  it("still rejects disallowed headers on fetchLink without any network call", async () => {
    const { client, calls } = setup(() => json({}));
    for (const headers of [{ Authorization: "Bearer x" }, { "X-Evil": "1" }]) {
      const err = await fail(
        client.fetchLink({ ...MSGS, headers }, `${G}/v1.0/me/messages?$skip=1`),
      );
      expect(err.code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("does not spread huge pages: 300k items with maxItems 1000", async () => {
    const big = new Array<number>(300_000).fill(0);
    const link = `${G}/v1.0/me/messages?$skip=300000`;
    const { client, calls } = setup(() => page(big, link));
    const out = await client.paginate(MSGS, { maxItems: 1000 });
    expect(out.items).toHaveLength(1000);
    expect(out).toMatchObject({ pages: 1, truncated: true });
    expect(out.nextLink).toBeUndefined(); // partially consumed page: no resume link
    expect(calls).toHaveLength(1);
  });

  it("keeps full-page vs partial-page semantics at the item ceiling", async () => {
    const link = `${G}/v1.0/me/messages?$skip=2`;
    const exact = setup(() => page(new Array<number>(1000).fill(1), link));
    const a = await exact.client.paginate(MSGS, { maxItems: 1000 });
    expect(a.items).toHaveLength(1000);
    expect(a).toMatchObject({ truncated: true, nextLink: link });
    const noMore = setup(() => page(new Array<number>(1000).fill(1)));
    expect(await noMore.client.paginate(MSGS, { maxItems: 1000 })).toMatchObject({
      truncated: false,
    });
  });

  it("rejects a nextLink that switches the principal and never sends the bearer to it", async () => {
    const alice: GraphRequest = {
      pathPattern: "/users/{user-id}/messages",
      pathParams: { "user-id": "alice@contoso.com" },
    };
    const swapped = setup((_c, n) =>
      n === 1 ? page([1], `${G}/v1.0/users/bob@contoso.com/messages?$skip=1`) : page([2]),
    );
    expect((await fail(swapped.client.paginate(alice))).code).toBe("invalid_input");
    expect(swapped.calls).toHaveLength(1);
    const ok = setup((_c, n) =>
      n === 1 ? page([1], `${G}/v1.0/users/alice@contoso.com/messages?$skip=1`) : page([2]),
    );
    expect((await ok.client.paginate(alice)).items).toEqual([1, 2]);
  });
});

describe("delta flows", () => {
  const DELTA: GraphRequest = {
    pathPattern: "/me/mailFolders/{mailFolder-id}/messages/delta()",
    pathParams: { "mailFolder-id": "inbox" },
  };

  it("walks nextLinks to the final deltaLink and exposes it", async () => {
    const dl = `${G}/v1.0/me/mailFolders('inbox')/messages/delta?$deltatoken=abc`;
    const { client, calls } = setup((_c, n) =>
      n === 1
        ? json({
            value: [1],
            "@odata.nextLink": `${G}/v1.0/me/mailFolders('inbox')/messages/delta?$skiptoken=s1`,
          })
        : json({ value: [2], "@odata.deltaLink": dl }),
    );
    expect(calls).toHaveLength(0);
    const out = await client.paginate(DELTA);
    expect(out).toEqual({ items: [1, 2], pages: 2, truncated: false, deltaLink: dl });
    expect(calls[0]?.url).toBe(`${G}/v1.0/me/mailFolders/inbox/messages/delta()`);
    // Later incremental sync: the stored deltaLink is validated and used.
    const next = await client.fetchLink(DELTA, dl);
    expect(next.status).toBe(200);
    expect(calls[2]?.url).toBe(dl);
  });

  it("rejects a deltaLink for another endpoint or host without sending the bearer", async () => {
    const { client, calls } = setup(() => json({}));
    for (const link of [
      `${G}/v1.0/me/events/delta?$deltatoken=x`,
      `${G}/v1.0/me/mailFolders/inbox/messages?$deltatoken=x`,
      "https://evil.example/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=x",
    ]) {
      expect((await fail(client.fetchLink(DELTA, link))).code).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });

  it("maps 410 Gone on a delta endpoint to restart_required, not on others", async () => {
    const gone = (): Response =>
      json({ error: { code: "syncStateNotFound", message: "gone" } }, { status: 410 });
    const { client } = setup(gone);
    const err = await fail(client.request(DELTA));
    expect(err.code).toBe("restart_required");
    expect(err.details?.["graphCode"]).toBe("syncStateNotFound");
    expect((await fail(client.request(MSGS))).code).toBe("graph_error");
  });

  it("supports the delta() pattern form (links without the parentheses)", async () => {
    const { client } = setup(() =>
      json({ value: [], "@odata.nextLink": `${G}/v1.0/me/calendarView/delta?$skiptoken=t` }),
    );
    const r = await client.request({ pathPattern: "/me/calendarView/delta()" });
    await expect(
      client.fetchLink({ pathPattern: "/me/calendarView/delta()" }, r.nextLink as string),
    ).resolves.toBeTruthy();
  });

  it("refuses the removed sites delta routes (D36: delegated access not supported)", async () => {
    const { client, calls } = setup(() => json({ value: [] }));
    for (const pathPattern of ["/sites/delta()", "/sites/delta"]) {
      expect((await fail(client.request({ pathPattern }))).code, pathPattern).toBe("invalid_input");
    }
    expect(calls).toHaveLength(0);
  });
});

describe("error normalization", () => {
  const body = { error: { code: "SomeCode", message: "Something broke" } };
  const cases: [number, string][] = [
    [400, "graph_error"],
    [401, "reauth_required"],
    [403, "capability_gap"],
    [404, "graph_error"],
    [410, "graph_error"],
    [429, "rate_limited"],
    [500, "graph_error"],
    [502, "graph_error"],
  ];
  it.each(cases)("HTTP %i -> %s", async (status, expected) => {
    const { client } = setup(
      () =>
        json(body, { status, headers: { "request-id": "req-1", "client-request-id": "cli-1" } }),
      { maxRetries: 0 },
    );
    const err = await fail(client.request(MSGS));
    expect(err.code).toBe(expected);
    expect(err.details).toMatchObject({
      endpoint: "GET /me/messages",
      status,
      graphCode: "SomeCode",
      requestId: "req-1",
      clientRequestId: "cli-1",
    });
  });

  it("truncates Graph text and never echoes non-JSON bodies", async () => {
    const long = setup(() =>
      json({ error: { code: "x".repeat(500), message: "m".repeat(5000) } }, { status: 500 }),
    );
    const e = await fail(long.client.request(MSGS));
    expect(String(e.details?.["graphMessage"]).length).toBeLessThan(250);
    expect(String(e.details?.["graphCode"]).length).toBeLessThan(80);
    const html = setup(() => new Response("<html>SECRET PAGE</html>", { status: 500 }));
    expect(JSON.stringify(await fail(html.client.request(MSGS)))).not.toContain("SECRET PAGE");
  });

  it("maps TokenError codes to reauth_required preserving the provider message", async () => {
    for (const code of ["reauth_required", "auth_misconfigured"] as const) {
      const calls: string[] = [];
      const client = new GraphClient({
        catalog,
        tokenProvider: {
          getAccessToken: () => Promise.reject(new TokenError(code, "fixed provider message")),
        },
        fetch: (url) => {
          calls.push(url);
          return Promise.resolve(json({}));
        },
      });
      const err = await fail(client.request(MSGS));
      expect(err.code).toBe("reauth_required");
      expect(err.message).toBe("fixed provider message");
      expect(calls).toHaveLength(0);
    }
  });

  it("maps unexpected provider failures to graph_error without leaking their message", async () => {
    const client = new GraphClient({
      catalog,
      tokenProvider: { getAccessToken: () => Promise.reject(new Error(`boom ${TOKEN}`)) },
      fetch: () => Promise.reject(new Error("must not be called")),
    });
    const err = await fail(client.request(MSGS));
    expect(err.code).toBe("graph_error");
    expect(JSON.stringify(err)).not.toContain(TOKEN);
  });

  it("rejects header-unsafe tokens before any network call", async () => {
    let called = false;
    const client = new GraphClient({
      catalog,
      tokenProvider: { getAccessToken: () => Promise.resolve("abc\r\nX-Evil: 1") },
      fetch: () => {
        called = true;
        return Promise.resolve(json({}));
      },
    });
    expect((await fail(client.request(MSGS))).code).toBe("reauth_required");
    expect(called).toBe(false);
  });

  it("classifies non-ASCII tokens as reauth_required before any network call", async () => {
    let called = false;
    const client = new GraphClient({
      catalog,
      tokenProvider: { getAccessToken: () => Promise.resolve("abcĀdef") },
      fetch: () => {
        called = true;
        return Promise.resolve(json({}));
      },
    });
    expect((await fail(client.request(MSGS))).code).toBe("reauth_required");
    expect(called).toBe(false);
  });

  it("bounds the wait on a token provider that never settles", async () => {
    let called = false;
    const client = new GraphClient({
      catalog,
      timeoutMs: 20,
      tokenProvider: { getAccessToken: () => new Promise<string>(() => {}) },
      fetch: () => {
        called = true;
        return Promise.resolve(json({}));
      },
    });
    expect((await fail(client.request(MSGS))).code).toBe("reauth_required");
    expect(called).toBe(false);
  });

  it("produces a valid envelope result via toResult()", async () => {
    const { client } = setup(() => json(body, { status: 403 }));
    const result = (await fail(client.request(MSGS))).toResult();
    expect(result.isError).toBe(true);
    expect(result.structuredContent).toMatchObject({
      ok: false,
      error: { code: "capability_gap" },
    });
  });
});

describe("token exfiltration", () => {
  it("the token never appears in errors, details, or JSON across failure modes, and goes only to the fixed origin", async () => {
    const echo = (c: Call): string => authOf(c); // hostile server echoing the header
    const scenarios: [
      string,
      (c: Call, n: number) => Response | Promise<Response>,
      GraphRequest?,
    ][] = [
      [
        "401",
        (c) =>
          json(
            { error: { code: "InvalidAuthenticationToken", message: echo(c) } },
            { status: 401 },
          ),
      ],
      [
        "403",
        (c) => json({ error: { code: echo(c), message: `denied ${TOKEN}` } }, { status: 403 }),
      ],
      ["500 echo", (c) => json({ error: { message: `${TOKEN} ${echo(c)}` } }, { status: 500 })],
      [
        "429",
        () =>
          json(
            { error: { message: TOKEN } },
            { status: 429, headers: { "retry-after": `${TOKEN}` } },
          ),
      ],
      [
        "redirect",
        () =>
          new Response(null, {
            status: 307,
            headers: { location: `https://evil.example/?t=${TOKEN}` },
          }),
      ],
      [
        "network",
        () => {
          throw new Error(`connect failed ${TOKEN}`);
        },
      ],
      ["bad json", () => new Response(`not json ${TOKEN}`, { status: 200 })],
      ["request-id echo", () => json({}, { status: 500, headers: { "request-id": TOKEN } })],
    ];
    for (const [name, responder, req] of scenarios) {
      const { client, calls } = setup(responder, { maxRetries: 0 });
      const err = await fail(client.request(req ?? MSGS));
      const everything = JSON.stringify([err.message, err.details, err, err.toResult(), err.stack]);
      expect(everything, name).not.toContain(TOKEN);
      expect(everything, name).not.toMatch(/authorization/i);
      for (const c of calls) expect(c.url.startsWith(`${G}/`)).toBe(true);
      for (const c of calls) expect(c.url).not.toContain(TOKEN);
    }
  });

  it("the token never appears in successful results", async () => {
    const { client } = setup(() => json({ value: [{ a: 1 }] }));
    const res = await client.request(MSGS);
    expect(JSON.stringify(res)).not.toContain(TOKEN);
  });

  it.each([
    ["a nested value", { value: [{ body: { content: `echo Bearer ${TOKEN}` } }] }],
    ["an object key", { value: [{ [TOKEN]: 1 }] }],
    ["the nextLink", { value: [], "@odata.nextLink": `${G}/v1.0/me/messages?t=${TOKEN}` }],
    ["the deltaLink", { value: [], "@odata.deltaLink": `${G}/v1.0/me/messages?t=${TOKEN}` }],
  ])("refuses a 2xx body that echoes the token in %s", async (_where, body) => {
    const { client } = setup(() => json(body));
    const err = await fail(client.request(MSGS));
    expect(err.code).toBe("graph_error");
    expect(JSON.stringify([err.message, err.details, err.toResult()])).not.toContain(TOKEN);
  });

  it("does not use console in runtime code", async () => {
    const { readFile } = await import("node:fs/promises");
    for (const f of ["client.ts", "errors.ts", "url.ts"]) {
      const src = await readFile(new URL(`./${f}`, import.meta.url), "utf8");
      expect(src).not.toMatch(/console\./);
    }
  });
});
