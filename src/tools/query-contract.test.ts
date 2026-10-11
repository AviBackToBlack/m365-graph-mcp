import { describe, expect, it } from "vitest";
import { GraphClientError } from "../graph/errors.js";
import { defineQueryContract, kqlQuote, kqlUnquote } from "./query-contract.js";

const contract = defineQueryContract({
  top: { default: 10, max: 200, pageMax: 50 },
  select: {
    allowed: ["id", "subject", "bodyPreview"],
    default: ["id", "subject"],
    requiredInBatch: true,
  },
  filter: { maxLength: 100 },
  search: { maxLength: 30 },
  orderby: { allowed: ["receivedDateTime", "subject", "from/emailAddress/address"], maxClauses: 3 },
  skip: { max: 100 },
  includeHiddenFolders: true,
});

const policy = (
  p: ReadonlyArray<readonly [string, string]>,
  headers: Record<string, string> = {},
): void => contract.batchPolicy.check(p, headers);
const rejects = (fn: () => unknown): void => {
  try {
    fn();
  } catch (err) {
    expect(err).toBeInstanceOf(GraphClientError);
    expect((err as GraphClientError).code).toBe("invalid_input");
    return;
  }
  throw new Error("expected invalid_input");
};

describe("kqlQuote / kqlUnquote", () => {
  it("wraps in one pair of double quotes and escapes backslash and quote", () => {
    expect(kqlQuote("from:john")).toBe('"from:john"');
    expect(kqlQuote('subject:"q3 plan"')).toBe('"subject:\\"q3 plan\\""');
    expect(kqlQuote("a\\b")).toBe('"a\\\\b"');
  });

  it("round-trips arbitrary text and refuses malformed wire forms", () => {
    for (const text of ['x"y', "x\\y", 'trail\\"', "\\\\", "plain words", 'a" OR from:b']) {
      expect(kqlUnquote(kqlQuote(text))).toBe(text);
    }
    for (const wire of [
      "pizza",
      '"pizza',
      'pizza"',
      '"a"b"',
      '"a\\"',
      '"a\\xb"',
      '"',
      "",
      '"a" OR "b"',
    ]) {
      expect(kqlUnquote(wire)).toBeUndefined();
    }
  });
});

describe("toRequest", () => {
  it("sends at most pageMax per request while limit keeps the requested top", () => {
    expect(contract.toRequest({ top: 200 })).toMatchObject({
      query: { $top: 50 },
      limit: 200,
    });
    expect(contract.toRequest({ top: 30 })).toMatchObject({ query: { $top: 30 }, limit: 30 });
    expect(contract.toRequest({}).query["$top"]).toBe(10);
  });

  it("applies the compact default select and default top", () => {
    const r = contract.toRequest({});
    expect(r.query).toEqual({ $top: 10, $select: "id,subject" });
    expect(r.limit).toBe(10);
    expect(r.headers).toEqual({});
  });

  it("builds every option and treats filter/search as data", () => {
    const r = contract.toRequest({
      top: 50,
      select: ["subject", "subject", "id"],
      filter: "isRead eq false",
      skip: 5,
      includeHiddenFolders: true,
    });
    expect(r.query).toEqual({
      $top: 50,
      $select: "subject,id",
      $filter: "isRead eq false",
      $skip: 5,
      includeHiddenFolders: true,
    });
    const s = contract.toRequest({ search: 'from:a" OR from:b' });
    expect(s.query["$search"]).toBe('"from:a\\" OR from:b"');
  });

  it("does not strip an enclosing quote pair: the text is searched as a phrase", () => {
    expect(contract.toRequest({ search: '"quarterly report"' }).query["$search"]).toBe(
      '"\\"quarterly report\\""',
    );
  });

  it("rejects out-of-contract options, bounds and forbidden combinations", () => {
    const bad: Record<string, unknown>[] = [
      { top: 201 },
      { top: 0 },
      { top: 1.5 },
      { top: "5" },
      { skip: 101 },
      { skip: -1 },
      { select: ["contentBytes"] },
      { filter: "isRead eq false", orderby: [{ property: "receivedDateTime" }] },
      { select: [] },
      { select: "id" },
      { orderby: [{ property: "bodyPreview" }] },
      { orderby: [{ property: "subject", direction: "up" }] },
      {
        orderby: [
          { property: "subject" },
          { property: "subject" },
          { property: "subject" },
          { property: "subject" },
        ],
      },
      { filter: "x".repeat(101) },
      { filter: "a\nb" },
      { search: "x".repeat(31) },
      { search: "   " },
      { search: '"' },
      { search: '""' },
      { search: "a\u0000b" },
      { search: "x", orderby: [{ property: "subject" }] },
      { search: "x", skip: 1 },
      { search: "x", filter: "isRead eq true" },
    ];
    for (const args of bad) rejects(() => contract.toRequest(args));
  });

  it("checks the filter/orderby rule locally without echoing values", () => {
    const ok = (filter: string, orderby: Array<{ property: string; direction?: string }>) =>
      contract.toRequest({ filter, orderby });
    const rcv = { property: "receivedDateTime", direction: "desc" };
    rejects(() => ok("isRead eq false", [rcv]));
    rejects(() => ok("subject eq 'receivedDateTime'", [rcv])); // literal is not a property
    rejects(() => ok("xreceivedDateTime ge 1", [rcv])); // not a whole token
    rejects(() => ok("receivedDateTime2 ge 1", [rcv]));
    rejects(() => ok("receivedDateTime ge 1 and subject eq 'x'", [{ property: "subject" }, rcv]));
    expect(() => ok("receivedDateTime ge 1 and isRead eq false", [rcv])).not.toThrow();
    expect(() =>
      ok("receivedDateTime ge 1 and subject eq 'x'", [rcv, { property: "subject" }]),
    ).not.toThrow();
    expect(() =>
      ok("from/emailAddress/address eq 'a@b.c'", [{ property: "from/emailAddress/address" }]),
    ).not.toThrow();
    try {
      ok("SECRETFILTER eq 1", [rcv]);
      throw new Error("expected throw");
    } catch (err) {
      const text = JSON.stringify((err as GraphClientError).toJSON());
      expect(text).not.toContain("SECRETFILTER");
      expect(text).toContain("orderby");
    }
  });

  it("never echoes offending values or unknown key names in errors", () => {
    try {
      contract.toRequest({
        filter: "x".repeat(101),
        "sneaky-key-name": 1,
        select: ["SECRET_VALUE"],
      });
      throw new Error("expected throw");
    } catch (err) {
      const text = JSON.stringify((err as GraphClientError).toJSON());
      expect(text).not.toContain("sneaky-key-name");
      expect(text).not.toContain("SECRET_VALUE");
      expect(text).not.toContain("xxxxxxxx");
    }
  });

  it("sends the body-content-type Prefer header only for tools that declare bodyType", () => {
    const withBody = defineQueryContract({
      select: { allowed: ["id", "body"], default: ["id"], requiredInBatch: false },
      bodyType: { default: "text" },
    });
    expect(withBody.toRequest({}).headers).toEqual({ Prefer: 'outlook.body-content-type="text"' });
    expect(withBody.toRequest({ bodyType: "html" }).headers).toEqual({
      Prefer: 'outlook.body-content-type="html"',
    });
    rejects(() => withBody.toRequest({ bodyType: "markdown" }));
    // Undeclared options are never turned into a query (unknown keys are rejected one layer
    // up by the strict tool schema, covered in mail.test.ts).
    expect(withBody.toRequest({ top: 5, expand: "x" }).query).toEqual({ $select: "id" });
    expect(contract.toRequest({}).headers).toEqual({});
  });

  it("advertises allowlists in the typed shape", () => {
    expect(Object.keys(contract.shape).sort()).toEqual(
      ["filter", "includeHiddenFolders", "orderby", "search", "select", "skip", "top"].sort(),
    );
  });
});

describe("batchPolicy (same contract on decoded wire pairs)", () => {
  it("accepts everything toRequest emits (round trip)", () => {
    const requests = [
      { select: ["id"] },
      { top: 50, select: ["subject", "id"], filter: "isRead eq false", skip: 5 },
      {
        select: ["id"],
        filter: "receivedDateTime ge 1 and isRead eq false",
        orderby: [{ property: "receivedDateTime", direction: "desc" }],
      },
      { select: ["id"], search: 'weird "quote" and \\ slash' },
      {
        select: ["id"],
        orderby: [{ property: "receivedDateTime", direction: "desc" }, { property: "subject" }],
      },
      { select: ["id"], includeHiddenFolders: false },
    ];
    for (const args of requests) {
      const { query } = contract.toRequest(args);
      const wire = Object.entries(query).map(([k, v]) => [k, String(v)] as [string, string]);
      expect(() => policy(wire)).not.toThrow();
    }
  });

  it("enforces the same bounds and allowlists", () => {
    const sel: [string, string] = ["$select", "id"];
    const bad: [string, string][][] = [
      [sel, ["$top", "51"]], // above pageMax: batch cannot paginate
      [sel, ["$top", "200"]],
      [sel, ["$top", "201"]],
      [sel, ["$filter", "isRead eq false"], ["$orderby", "receivedDateTime desc"]],
      [
        sel,
        ["$filter", "subject eq 'x' and receivedDateTime ge 1"],
        ["$orderby", "receivedDateTime,subject"],
      ],
      [sel, ["$top", "0"]],
      [sel, ["$top", "05"]],
      [sel, ["$top", "+5"]],
      [sel, ["$top", "1e1"]],
      [sel, ["$top", "5.0"]],
      [sel, ["$top", "5x"]],
      [sel, ["$skip", "101"]],
      [sel, ["$skip", "-1"]],
      [["$select", "id,contentBytes"]],
      [["$select", "id, subject"]],
      [["$select", ""]],
      [["$select", "id,,subject"]],
      [sel, ["$orderby", "bodyPreview"]],
      [sel, ["$orderby", "subject up"]],
      [sel, ["$orderby", "subject,subject,subject,subject"]],
      [sel, ["$orderby", "subject  desc"]],
      [sel, ["$filter", "x".repeat(101)]],
      [sel, ["$filter", "a\nb"]],
      [sel, ["$search", "x"]],
      [sel, ["$search", '"x" OR "y"']],
      [sel, ["$search", '"' + "x".repeat(31) + '"']],
      [sel, ["$search", '""']],
      [sel, ["includeHiddenFolders", "TRUE"]],
    ];
    for (const p of bad) rejects(() => policy(p));
  });

  it("rejects unknown, $-less, re-cased, duplicate and expand-style names", () => {
    const sel: [string, string] = ["$select", "id"];
    for (const name of [
      "top",
      "select",
      "$TOP",
      "$Select",
      "$expand",
      "expand",
      "$skiptoken",
      "$count",
      "$format",
      "$apply",
      "access_token",
      "$top ",
    ]) {
      rejects(() => policy([sel, [name, "1"]]));
    }
    rejects(() => policy([sel, ["$top", "5"], ["$top", "5000"]]));
    rejects(() => policy([sel, sel]));
  });

  it("projects items and collections to the select allowlist only", () => {
    const item = {
      id: "1",
      subject: "s",
      "@odata.type": "t",
      "@odata.etag": "e",
      "@odata.nextLink": "n",
      body: "SECRET",
    };
    expect(contract.projectItem(item)).toEqual({
      id: "1",
      subject: "s",
      "@odata.type": "t",
      "@odata.etag": "e",
    });
    expect(contract.projectItem("x")).toBeUndefined();
    expect(contract.projectItem([item])).toBeUndefined();
    const projectQ = contract.batchPolicy.project as (
      b: unknown,
      q: ReadonlyArray<readonly [string, string]>,
    ) => unknown;
    const project = (b: unknown) => projectQ(b, [["$top", "5"]]);
    expect(
      project({
        "@odata.context": "c",
        "@odata.nextLink": "n",
        extra: 1,
        value: [item, "junk", null],
      }),
    ).toEqual({
      value: [{ id: "1", subject: "s", "@odata.type": "t", "@odata.etag": "e" }],
      hasMore: true, // the nextLink is dropped but signals more results
    });
    expect(project({ value: "nope" })).toEqual({});
    expect(project([1])).toEqual({});
    expect(project(null)).toEqual({});
  });

  it("clips batch collections to the requested $top, or the default when omitted", () => {
    const projectQ = contract.batchPolicy.project as (
      b: unknown,
      q: ReadonlyArray<readonly [string, string]>,
    ) => unknown;
    const items = Array.from({ length: 300 }, (_, i) => ({ id: String(i), subject: "s" }));
    const one = projectQ({ value: items.slice(0, 3) }, [["$top", "1"]]) as {
      value: unknown[];
      hasMore?: boolean;
    };
    expect(one.value).toHaveLength(1);
    expect(one.hasMore).toBe(true);
    const exact = projectQ({ value: items.slice(0, 2) }, [["$top", "2"]]) as {
      hasMore?: boolean;
    };
    expect(exact.hasMore).toBeUndefined();
    const link = { "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=2" };
    const fullPage = projectQ({ value: items.slice(0, 2), ...link }, [["$top", "2"]]) as {
      value: unknown[];
      hasMore?: boolean;
    };
    expect(fullPage.value).toHaveLength(2);
    expect(fullPage.hasMore).toBe(true);
    expect(JSON.stringify(fullPage)).not.toContain("nextLink");
    const shortPage = projectQ({ value: items.slice(0, 1), ...link }, [["$top", "5"]]) as {
      value: unknown[];
      hasMore?: boolean;
    };
    expect(shortPage.value).toHaveLength(1);
    expect(shortPage.hasMore).toBe(true);
    const dflt = projectQ({ value: items }, []) as { value: unknown[]; hasMore?: boolean };
    expect(dflt.value).toHaveLength(10);
    expect(dflt.hasMore).toBe(true);
  });

  it("requires the body-content-type Prefer in a batch for body tools", () => {
    const withBody = defineQueryContract({
      select: { allowed: ["id", "body"], default: ["id"], requiredInBatch: false },
      bodyType: { default: "text" },
    });
    const prefer = (type: string) => ({ Prefer: `outlook.body-content-type="${type}"` });
    expect(() => withBody.batchPolicy.check([], prefer("text"))).not.toThrow();
    expect(() => withBody.batchPolicy.check([], prefer("html"))).not.toThrow();
    rejects(() => withBody.batchPolicy.check([], {}));
    rejects(() => withBody.batchPolicy.check([["$select", "id"]], {}));
    expect(withBody.batchPolicy.project).toBeUndefined();
  });

  it("applies the search combination rules and requires $select where declared", () => {
    const sel: [string, string] = ["$select", "id"];
    const q: [string, string] = ["$search", '"x"'];
    rejects(() => policy([sel, q, ["$orderby", "subject"]]));
    rejects(() => policy([sel, q, ["$skip", "1"]]));
    rejects(() => policy([sel, q, ["$filter", "isRead eq true"]]));
    rejects(() => policy([]));
    rejects(() => policy([["$top", "5"]]));
    expect(() => policy([sel, q, ["$top", "5"]])).not.toThrow();
  });
});

describe("window / timeZone / requireInFilter extensions", () => {
  const cal = defineQueryContract({
    top: { default: 10, max: 100 },
    select: { allowed: ["id", "subject", "body"], default: ["id"], requiredInBatch: false },
    filter: { maxLength: 100 },
    orderby: { allowed: ["start/dateTime"], maxClauses: 1, requireInFilter: false },
    bodyType: { default: "text" },
    timeZone: true,
    window: { maxDays: 7 },
  });
  const W = { startDateTime: "2024-01-01T00:00:00Z", endDateTime: "2024-01-08T00:00:00Z" };
  const wire = [
    ["startDateTime", W.startDateTime],
    ["endDateTime", W.endDateTime],
  ] as const;

  it("builds window query values and one combined Prefer (timezone first)", () => {
    const req = cal.toRequest({ ...W, timeZone: "europe/berlin", bodyType: "html" });
    expect(req.query["startDateTime"]).toBe(W.startDateTime);
    expect(req.query["endDateTime"]).toBe(W.endDateTime);
    expect(req.headers).toEqual({
      Prefer: 'outlook.timezone="Europe/Berlin", outlook.body-content-type="html"',
    });
    expect(cal.toRequest({ ...W }).headers).toEqual({
      Prefer: 'outlook.body-content-type="text"',
    });
  });

  it("enforces required, ordered and bounded windows on both paths", () => {
    rejects(() => cal.toRequest({}));
    rejects(() => cal.toRequest({ startDateTime: W.startDateTime }));
    rejects(() => cal.toRequest({ startDateTime: W.endDateTime, endDateTime: W.startDateTime }));
    rejects(() =>
      cal.toRequest({ startDateTime: W.startDateTime, endDateTime: "2024-01-08T00:00:01Z" }),
    );
    const p = (q: ReadonlyArray<readonly [string, string]>) =>
      cal.batchPolicy.check(q, { Prefer: 'outlook.body-content-type="text"' });
    p(wire);
    rejects(() => p([]));
    rejects(() => p([wire[0]]));
    rejects(() => p([...wire, wire[0]]));
    rejects(() => p([wire[0], ["endDateTime", "2024-01-08T00:00:01Z"]]));
    rejects(() => p([wire[0], ["endDateTime", "2024-01-08T00:00:00+00:00"]]));
  });

  it("batch Prefer: canonical combined form only", () => {
    const ok = (prefer: string) => cal.batchPolicy.check(wire, { Prefer: prefer });
    ok('outlook.body-content-type="text"');
    ok('outlook.timezone="UTC", outlook.body-content-type="html"');
    for (const bad of [
      'outlook.timezone="UTC"',
      'outlook.body-content-type="text", outlook.timezone="UTC"',
      'outlook.timezone="UTC", outlook.body-content-type="text", outlook.allow-unsafe-html',
      'outlook.timezone="utc", outlook.body-content-type="text"',
      'outlook.timezone="Nowhere/Land", outlook.body-content-type="text"',
    ]) {
      rejects(() => ok(bad));
    }
    rejects(() => cal.batchPolicy.check(wire, {}));
    rejects(() =>
      cal.batchPolicy.check(wire, { Prefer: 'outlook.body-content-type="text"', X: "1" }),
    );
  });

  it("requireInFilter:false allows a filter that omits the orderby property", () => {
    cal.toRequest({
      ...W,
      filter: "isAllDay eq false",
      orderby: [{ property: "start/dateTime" }],
    });
    // Default (mail) contract keeps the Graph messages rule.
    rejects(() =>
      contract.toRequest({
        filter: "isRead eq false",
        orderby: [{ property: "receivedDateTime" }],
      }),
    );
  });
});

describe("search.siteSearch (GET /sites?search=)", () => {
  const sites = defineQueryContract({
    top: { default: 25, max: 200, wire: false },
    search: { maxLength: 50, siteSearch: true },
  });
  const check = (p: ReadonlyArray<readonly [string, string]>, h: Record<string, string> = {}) =>
    sites.batchPolicy.check(p, h);

  it("sends the bare `search` name with the verbatim text and never `$search`", () => {
    const { query } = sites.toRequest({ search: 'a "b" \\ c' });
    expect(query).toEqual({ search: 'a "b" \\ c' });
    expect(Object.keys(query)).not.toContain("$search");
  });

  it("requires search (typed and in batch) and refuses text without a letter or digit", () => {
    rejects(() => sites.toRequest({}));
    rejects(() => sites.toRequest({ top: 5 }));
    for (const search of ["", " ", "*", "?*", "a\nb", "x".repeat(51)]) {
      rejects(() => sites.toRequest({ search }));
      rejects(() => check([["search", search]]));
    }
    rejects(() => check([]));
    rejects(() => check([["$top", "5"]]));
  });

  it("batch accepts only the bare `search` pair, once, and nothing else", () => {
    check([["search", "hr team"]]);
    rejects(() => check([["$search", '"hr"']]));
    rejects(() => check([["$search", "hr"]]));
    rejects(() =>
      check([
        ["search", "hr"],
        ["search", "hr"],
      ]),
    );
    rejects(() =>
      check([
        ["search", "hr"],
        ["$top", "5"],
      ]),
    );
    rejects(() =>
      check([
        ["search", "hr"],
        ["$select", "id"],
      ]),
    );
    rejects(() => check([["Search", "hr"]]));
    rejects(() => check([["search", "hr"]], { ConsistencyLevel: "eventual" }));
  });

  it("is not available together with plainName or directoryProperties", () => {
    expect(() =>
      defineQueryContract({ search: { maxLength: 10, siteSearch: true, plainName: true } }),
    ).toThrow(/siteSearch/);
    expect(() =>
      defineQueryContract({
        search: { maxLength: 10, siteSearch: true, directoryProperties: ["displayName"] },
      }),
    ).toThrow(/siteSearch/);
  });

  it("leaves every other contract on the quoted $search form", () => {
    expect(contract.toRequest({ search: "from:john" }).query["$search"]).toBe('"from:john"');
    rejects(() => policy([["search", "from:john"]]));
  });
});
