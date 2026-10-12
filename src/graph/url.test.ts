import { describe, expect, it } from "vitest";
import type { EndpointDefinition } from "../catalog.js";
import { GraphClientError } from "./errors.js";
import { buildPath, buildQuery, buildUrl, validateGraphLink } from "./url.js";

const ep = (pathPattern: string, extra: Record<string, unknown> = {}): EndpointDefinition => ({
  pathPattern,
  method: "get",
  toolName: "t",
  ...extra,
});

const MESSAGES = ep("/me/mailFolders/{mailFolder-id}/messages");
const DELTA = ep("/me/mailFolders/{mailFolder-id}/messages/delta()");
const EVENTS_DELTA = ep("/me/events/delta()");
const INBOX = { "mailFolder-id": "inbox" };
const GUID_A = "f17e0000-0000-4000-8000-000000000004";
const GUID_B = "f17e0000-0000-4000-8000-000000000005";
const SITE_PATH = ep("/sites/{site-id}:/{path}", { skipEncoding: ["path"] });

function code(fn: () => unknown): string {
  try {
    fn();
  } catch (e) {
    if (e instanceof GraphClientError) return e.code;
    throw e;
  }
  return "no-error";
}

describe("buildPath", () => {
  it("encodes each parameter as one path segment", () => {
    expect(buildPath(MESSAGES, { "mailFolder-id": "AAMk+a=b-c_d" })).toBe(
      "/me/mailFolders/AAMk%2Ba=b-c_d/messages",
    );
    expect(buildPath(MESSAGES, INBOX)).toBe("/me/mailFolders/inbox/messages");
  });

  it.each(["", ".", ".."])("rejects %j", (value) => {
    expect(code(() => buildPath(MESSAGES, { "mailFolder-id": value }))).toBe("invalid_input");
  });

  it("rejects separator, traversal and backslash input (no id contains them)", () => {
    for (const value of ["../x", "a\\b", "a/b", "a?b", "a#b", "a b", "AAMk1/x"]) {
      expect(
        code(() => buildPath(MESSAGES, { "mailFolder-id": value })),
        value,
      ).toBe("invalid_input");
    }
  });

  it("rejects missing and unexpected params", () => {
    expect(code(() => buildPath(MESSAGES, {}))).toBe("invalid_input");
    expect(code(() => buildPath(MESSAGES, { "mailFolder-id": "inbox", extra: "y" }))).toBe(
      "invalid_input",
    );
  });

  it("splits only catalog skipEncoding params and validates each segment", () => {
    expect(
      buildPath(SITE_PATH, { "site-id": "contoso.sharepoint.com", path: "sites/hr team" }),
    ).toBe("/sites/contoso.sharepoint.com:/sites/hr%20team");
    const host = "contoso.sharepoint.com";
    for (const bad of [
      "sites/../x",
      "/sites/x",
      "sites//x",
      "sites/x/",
      "..",
      "sites/x:/drive",
      "sites/%2e%2e/x",
      "sites/%252e%252e/x",
      "%2e",
    ]) {
      expect(code(() => buildPath(SITE_PATH, { "site-id": host, path: bad }))).toBe(
        "invalid_input",
      );
    }
    expect(code(() => buildPath(SITE_PATH, { "site-id": "a/b", path: "x" }))).toBe("invalid_input");
  });

  it.each(["x:y", "x%3Ay", "x%3ay", "x%253Ay", "x%25253Ay", "sites/a%3A/drive", "%3A", "x%zz"])(
    "rejects an encoded or raw colon / bad escape in {path}: %s",
    (bad) => {
      expect(
        code(() => buildPath(SITE_PATH, { "site-id": "contoso.sharepoint.com", path: bad })),
      ).toBe("invalid_input");
    },
  );

  it.each([
    "sites/a?x=1",
    "sites/a%3Fb",
    "sites/a#b",
    "sites/a%23b",
    "sites\\a",
    "sites/a%5Cb",
    "sites/a%255Cb",
    "sites/a\nb",
    "sites/a\u0000b",
    "sites/a%00b",
    "sites/a\u007fb",
  ])("rejects a query/fragment/backslash/control character in {path}: %j", (bad) => {
    expect(
      code(() => buildPath(SITE_PATH, { "site-id": "contoso.sharepoint.com", path: bad })),
    ).toBe("invalid_input");
  });

  it.each(["sites/hr%2Fdrive", "sites/hr%2fdrive", "sites/hr%252Fdrive", "sites%2Fhr"])(
    "rejects an encoded '/' in {path} (batch parity): %j",
    (bad) => {
      expect(
        code(() => buildPath(SITE_PATH, { "site-id": "contoso.sharepoint.com", path: bad })),
      ).toBe("invalid_input");
    },
  );

  it("requires a plain hostname as {site-id} before ':/' (this pattern only)", () => {
    const ok = "contoso.sharepoint.com";
    expect(buildPath(SITE_PATH, { "site-id": ok, path: "sites/hr" })).toBe(
      "/sites/contoso.sharepoint.com:/sites/hr",
    );
    for (const id of ["root", GUID_A, `${GUID_A},${GUID_B}`, `${ok},${GUID_A},${GUID_B}`, "host"]) {
      expect(
        code(() => buildPath(SITE_PATH, { "site-id": id, path: "sites/hr" })),
        id,
      ).toBe("invalid_input");
    }
    // The shared site-id shape (root / composite / GUID) is unchanged for every other site tool.
    for (const id of ["root", GUID_A, `${GUID_A},${GUID_B}`, `${ok},${GUID_A},${GUID_B}`]) {
      expect(buildPath(ep("/sites/{site-id}"), { "site-id": id })).toBe(
        `/sites/${encodeURIComponent(id)}`,
      );
    }
  });

  it("rejects a non-hostname {site-id} in a {path} continuation link too", () => {
    const G = "https://graph.microsoft.com";
    for (const id of ["root", `${GUID_A},${GUID_B}`, `h.sharepoint.com,${GUID_A},${GUID_B}`]) {
      expect(
        code(() =>
          validateGraphLink(
            `${G}/v1.0/sites/${encodeURIComponent(id)}:/sites/hr`,
            SITE_PATH,
            "v1.0",
            { "site-id": id, path: "sites/hr" },
          ),
        ),
        id,
      ).toBe("invalid_input");
    }
  });

  it("rejects OData control segments inside {path}", () => {
    for (const bad of ["sites/$value", "$count", "sites/microsoft.graph.x", "sites/a(1)"]) {
      expect(
        code(() => buildPath(SITE_PATH, { "site-id": "contoso.sharepoint.com", path: bad })),
        bad,
      ).toBe("invalid_input");
    }
  });
});

describe("buildQuery / buildUrl", () => {
  it("keeps $ names literal and encodes values", () => {
    expect(
      buildQuery({ $top: 5, $select: "a,b", $filter: "x eq 'y z'&q=1", skip: undefined }),
    ).toBe("$top=5&$select=a%2Cb&$filter=x%20eq%20'y%20z'%26q%3D1");
  });

  it("rejects hostile query names and non-finite numbers", () => {
    expect(code(() => buildQuery({ "a=b&c": "x" }))).toBe("invalid_input");
    expect(code(() => buildQuery({ "": "x" }))).toBe("invalid_input");
    expect(code(() => buildQuery({ $top: Number.NaN }))).toBe("invalid_input");
  });

  it.each([
    ["array", ["id", "subject"]],
    ["null", null],
    ["object", { id: "subject" }],
    ["bigint", 7n],
    ["symbol", Symbol("injected")],
    ["function", () => "id"],
  ])("rejects non-scalar runtime query value %s rather than coercing it", (_label, value) => {
    // Compile-time QueryValue cannot enforce safety for untrusted runtime inputs.
    const query = { $select: value } as unknown as Parameters<typeof buildQuery>[0];
    expect(code(() => buildQuery(query))).toBe("invalid_input");
  });

  it.each([
    ["null", null],
    ["array", []],
    ["string", "query"],
    ["number", 7],
    ["Date", new Date("2026-01-01T00:00:00Z")],
  ])("rejects a non-record runtime query %s", (_label, value) => {
    const query = value as unknown as Parameters<typeof buildQuery>[0];
    expect(code(() => buildQuery(query))).toBe("invalid_input");
  });

  it("permits omitted and null-prototype plain query records", () => {
    expect(buildQuery()).toBe("");
    const query = Object.assign(Object.create(null) as Record<string, string>, { $top: "10" });
    expect(buildQuery(query)).toBe("$top=10");
  });

  it("bounds diagnostics for oversized query names without echoing user input", () => {
    const hugeName = "$" + "x".repeat(100_000);
    const query = { [hugeName]: [] } as unknown as Parameters<typeof buildQuery>[0];
    let error: unknown;
    try {
      buildQuery(query);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(GraphClientError);
    expect((error as GraphClientError).code).toBe("invalid_input");
    expect((error as GraphClientError).message.length).toBeLessThan(256);
    expect((error as GraphClientError).message).not.toContain(hugeName);
  });

  it("preserves valid scalar query types and omitted keys", () => {
    expect(buildQuery({ $top: 0, $count: false, $filter: "a eq 'b'", skip: undefined })).toBe(
      "$top=0&$count=false&$filter=a%20eq%20'b'",
    );
  });

  it("builds on the fixed origin with version prefix and caps length", () => {
    expect(buildUrl("v1.0", "/me", { $top: 1 })).toBe("https://graph.microsoft.com/v1.0/me?$top=1");
    expect(buildUrl("beta", "/me/profile", undefined)).toBe(
      "https://graph.microsoft.com/beta/me/profile",
    );
    expect(code(() => buildUrl("v1.0", "/me", { $filter: "x".repeat(9000) }))).toBe(
      "invalid_input",
    );
  });
});

describe("validateGraphLink", () => {
  const G = "https://graph.microsoft.com";

  describe("with the original query: a continuation may not widen it", () => {
    const link = (q: string) => `${G}/v1.0/me/mailFolders/inbox/messages?${q}`;
    const check = (q: string, original: Record<string, string>) =>
      validateGraphLink(link(q), MESSAGES, "v1.0", INBOX, original);

    it("keeps paging state and options the original sent (any order or spelling)", () => {
      const original = { $select: "id,subject", $expand: "attachments($select=id,name)" };
      for (const q of [
        "$skiptoken=abc",
        "$skip=10&$top=10",
        "$deltatoken=t",
        "$select=subject,id&$skiptoken=abc",
        "select=ID&$skiptoken=abc",
        "$expand=attachments($select=id,name)&$select=id",
        "startDateTime=2026-01-01&$skiptoken=abc",
      ]) {
        expect(check(q, original), q).toBe(link(q));
      }
      expect(
        check("$filter=isRead%20eq%20false&$skiptoken=a", { $filter: "isRead eq false" }),
      ).toBe(link("$filter=isRead%20eq%20false&$skiptoken=a"));
    });

    it.each([
      ["$expand=attachments", {}],
      ["expand=attachments", {}],
      ["$EXPAND=attachments", {}],
      ["$select=id", {}],
      ["$filter=x%20eq%201", {}],
      ["$orderby=id", {}],
      ["$search=x", {}],
      ["$count=true", {}],
      ["$apply=x", {}],
      ["$compute=x", {}],
      ["$format=json", {}],
      ["$select=id,body", { $select: "id" }],
      ["$expand=attachments,extensions", { $expand: "attachments" }],
      ["$expand=attachments&$expand=extensions", { $expand: "attachments" }],
      ["$expand=attachments($select=id,contentBytes)", { $expand: "attachments($select=id)" }],
    ])("refuses %j (original %j)", (q, original) => {
      expect(() => check(q, original)).toThrow(GraphClientError);
    });

    it("rejects invalid original query values even when the link only carries paging state", () => {
      const continuation = link("$skip=10");
      for (const value of [[], { id: "subject" }, null, 8n, Symbol("oops")]) {
        const original = { $select: value } as unknown as Parameters<typeof validateGraphLink>[4];
        expect(code(() => validateGraphLink(continuation, MESSAGES, "v1.0", INBOX, original))).toBe(
          "invalid_input",
        );
      }
      const badRecord = new Date("2026-01-01T00:00:00Z") as unknown as Parameters<
        typeof validateGraphLink
      >[4];
      expect(code(() => validateGraphLink(continuation, MESSAGES, "v1.0", INBOX, badRecord))).toBe(
        "invalid_input",
      );
    });

    it("does not compare query options when no original query is given (delta token extraction)", () => {
      expect(validateGraphLink(link("$select=id&$skiptoken=a"), MESSAGES, "v1.0", INBOX)).toBe(
        link("$select=id&$skiptoken=a"),
      );
    });
  });

  it("accepts and rebuilds legitimate nextLinks", () => {
    expect(
      validateGraphLink(
        `${G}/v1.0/me/mailFolders/inbox/messages?$skip=10&$top=10`,
        MESSAGES,
        "v1.0",
        INBOX,
      ),
    ).toBe(`${G}/v1.0/me/mailFolders/inbox/messages?$skip=10&$top=10`);
    expect(
      validateGraphLink(
        `${G}/v1.0/me/mailFolders('inbox')/messages?$skiptoken=abc`,
        MESSAGES,
        "v1.0",
        INBOX,
      ),
    ).toContain("/me/mailFolders('inbox')/messages?$skiptoken=abc");
  });

  it("accepts delta links for the same delta endpoint (with and without ())", () => {
    expect(
      validateGraphLink(
        `${G}/v1.0/me/mailFolders('inbox')/messages/delta?$deltatoken=t`,
        DELTA,
        "v1.0",
        INBOX,
      ),
    ).toContain("$deltatoken=t");
    expect(
      validateGraphLink(`${G}/v1.0/me/events/delta?$deltatoken=t`, EVENTS_DELTA, "v1.0"),
    ).toContain("delta?$deltatoken=t");
    expect(
      validateGraphLink(`${G}/v1.0/me/events/delta()?$skiptoken=t`, EVENTS_DELTA, "v1.0"),
    ).toContain("$skiptoken=t");
  });

  it("accepts site-by-path links", () => {
    expect(
      validateGraphLink(`${G}/v1.0/sites/h.sharepoint.com:/sites/hr`, SITE_PATH, "v1.0", {
        "site-id": "h.sharepoint.com",
        path: "sites/hr",
      }),
    ).toContain("/sites/h.sharepoint.com:/sites/hr");
  });

  const bad: [string, string][] = [
    ["other host", "https://evil.example/v1.0/me/mailFolders/i/messages"],
    ["lookalike suffix", "https://graph.microsoft.com.evil.com/v1.0/me/mailFolders/i/messages"],
    ["lookalike prefix", "https://evilgraph.microsoft.com/v1.0/me/mailFolders/i/messages"],
    ["trailing dot host", "https://graph.microsoft.com./v1.0/me/mailFolders/i/messages"],
    ["http scheme", "http://graph.microsoft.com/v1.0/me/mailFolders/i/messages"],
    ["userinfo", "https://user:pw@graph.microsoft.com/v1.0/me/mailFolders/i/messages"],
    ["userinfo host trick", "https://graph.microsoft.com@evil.com/v1.0/me/mailFolders/i/messages"],
    ["explicit port", "https://graph.microsoft.com:8443/v1.0/me/mailFolders/i/messages"],
    ["explicit default port", "https://graph.microsoft.com:443/v1.0/me/mailFolders/i/messages"],
    ["empty port", "https://graph.microsoft.com:/v1.0/me/mailFolders/i/messages"],
    ["trailing-dot host", "https://graph.microsoft.com./v1.0/me/mailFolders/i/messages"],
    ["different version", "https://graph.microsoft.com/beta/me/mailFolders/i/messages"],
    ["no version", "https://graph.microsoft.com/me/mailFolders/i/messages"],
    ["different endpoint", "https://graph.microsoft.com/v1.0/me/messages"],
    ["different endpoint (admin)", "https://graph.microsoft.com/v1.0/users"],
    ["extra path suffix", "https://graph.microsoft.com/v1.0/me/mailFolders/i/messages/x/y"],
    ["protocol-relative", "//evil.example/v1.0/me/mailFolders/i/messages"],
    ["relative path", "/v1.0/me/mailFolders/i/messages"],
    ["javascript scheme", "javascript:alert(1)"],
    ["file scheme", "file:///etc/passwd"],
    ["backslash host", "https://evil.example\\@graph.microsoft.com/v1.0/me/mailFolders/i/messages"],
    ["backslash path", "https://graph.microsoft.com/v1.0\\..\\me/mailFolders/i/messages"],
    ["encoded slash", "https://graph.microsoft.com/v1.0/me/mailFolders/a%2Fb/messages"],
    ["encoded dots", "https://graph.microsoft.com/v1.0/me/%2e%2e/me/mailFolders/i/messages"],
    ["fragment", "https://graph.microsoft.com/v1.0/me/mailFolders/i/messages#x"],
    ["whitespace", " https://graph.microsoft.com/v1.0/me/mailFolders/i/messages"],
    [
      "newline injection",
      "https://graph.microsoft.com/v1.0/me/mailFolders/i/messages\r\nHost: evil",
    ],
    ["decimal/IP host", "https://2852039166/v1.0/me/mailFolders/i/messages"],
    [
      "unicode lookalike host",
      "https://graph.microsoft.com/v1.0/me/mailFolders/i/messages".replace("graph", "grарh"),
    ],
    ["fullwidth host", "https://ｇraph.microsoft.com.evil.com/v1.0/me/mailFolders/i/messages"],
    ["empty", ""],
  ];
  it.each(bad)("rejects %s", (_name, link) => {
    expect(code(() => validateGraphLink(link, MESSAGES, "v1.0", INBOX))).toBe("invalid_input");
  });

  it("rejects non-strings and wrong-endpoint delta links", () => {
    expect(code(() => validateGraphLink(undefined, MESSAGES, "v1.0"))).toBe("invalid_input");
    expect(code(() => validateGraphLink(42, MESSAGES, "v1.0"))).toBe("invalid_input");
    expect(
      code(() => validateGraphLink(`${G}/v1.0/me/calendarView/delta?x=1`, EVENTS_DELTA, "v1.0")),
    ).toBe("invalid_input");
    expect(code(() => validateGraphLink(`${G}/v1.0/me/events/delta`, MESSAGES, "v1.0"))).toBe(
      "invalid_input",
    );
  });
});

describe("validateGraphLink principal binding", () => {
  const G = "https://graph.microsoft.com";
  const USER_MSGS = ep("/users/{user-id}/messages");
  const USER_FOLDER_DELTA = ep("/users/{user-id}/mailFolders/{mailFolder-id}/messages/delta()");
  const ALICE = { "user-id": "alice@contoso.com" };

  it("accepts legitimate continuations for the same principal", () => {
    expect(
      validateGraphLink(
        `${G}/v1.0/users/alice@contoso.com/messages?$skip=10`,
        USER_MSGS,
        "v1.0",
        ALICE,
      ),
    ).toContain("/users/alice@contoso.com/messages");
    // case, percent-encoding and OData key syntax are equivalent spellings of the same id
    expect(
      validateGraphLink(
        `${G}/v1.0/users/ALICE%40contoso.com/messages?$skip=10`,
        USER_MSGS,
        "v1.0",
        ALICE,
      ),
    ).toBeTruthy();
    expect(
      validateGraphLink(
        `${G}/v1.0/users('alice@contoso.com')/messages?$skip=10`,
        USER_MSGS,
        "v1.0",
        ALICE,
      ),
    ).toBeTruthy();
  });

  it("rejects a nextLink that switches the user", () => {
    for (const link of [
      `${G}/v1.0/users/bob@contoso.com/messages?$skip=10`,
      `${G}/v1.0/users('bob@contoso.com')/messages?$skip=10`,
      `${G}/v1.0/users/alice@contoso.com.evil/messages`,
      `${G}/v1.0/users/%25/messages`,
    ]) {
      expect(code(() => validateGraphLink(link, USER_MSGS, "v1.0", ALICE))).toBe("invalid_input");
    }
    // missing original params fail closed
    expect(
      code(() =>
        validateGraphLink(`${G}/v1.0/users/alice@contoso.com/messages`, USER_MSGS, "v1.0"),
      ),
    ).toBe("invalid_input");
  });

  it("rejects /me rewritten to /users/{id} and users swapped in delta links", () => {
    expect(
      code(() => validateGraphLink(`${G}/v1.0/users/bob/messages`, ep("/me/messages"), "v1.0")),
    ).toBe("invalid_input");
    expect(
      code(() =>
        validateGraphLink(
          `${G}/v1.0/users/bob@contoso.com/mailFolders('inbox')/messages/delta?$deltatoken=t`,
          USER_FOLDER_DELTA,
          "v1.0",
          { ...ALICE, ...INBOX },
        ),
      ),
    ).toBe("invalid_input");
  });

  it("accepts a delta link for the same user and tolerates a canonicalized folder id", () => {
    const params = { ...ALICE, ...INBOX };
    expect(
      validateGraphLink(
        `${G}/v1.0/users/alice@contoso.com/mailFolders('AAMkAD0=')/messages/delta?$deltatoken=t`,
        USER_FOLDER_DELTA,
        "v1.0",
        params,
      ),
    ).toContain("$deltatoken=t");
  });

  it("binds every non-canonicalizable parameter (group, site, calendar)", () => {
    const grp = ep("/groups/{group-id}/events");
    expect(
      validateGraphLink(`${G}/v1.0/groups/${GUID_A}/events?$skip=1`, grp, "v1.0", {
        "group-id": GUID_A,
      }),
    ).toBeTruthy();
    expect(
      code(() =>
        validateGraphLink(`${G}/v1.0/groups/${GUID_B}/events?$skip=1`, grp, "v1.0", {
          "group-id": GUID_A,
        }),
      ),
    ).toBe("invalid_input");
  });

  it("compares opaque ids exactly; only user/group/team ids ignore case", () => {
    const cal = ep("/me/calendars/{calendar-id}/events");
    const params = { "calendar-id": "AAMkAbC=" };
    expect(
      validateGraphLink(`${G}/v1.0/me/calendars/AAMkAbC=/events?$skip=1`, cal, "v1.0", params),
    ).toBeTruthy();
    expect(
      code(() =>
        validateGraphLink(`${G}/v1.0/me/calendars/AAMKABC=/events?$skip=1`, cal, "v1.0", params),
      ),
    ).toBe("invalid_input");
    const grp = ep("/groups/{group-id}/events");
    expect(
      validateGraphLink(`${G}/v1.0/groups/${GUID_A.toUpperCase()}/events?$skip=1`, grp, "v1.0", {
        "group-id": GUID_A,
      }),
    ).toBeTruthy();
  });
});
