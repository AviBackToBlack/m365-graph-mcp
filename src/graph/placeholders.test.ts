import { beforeAll, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointDefinition } from "../catalog.js";
import { validateBatch } from "./batch.js";
import { GraphClientError } from "./errors.js";
import { buildPath, validateGraphLink, validateRelativeGraphUrl } from "./url.js";

const GUID_A = "f17e0000-0000-4000-8000-000000000004";
const GUID_B = "f17e0000-0000-4000-8000-000000000005";
const EXCHANGE_ID =
  "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAFNAAA=";
const SHORT_ID = "AAMkAGI2TG93AAA=";
const G = "https://graph.microsoft.com";

const ep = (pathPattern: string, extra: Record<string, unknown> = {}): EndpointDefinition => ({
  pathPattern,
  method: "get",
  toolName: "t",
  ...extra,
});

const VALID_IDS: Record<string, string[]> = {
  "group-id": [GUID_A, GUID_A.toUpperCase()],
  "team-id": [GUID_A],
  "list-id": [GUID_A],
  "columnDefinition-id": [GUID_A],
  "user-id": [
    GUID_A,
    "alice@contoso.com",
    "o'brien+tag@sub.contoso.co.uk",
    "bob_contoso.com#EXT#@t.onmicrosoft.com",
  ],
  "listItem-id": ["1", "4021"],
  "chatMessage-id": ["1662678714353"],
  "chat-id": [
    "19:b8577894a63548969c5c92bb9c80c5e1@thread.v2",
    "19:meeting_ZDZlYTYxOWUtYzdlMi00ZmMxLWIxMTAtN2YzODZlZjAxYzI4@thread.v2",
    `19:${GUID_A}_${GUID_B}@unq.gbl.spaces`,
    "19:abc@thread.tacv2",
    "48:notes",
  ],
  "channel-id": ["19:abc123def@thread.tacv2", "19:abc123def@thread.skype", "19:abc@thread.v2"],
  "message-id": [SHORT_ID, EXCHANGE_ID, "AQMkADAwATM0MDAAMS0xYzBiLTA="],
  "event-id": [SHORT_ID, EXCHANGE_ID],
  "attachment-id": [SHORT_ID, EXCHANGE_ID],
  "calendar-id": [SHORT_ID, EXCHANGE_ID],
  "mailFolder-id": [
    ...[
      "inbox",
      "Inbox",
      "drafts",
      "sentitems",
      "deleteditems",
      "archive",
      "junkemail",
      "outbox",
      "clutter",
      "conflicts",
      "conversationhistory",
      "localfailures",
      "msgfolderroot",
      "recoverableitemsdeletions",
      "scheduled",
      "searchfolders",
      "serverfailures",
      "syncissues",
    ],
    EXCHANGE_ID,
  ],
  "drive-id": ["b!Xk9-_abcDEF123", "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K"],
  "site-id": [
    "root",
    "contoso.sharepoint.com",
    `contoso.sharepoint.com,${GUID_A},${GUID_B}`,
    `${GUID_A},${GUID_B}`,
    GUID_A,
  ],
};

/** Values that must be rejected for EVERY placeholder (OData control / route-like / malformed). */
const HOSTILE_VALUES = [
  "delta",
  "Delta",
  "DELTA",
  "delta()",
  "$count",
  "$COUNT",
  "%24count",
  "%2524count",
  "$ref",
  "$value",
  "$metadata",
  "$batch",
  "$entity",
  "microsoft.graph.getByIds()",
  "microsoft.graph.getByIds",
  "Microsoft.Graph.foo",
  "foo()",
  "getByIds",
  "count",
  "ref",
  "value",
  "metadata",
  "..",
  ".",
  "",
  "a/b",
  "a\\b",
  "x%2Fy",
  "%2e%2e",
  "a b",
  "inbox/delta",
  "inbox()",
  "('inbox')",
];

function code(fn: () => unknown): string {
  try {
    fn();
  } catch (e) {
    if (e instanceof GraphClientError) return e.code;
    throw e;
  }
  return "no-error";
}

describe("placeholder shapes via buildPath (direct-tool path)", () => {
  const names = Object.keys(VALID_IDS);

  it("has a shape for every placeholder used by the catalog", async () => {
    const used = new Set(
      (await loadEndpointCatalog()).flatMap((e) =>
        [...e.pathPattern.matchAll(/\{([^{}/]+)\}/g)].map((m) => m[1] as string),
      ),
    );
    used.delete("path"); // skipEncoding value, validated separately
    // attachment-id has a shape for future entries but no catalog pattern uses it yet.
    expect([...used].sort()).toEqual(names.filter((n) => n !== "attachment-id").sort());
  });

  for (const name of names) {
    const entry = ep(`/x/{${name}}/y`);
    it(`accepts realistic ${name} values`, () => {
      for (const v of VALID_IDS[name] as string[]) {
        expect(() => buildPath(entry, { [name]: v }), v).not.toThrow();
      }
    });
    it(`rejects hostile and route-like ${name} values`, () => {
      for (const v of HOSTILE_VALUES) {
        expect(
          code(() => buildPath(entry, { [name]: v })),
          `${name}=${v}`,
        ).toBe("invalid_input");
      }
    });
  }

  it("rejects a placeholder that has no shape", () => {
    expect(code(() => buildPath(ep("/x/{unknown-id}"), { "unknown-id": "abc123" }))).toBe(
      "invalid_input",
    );
  });

  it("rejects the documented escapes on catalog-style routes", () => {
    expect(code(() => buildPath(ep("/groups/{group-id}"), { "group-id": "delta" }))).toBe(
      "invalid_input",
    );
    expect(code(() => buildPath(ep("/users/{user-id}/messages"), { "user-id": "delta()" }))).toBe(
      "invalid_input",
    );
    expect(
      code(() =>
        buildPath(ep("/me/mailFolders/{mailFolder-id}/messages"), { "mailFolder-id": "delta" }),
      ),
    ).toBe("invalid_input");
  });

  it("requires exactly one @ and a dotted domain for UPNs", () => {
    for (const v of ["a@b@contoso.com", "a@contoso", "@contoso.com", "alice@contoso.com/x"]) {
      expect(
        code(() => buildPath(ep("/users/{user-id}"), { "user-id": v })),
        v,
      ).toBe("invalid_input");
    }
  });

  it("encodes accepted values as one segment (UPN @, Teams colon, Exchange =+)", () => {
    expect(buildPath(ep("/users/{user-id}/m"), { "user-id": "alice@contoso.com" })).toBe(
      "/users/alice%40contoso.com/m",
    );
    expect(buildPath(ep("/chats/{chat-id}"), { "chat-id": "19:abc@thread.v2" })).toBe(
      "/chats/19%3Aabc%40thread.v2",
    );
    expect(buildPath(ep("/me/messages/{message-id}"), { "message-id": "AAMk+a=b-c_d" })).toBe(
      "/me/messages/AAMk%2Ba=b-c_d",
    );
  });
});

describe("placeholder shapes via the batch validator (real catalog)", () => {
  let permitted: EndpointDefinition[];
  beforeAll(async () => {
    permitted = (await loadEndpointCatalog()).filter((e) => e.method === "get");
  });

  const accepts = (url: string): EndpointDefinition =>
    validateRelativeGraphUrl(url, permitted, "v1.0").entry;
  const rejects = (url: string): void => {
    expect(
      code(() => validateRelativeGraphUrl(url, permitted, "v1.0")),
      url,
    ).toBe("invalid_input");
  };

  it("accepts realistic ids for every placeholder", () => {
    expect(accepts(`/groups/${GUID_A}`).toolName).toBe("get-group");
    expect(accepts(`/teams/${GUID_A}`).toolName).toBe("get-team");
    expect(accepts("/users/alice@contoso.com/messages").toolName).toBe(
      "list-shared-mailbox-messages",
    );
    expect(accepts("/users/alice%40contoso.com/messages").toolName).toBe(
      "list-shared-mailbox-messages",
    );
    expect(accepts("/users('alice@contoso.com')/messages").toolName).toBe(
      "list-shared-mailbox-messages",
    );
    expect(accepts(`/users/${GUID_A}/presence`).toolName).toBe("get-user-presence");
    expect(accepts("/me/mailFolders/inbox/messages").toolName).toBe("list-mail-folder-messages");
    expect(accepts("/me/mailFolders('inbox')/messages/delta()").toolName).toBe(
      "list-mail-folder-messages-delta",
    );
    expect(accepts(`/me/mailFolders/${EXCHANGE_ID}/childFolders`).toolName).toBe(
      "list-mail-child-folders",
    );
    expect(accepts(`/me/messages/${SHORT_ID}`).toolName).toBe("get-mail-message");
    expect(accepts(`/me/messages/${EXCHANGE_ID}/attachments`).toolName).toBe(
      "list-mail-attachments",
    );
    expect(accepts(`/me/events/${SHORT_ID}`).toolName).toBe("get-calendar-event");
    expect(accepts(`/me/calendars/${SHORT_ID}/events/${EXCHANGE_ID}/instances`).toolName).toBe(
      "list-calendar-event-instances",
    );
    expect(accepts("/chats/19:b8577894a63548969c5c92bb9c80c5e1@thread.v2").toolName).toBe(
      "get-chat",
    );
    expect(
      accepts("/chats/19%3Ab8577894a63548969c5c92bb9c80c5e1%40thread.v2/messages").toolName,
    ).toBe("list-chat-messages");
    expect(
      accepts(`/teams/${GUID_A}/channels/19:abc123@thread.tacv2/messages/1662678714353`).toolName,
    ).toBe("get-channel-message");
    expect(accepts("/sites/root").toolName).toBe("get-sharepoint-site");
    expect(accepts("/sites/contoso.sharepoint.com").toolName).toBe("get-sharepoint-site");
    expect(
      accepts(`/sites/contoso.sharepoint.com,${GUID_A},${GUID_B}/lists/${GUID_A}/items/12`)
        .toolName,
    ).toBe("get-sharepoint-site-list-item");
    expect(accepts(`/sites/root/lists/${GUID_A}/columns/${GUID_B}`).toolName).toBe(
      "get-sharepoint-list-column",
    );
    expect(accepts("/sites/root/drives/b!Xk9-_abcDEF123").toolName).toBe(
      "get-sharepoint-site-drive-by-id",
    );
    // D32: the site `items` collection (non-enumerable, abstract baseItem) is not in the catalog.
    rejects("/sites/root/items");
    rejects("/sites/root/items/01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K");
    rejects(
      `/sites/contoso.sharepoint.com,${GUID_A},${GUID_B}/items/01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K`,
    );
    // D36: sites delta is application-only (delegated not supported) and not in the catalog.
    rejects("/sites/delta()");
    rejects("/sites/delta");
    rejects("/sites/delta?token=latest");
    // Beta-only entry (D2): reachable only under the beta version, never v1.0.
    expect(validateRelativeGraphUrl("/me/events/delta()", permitted, "beta").entry.toolName).toBe(
      "list-calendar-events-delta",
    );
    rejects("/me/events/delta()");
  });

  it.each([
    "/groups/delta",
    "/groups/%64elta",
    "/groups/%2564elta",
    "/groups/Delta",
    "/groups/delta()",
    "/groups/$count",
    "/groups/%24count",
    "/groups/%2524count",
    "/groups/$ref",
    "/groups/$value",
    "/groups/$metadata",
    "/groups/microsoft.graph.getByIds()",
    "/groups/foo()",
    "/groups/getByIds",
    "/groups/getByIds/events",
    "/groups/delta/members",
    "/groups/microsoft.graph.delta/members",
    "/users/delta()/messages",
    "/users/delta/messages",
    "/users/$count/messages",
    "/users/microsoft.graph.getByIds()/presence",
    "/users('delta()')/messages",
    "/me/mailFolders/delta/messages",
    "/me/mailFolders/delta/childFolders",
    "/me/mailFolders/Delta/messages/delta()",
    "/me/mailFolders/$count/messages",
    "/me/mailFolders/microsoft.graph.x()/messages",
    "/me/messages/$value",
    "/me/messages/delta",
    "/me/messages/getByIds",
    "/me/messages/%24value",
    "/me/messages/%2524value/attachments",
    "/me/events/foo()",
    "/me/calendars/delta/events",
    "/me/calendars/$count/events",
    "/chats/delta",
    "/chats/getAllMessages",
    "/chats/19:abc@thread.v2/messages/delta",
    "/chats/19:abc@thread.v2/messages/$count",
    "/teams/delta",
    "/teams/getAllMessages",
    `/teams/${GUID_A}/channels/getAllMessages`,
    `/teams/${GUID_A}/channels/delta`,
    `/teams/${GUID_A}/channels/19:abc@thread.tacv2/messages/delta`,
    "/sites/getAllSites",
    "/sites/$count",
    "/sites/root/drives/delta",
    "/sites/root/items/delta",
    `/sites/root/lists/${GUID_A}/items/delta`,
    `/sites/root/lists/${GUID_A}/columns/$count`,
  ])("rejects route-escape value %s", (url) => {
    rejects(url);
  });

  it("does not let a failed capture be steered onto a different entry", () => {
    // `delta` matches /groups/{group-id}; with no other candidate the path must be rejected.
    expect(code(() => validateRelativeGraphUrl("/groups/delta", permitted, "v1.0"))).toBe(
      "invalid_input",
    );
    // The legitimate literal delta entry still wins where the catalog has one.
    expect(
      validateRelativeGraphUrl("/me/events/delta()", permitted, "beta").entry.pathPattern,
    ).toBe("/me/events/delta()");
    // A valid-shape capture still resolves to the placeholder entry.
    expect(accepts(`/me/events/${SHORT_ID}`).pathPattern).toBe("/me/events/{event-id}");
  });

  it("applies the same checks to batch subrequests", () => {
    for (const url of [
      "/groups/delta",
      "/users/delta()/messages",
      "/me/mailFolders/delta/messages",
    ]) {
      expect(
        code(() =>
          validateBatch(
            { requests: [{ id: "1", method: "GET", url }] },
            new Map(permitted.map((e) => [e, { check: () => undefined }])),
            "v1.0",
          ),
        ),
        url,
      ).toBe("invalid_input");
    }
  });

  describe("{path} colon in any encoding", () => {
    const site = "/sites/contoso.sharepoint.com:";
    it.each(["%3A", "%3a", "%253A", "%253a", "%25253A", "%2525253A"])(
      "rejects %s inside {path}",
      (enc) => {
        rejects(`${site}/sites/a${enc}/drive`);
        rejects(`${site}/sites/a${enc}`);
      },
    );
    it("rejects malformed escapes inside {path} and OData control segments", () => {
      rejects(`${site}/sites/a%zz`);
      rejects(`${site}/sites/$value`);
      rejects(`${site}/sites/microsoft.graph.x`);
    });
    it("still accepts ordinary site paths", () => {
      expect(accepts(`${site}/sites/hr%20team`).toolName).toBe("get-sharepoint-site-by-path");
      expect(accepts(`${site}/teams/1drvteam`).toolName).toBe("get-sharepoint-site-by-path");
    });
  });
});

describe("validateGraphLink shape checks", () => {
  const MESSAGES = ep("/me/mailFolders/{mailFolder-id}/messages");
  const INBOX = { "mailFolder-id": "inbox" };

  it("lets the binding-exempt mailFolder-id be canonicalized to a real opaque id", () => {
    expect(
      validateGraphLink(
        `${G}/v1.0/me/mailFolders('${EXCHANGE_ID}')/messages?$skiptoken=a`,
        MESSAGES,
        "v1.0",
        INBOX,
      ),
    ).toContain("$skiptoken=a");
    expect(
      validateGraphLink(`${G}/v1.0/me/mailFolders/${SHORT_ID}/messages`, MESSAGES, "v1.0", INBOX),
    ).toBeTruthy();
  });

  it("rejects route-like folder ids even though the folder id is not bound", () => {
    for (const v of ["delta", "$count", "foo()", "microsoft.graph.x", "%24count", "a%20b"]) {
      expect(
        code(() =>
          validateGraphLink(`${G}/v1.0/me/mailFolders/${v}/messages`, MESSAGES, "v1.0", INBOX),
        ),
        v,
      ).toBe("invalid_input");
    }
  });

  it("rejects route-like bound values too (user switch to delta())", () => {
    expect(
      code(() =>
        validateGraphLink(
          `${G}/v1.0/users/delta()/messages`,
          ep("/users/{user-id}/messages"),
          "v1.0",
          { "user-id": "delta()" },
        ),
      ),
    ).toBe("invalid_input");
  });

  it.each(["a%3Ab", "a%3ab", "a%253Ab"])("rejects colon encoding %s in {path} links", (v) => {
    const site = ep("/sites/{site-id}:/{path}", { skipEncoding: ["path"] });
    expect(
      code(() =>
        validateGraphLink(`${G}/v1.0/sites/h.sharepoint.com:/${v}`, site, "v1.0", {
          "site-id": "h.sharepoint.com",
          path: v,
        }),
      ),
    ).toBe("invalid_input");
  });
});

describe("get-user-presence user-id is GUID-only (D39), the shared user-id shape is unchanged", () => {
  const PRESENCE = "/users/{user-id}/presence";
  const NOT_GUIDS = [
    "alice@contoso.com",
    "alice%40contoso.com",
    "o'brien+tag@sub.contoso.co.uk",
    "bob_contoso.com#EXT#@t.onmicrosoft.com",
    GUID_A + "x",
    `{${GUID_A}}`,
    GUID_A.replaceAll("-", ""),
    "me",
  ];
  let permitted: EndpointDefinition[];
  beforeAll(async () => {
    permitted = (await loadEndpointCatalog()).filter((e) => e.method === "get");
  });
  const batchOf = (url: string) =>
    validateBatch(
      { requests: [{ id: "1", method: "GET", url }] },
      new Map(permitted.map((e) => [e, { check: () => undefined }])),
      "v1.0",
    );

  it("buildPath accepts a GUID (either case) and rejects every UPN/email form", () => {
    for (const v of [GUID_A, GUID_A.toUpperCase()]) {
      expect(buildPath(ep(PRESENCE), { "user-id": v })).toBe(`/users/${v}/presence`);
    }
    for (const v of NOT_GUIDS) {
      expect(
        code(() => buildPath(ep(PRESENCE), { "user-id": v })),
        v,
      ).toBe("invalid_input");
    }
  });

  it("the error never echoes the rejected email", () => {
    try {
      buildPath(ep(PRESENCE), { "user-id": "alice@contoso.com" });
      expect.unreachable();
    } catch (e) {
      expect((e as GraphClientError).message).not.toContain("alice");
    }
  });

  it("graph-batch refuses the same UPN/email forms (parity with the single path), accepts GUIDs", () => {
    expect(batchOf(`/users/${GUID_A}/presence`).requests).toHaveLength(1);
    expect(batchOf(`/users/${GUID_A.toUpperCase()}/presence`).requests).toHaveLength(1);
    for (const url of [
      "/users/alice@contoso.com/presence",
      "/users/alice%40contoso.com/presence",
      "/users('alice@contoso.com')/presence",
      "/users/bob_contoso.com%23EXT%23@t.onmicrosoft.com/presence",
      `/users/${GUID_A}x/presence`,
    ]) {
      expect(
        code(() => batchOf(url)),
        url,
      ).toBe("invalid_input");
    }
    expect(batchOf(`/users('${GUID_A}')/presence`).requests).toHaveLength(1);
  });

  it("other /users/{user-id}/... routes still accept a UPN (shared validator untouched)", () => {
    expect(buildPath(ep("/users/{user-id}/messages"), { "user-id": "alice@contoso.com" })).toBe(
      "/users/alice%40contoso.com/messages",
    );
    expect(buildPath(ep("/users/{user-id}"), { "user-id": "alice@contoso.com" })).toBe(
      "/users/alice%40contoso.com",
    );
    expect(
      validateRelativeGraphUrl("/users/alice@contoso.com/messages", permitted, "v1.0").entry
        .toolName,
    ).toBe("list-shared-mailbox-messages");
  });
});
