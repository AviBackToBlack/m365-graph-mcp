import { describe, expect, it } from "vitest";
import {
  ACCESS_TOKEN_ENV_VAR,
  DEFAULT_CLOCK_SKEW_MS,
  DEFAULT_EXPIRY_BUFFER_MS,
  MAX_TOKEN_LENGTH,
  ManualTokenProvider,
  TENANT_ID_ENV_VAR,
} from "./manual-token-provider.js";
import { TokenError } from "./token-provider.js";

const NOW_MS = Date.UTC(2030, 0, 1, 12, 0, 0);
const NOW_S = NOW_MS / 1000;
const TENANT = "11111111-2222-3333-4444-555555555555";
const GRAPH_AUD = "https://graph.microsoft.com";

function b64url(value: unknown): string {
  return Buffer.from(typeof value === "string" ? value : JSON.stringify(value)).toString(
    "base64url",
  );
}

/** Builds a fake, unsigned-in-spirit JWT; the signature segment is a dummy. */
function fakeJwt(claims: Record<string, unknown>): string {
  return `${b64url({ alg: "RS256", typ: "JWT" })}.${b64url(claims)}.FAKE_SIGNATURE_SEGMENT`;
}

function validClaims(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    aud: GRAPH_AUD,
    exp: NOW_S + 3600,
    tid: TENANT,
    marker: "SECRET_PAYLOAD_MARKER",
    ...overrides,
  };
}

function provider(
  env: Record<string, string | undefined>,
  opts: {
    now?: () => number;
    tenantId?: string;
    expiryBufferMs?: number;
    clockSkewMs?: number;
  } = {},
): ManualTokenProvider {
  return new ManualTokenProvider({ env, now: opts.now ?? (() => NOW_MS), ...opts });
}

async function failure(p: ManualTokenProvider): Promise<TokenError> {
  try {
    await p.getAccessToken();
  } catch (error) {
    expect(error).toBeInstanceOf(TokenError);
    return error as TokenError;
  }
  throw new Error("expected getAccessToken to reject");
}

describe("ManualTokenProvider", () => {
  it("returns a valid token", async () => {
    const token = fakeJwt(validClaims());
    await expect(provider({ [ACCESS_TOKEN_ENV_VAR]: token }).getAccessToken()).resolves.toBe(token);
  });

  it("trims whitespace and strips a leading Bearer prefix", async () => {
    const token = fakeJwt(validClaims());
    await expect(
      provider({ [ACCESS_TOKEN_ENV_VAR]: `  Bearer ${token}\n` }).getAccessToken(),
    ).resolves.toBe(token);
  });

  it.each(["https://graph.microsoft.com/", "00000003-0000-0000-c000-000000000000"])(
    "accepts Graph audience %s",
    async (aud) => {
      const token = fakeJwt(validClaims({ aud }));
      await expect(provider({ [ACCESS_TOKEN_ENV_VAR]: token }).getAccessToken()).resolves.toBe(
        token,
      );
    },
  );

  it.each([undefined, "", "   ", "Bearer "])(
    "rejects missing token env (%j) with reauth_required",
    async (value) => {
      const err = await failure(provider({ [ACCESS_TOKEN_ENV_VAR]: value }));
      expect(err.code).toBe("reauth_required");
      expect(err.message).toContain(ACCESS_TOKEN_ENV_VAR);
    },
  );

  const OPAQUE_TOKENS: Array<[string, string]> = [
    ["non-JWT string", "opaque-token_value~123"],
    ["two segments", "aaa.bbb"],
    ["four segments", "aaa.bbb.ccc.ddd"],
    ["5-segment JWE shape", "aGRy.ZW5jS2V5.aXY.Y2lwaGVy.dGFn"],
    ["empty signature", `${b64url({})}.${b64url(validClaims())}.`],
    ["non-base64url chars", "aa$a.bb$b.cc$c"],
    ["3-segment non-JSON payload", `${b64url({})}.${b64url("not json{")}.sig`],
    ["payload JSON array", `${b64url({})}.${b64url([1])}.sig`],
    ["payload JSON null", `${b64url({})}.${b64url("null")}.sig`],
  ];

  it.each(OPAQUE_TOKENS)("passes through opaque token without a pin: %s", async (_n, token) => {
    await expect(provider({ [ACCESS_TOKEN_ENV_VAR]: token }).getAccessToken()).resolves.toBe(token);
  });

  it("trims and strips Bearer from opaque tokens", async () => {
    await expect(
      provider({ [ACCESS_TOKEN_ENV_VAR]: " Bearer opaque-token " }).getAccessToken(),
    ).resolves.toBe("opaque-token");
  });

  it.each(OPAQUE_TOKENS)("rejects opaque token with a tenant pin: %s", async (_n, token) => {
    const err = await failure(
      provider({ [ACCESS_TOKEN_ENV_VAR]: token, [TENANT_ID_ENV_VAR]: TENANT }),
    );
    expect(err.code).toBe("reauth_required");
    expect(err.message).toContain("opaque");
    expect(err.message).toContain("tenant pin");
    expect(err.message).not.toContain(token);
  });

  it.each([
    ["inner space", "abc def"],
    ["inner tab", "abc\tdef"],
    ["inner newline", "abc\ndef"],
    ["NUL", "abc\u0000def"],
    ["DEL", "abc\u007fdef"],
    ["non-ASCII", "abc\u00e9def"],
    ["whitespace in JWT form", `${fakeJwt(validClaims())} x`],
    ["newline in JWT form", fakeJwt(validClaims()).replace(".", "\n.")],
  ])("rejects header-unsafe token: %s", async (_n, token) => {
    const err = await failure(provider({ [ACCESS_TOKEN_ENV_VAR]: token }));
    expect(err.code).toBe("reauth_required");
    expect(err.message).toContain("whitespace, control");
    expect(err.message).not.toContain(token);
  });

  it("rejects over-length tokens but accepts the maximum", async () => {
    const max = "a".repeat(MAX_TOKEN_LENGTH);
    await expect(provider({ [ACCESS_TOKEN_ENV_VAR]: max }).getAccessToken()).resolves.toBe(max);
    const over = "a".repeat(MAX_TOKEN_LENGTH + 1);
    const err = await failure(provider({ [ACCESS_TOKEN_ENV_VAR]: over }));
    expect(err.code).toBe("reauth_required");
    expect(err.message).toContain("maximum");
    expect(err.message).not.toContain(over.slice(0, 50));
  });

  it.each([
    ["different audience", { aud: "https://management.azure.com" }],
    ["missing audience", { aud: undefined }],
    ["array audience", { aud: [GRAPH_AUD] }],
    ["lookalike audience", { aud: "https://graph.microsoft.com.evil.example" }],
  ])("rejects %s", async (_name, overrides) => {
    const err = await failure(
      provider({ [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims(overrides)) }),
    );
    expect(err.code).toBe("reauth_required");
    expect(err.message).toContain("audience");
  });

  it("rejects an expired token and names the expiry time", async () => {
    const exp = NOW_S - 10;
    const err = await failure(provider({ [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ exp })) }));
    expect(err.code).toBe("reauth_required");
    expect(err.message).toContain(new Date(exp * 1000).toISOString());
  });

  it.each([undefined, "soon", null])("rejects missing or invalid exp (%j)", async (exp) => {
    const err = await failure(provider({ [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ exp })) }));
    expect(err.code).toBe("reauth_required");
  });

  it("rejects an absurdly negative exp without throwing a RangeError", async () => {
    const err = await failure(
      provider({ [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ exp: -1e300 })) }),
    );
    expect(err).toBeInstanceOf(TokenError);
  });

  it.each([1e308, -1e308, 8.64e12 + 1, -(8.64e12 + 1)])(
    "rejects out-of-range exp %j without echoing it",
    async (exp) => {
      const err = await failure(
        provider({ [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ exp })) }),
      );
      expect(err.code).toBe("reauth_required");
      expect(err.message).toContain("out-of-range");
      expect(err.message).not.toContain(String(exp));
    },
  );

  it.each([1e308, -1e308, 8.64e12 + 1, -(8.64e12 + 1)])(
    "rejects out-of-range nbf %j without echoing it",
    async (nbf) => {
      const err = await failure(
        provider({ [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ nbf })) }),
      );
      expect(err.code).toBe("reauth_required");
      expect(err.message).toContain("out-of-range");
      expect(err.message).not.toContain(String(nbf));
    },
  );

  it.each([
    ["expiryBufferMs", -1],
    ["expiryBufferMs", Number.NaN],
    ["expiryBufferMs", Number.POSITIVE_INFINITY],
    ["expiryBufferMs", Number.NEGATIVE_INFINITY],
    ["expiryBufferMs", 3_600_001],
    ["clockSkewMs", -1],
    ["clockSkewMs", Number.NaN],
    ["clockSkewMs", Number.POSITIVE_INFINITY],
    ["clockSkewMs", Number.NEGATIVE_INFINITY],
    ["clockSkewMs", 3_600_001],
  ])("rejects invalid %s option %j at construction", (name, value) => {
    let caught: unknown;
    try {
      provider({}, { [name]: value });
    } catch (error) {
      caught = error;
    }
    expect(caught).toBeInstanceOf(TokenError);
    expect((caught as TokenError).code).toBe("auth_misconfigured");
    expect((caught as TokenError).message).not.toContain(String(value));
  });

  it("accepts zero and the maximum for the time options", () => {
    expect(() => provider({}, { expiryBufferMs: 0, clockSkewMs: 3_600_000 })).not.toThrow();
  });

  it("applies the expiry buffer at the exact boundary", async () => {
    const exp = NOW_S + DEFAULT_EXPIRY_BUFFER_MS / 1000;
    const env = { [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ exp })) };
    // now == exp - buffer: rejected; one ms earlier: accepted.
    expect((await failure(provider(env))).code).toBe("reauth_required");
    await expect(provider(env, { now: () => NOW_MS - 1 }).getAccessToken()).resolves.toBeTypeOf(
      "string",
    );
  });

  it("applies clock skew to nbf at the exact boundary", async () => {
    const nbf = NOW_S + DEFAULT_CLOCK_SKEW_MS / 1000;
    const env = { [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ nbf })) };
    await expect(provider(env).getAccessToken()).resolves.toBeTypeOf("string");
    const err = await failure(provider(env, { now: () => NOW_MS - 1 }));
    expect(err.code).toBe("reauth_required");
    expect(err.message).toContain("nbf");
  });

  it.each([NaN, Infinity, -Infinity, 8.64e15 + 1])(
    "fails closed when the clock returns an invalid time (%s)",
    async (bad) => {
      const env = { [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims()) };
      const err = await failure(provider(env, { now: () => bad }));
      expect(err.code).toBe("auth_misconfigured");
    },
  );

  it("rejects a non-numeric nbf", async () => {
    const err = await failure(
      provider({ [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ nbf: "now" })) }),
    );
    expect(err.code).toBe("reauth_required");
  });

  it("re-checks expiry on every call as the clock advances", async () => {
    let now = NOW_MS;
    const token = fakeJwt(validClaims({ exp: NOW_S + 120 }));
    const p = provider({ [ACCESS_TOKEN_ENV_VAR]: token }, { now: () => now });
    await expect(p.getAccessToken()).resolves.toBe(token);
    now += 60_000;
    await expect(p.getAccessToken()).resolves.toBe(token);
    now += 60_000;
    expect((await failure(p)).code).toBe("reauth_required");
    now -= 120_000; // clock moving backwards is simply re-evaluated
    await expect(p.getAccessToken()).resolves.toBe(token);
  });

  it("picks up a refreshed token from the environment on later calls", async () => {
    const env: Record<string, string | undefined> = {
      [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ exp: NOW_S - 100 })),
    };
    const p = provider(env);
    expect((await failure(p)).code).toBe("reauth_required");
    const fresh = fakeJwt(validClaims());
    env[ACCESS_TOKEN_ENV_VAR] = fresh;
    await expect(p.getAccessToken()).resolves.toBe(fresh);
  });

  describe("tenant pin", () => {
    it("accepts a matching tid (case-insensitive) from options", async () => {
      const token = fakeJwt(validClaims({ tid: TENANT.toUpperCase() }));
      await expect(
        provider({ [ACCESS_TOKEN_ENV_VAR]: token }, { tenantId: TENANT }).getAccessToken(),
      ).resolves.toBe(token);
    });

    it("accepts a matching tid from the environment", async () => {
      const token = fakeJwt(validClaims());
      const env = { [ACCESS_TOKEN_ENV_VAR]: token, [TENANT_ID_ENV_VAR]: TENANT };
      await expect(provider(env).getAccessToken()).resolves.toBe(token);
    });

    it("rejects a mismatched or missing tid", async () => {
      for (const tid of ["99999999-2222-3333-4444-555555555555", undefined, 5]) {
        const env = {
          [ACCESS_TOKEN_ENV_VAR]: fakeJwt(validClaims({ tid })),
          [TENANT_ID_ENV_VAR]: TENANT,
        };
        const err = await failure(provider(env));
        expect(err.code).toBe("reauth_required");
        expect(err.message).toContain("tenant");
      }
    });

    it("does not enforce a tenant when no pin is configured", async () => {
      const token = fakeJwt(validClaims({ tid: undefined }));
      await expect(provider({ [ACCESS_TOKEN_ENV_VAR]: token }).getAccessToken()).resolves.toBe(
        token,
      );
    });

    it.each([
      "not-a-guid",
      "{11111111-2222-3333-4444-555555555555}",
      "common",
      "11111111222233334444555555555555",
    ])("rejects invalid pin %j as auth_misconfigured at construction", (pin) => {
      for (const make of [
        () => new ManualTokenProvider({ env: {}, tenantId: pin }),
        () => new ManualTokenProvider({ env: { [TENANT_ID_ENV_VAR]: pin } }),
      ]) {
        try {
          make();
        } catch (error) {
          expect(error).toBeInstanceOf(TokenError);
          expect((error as TokenError).code).toBe("auth_misconfigured");
          continue;
        }
        throw new Error("expected constructor to throw");
      }
    });

    it("treats a blank tenant env var as unset", async () => {
      const token = fakeJwt(validClaims());
      const env = { [ACCESS_TOKEN_ENV_VAR]: token, [TENANT_ID_ENV_VAR]: "  " };
      await expect(provider(env).getAccessToken()).resolves.toBe(token);
    });
  });

  describe("token never leaks through error surfaces", () => {
    const token = fakeJwt(validClaims({ exp: NOW_S - 100 }));
    const [header, payload, signature] = token.split(".") as [string, string, string];
    const fragments = [
      token,
      token.slice(0, 20),
      header,
      payload,
      signature,
      "SECRET_PAYLOAD_MARKER",
    ];

    function assertNoLeak(err: TokenError): void {
      const surfaces = [
        err.message,
        String(err),
        err.toString(),
        JSON.stringify(err),
        JSON.stringify(err, Object.getOwnPropertyNames(err)),
        err.stack ?? "",
        String(err.cause ?? ""),
      ];
      expect(err.cause).toBeUndefined();
      for (const surface of surfaces) {
        for (const fragment of fragments) {
          expect(surface).not.toContain(fragment);
        }
      }
    }

    const cases: Array<[string, Record<string, unknown> | string, string?]> = [
      ["expired", validClaims({ exp: NOW_S - 100 })],
      ["wrong audience", validClaims({ aud: "https://example.invalid" })],
      ["tenant mismatch", validClaims({ tid: "99999999-2222-3333-4444-555555555555" }), TENANT],
      ["not yet valid", validClaims({ nbf: NOW_S + 86400 })],
      [
        "bad json payload",
        `${b64url({})}.${b64url("SECRET_PAYLOAD_MARKER{")}.FAKE_SIGNATURE_SEGMENT`,
        TENANT,
      ],
      ["opaque with pin", "SECRET_PAYLOAD_MARKER.opaque.token.with.five", TENANT],
      ["non-jwt with marker and pin", "SECRET_PAYLOAD_MARKER", TENANT],
      ["header-unsafe", "SECRET_PAYLOAD_MARKER with space"],
      ["over-length", `SECRET_PAYLOAD_MARKER${"a".repeat(MAX_TOKEN_LENGTH)}`],
    ];

    it.each(cases)("%s", async (_name, claimsOrRaw, tenantId) => {
      const value = typeof claimsOrRaw === "string" ? claimsOrRaw : fakeJwt(claimsOrRaw);
      const opts = tenantId === undefined ? {} : { tenantId };
      const err = await failure(provider({ [ACCESS_TOKEN_ENV_VAR]: value }, opts));
      assertNoLeak(err);
    });
  });
});
