import { execFileSync } from "node:child_process";
import { rmSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import type { Tool } from "@modelcontextprotocol/sdk/types.js";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { loadEndpointCatalog, type EndpointCatalog } from "./catalog.js";
import { z } from "zod";
import { ERROR_CODES, errorResult, successResult } from "./envelope.js";
import {
  RESERVED_INPUT_NAMES,
  UNTRUSTED_CONTENT_NOTICE,
  derivePathParamSchema,
} from "./registry.js";
import { createServer } from "./server.js";

let catalog: EndpointCatalog;
let client: Client;
let tools: Tool[];

async function connect(server = createServer(catalog)): Promise<Client> {
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const c = new Client({ name: "test", version: "0.0.0" });
  await Promise.all([server.connect(serverTransport), c.connect(clientTransport)]);
  return c;
}

beforeAll(async () => {
  catalog = await loadEndpointCatalog();
  client = await connect();
  tools = (await client.listTools()).tools;
});

afterAll(async () => {
  await client.close();
});

describe("tool registration", () => {
  it("registers exactly the catalog tools", () => {
    expect(tools).toHaveLength(76);
    expect(new Set(tools.map((t) => t.name))).toEqual(new Set(catalog.map((e) => e.toolName)));
  });

  it("has no generic proxy-style tool names", () => {
    const pattern = /request|proxy|fetch|http|url/i;
    expect(tools.filter((t) => pattern.test(t.name)).map((t) => t.name)).toEqual([]);
  });

  it("exposes no url/method/path parameters or absolute-URL inputs", () => {
    for (const tool of tools) {
      const props = Object.keys(tool.inputSchema.properties ?? {});
      for (const p of props) {
        expect(RESERVED_INPUT_NAMES).not.toContain(p.toLowerCase());
      }
      expect(JSON.stringify(tool.inputSchema.properties ?? {})).not.toMatch(
        /"format":"uri"|https?:\/\//i,
      );
    }
  });

  it("derives only required string params from path placeholders", () => {
    expect(Object.keys(derivePathParamSchema("/me/mailFolders/{mailFolder-id}/messages"))).toEqual([
      "mailFolder-id",
    ]);
    expect(Object.keys(derivePathParamSchema("/$batch"))).toEqual([]);
  });

  it("marks every tool read-only, non-destructive, open-world", () => {
    for (const tool of tools) {
      expect(tool.annotations).toMatchObject({
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: true,
      });
    }
  });

  it("includes the untrusted-content notice in every description", () => {
    for (const tool of tools) {
      expect(tool.description).toContain(UNTRUSTED_CONTENT_NOTICE);
    }
  });

  it("describes graph-batch as GET-only with relative URLs", () => {
    const batch = tools.find((t) => t.name === "graph-batch");
    expect(batch?.description).toMatch(/only GET/i);
    expect(batch?.description).toMatch(/relative/i);
    expect(batch?.description).not.toMatch(/\bbody\b|POST|PATCH|DELETE|dependsOn/);
  });
});

describe("input validation", () => {
  type Env = { ok: boolean; error: { code: string; message: string; details?: unknown } };
  const call = (name: string, args: Record<string, unknown>) =>
    client.callTool({ name, arguments: args });

  it("returns invalid_input when a required path param is missing", async () => {
    const result = await call("get-mail-message", {});
    expect(result.isError).toBe(true);
    const env = result.structuredContent as Env;
    expect(env.ok).toBe(false);
    expect(env.error.code).toBe("invalid_input");
    expect(env.error.message).toContain("message-id");
  });

  it("returns invalid_input for wrong types and empty strings without echoing values", async () => {
    const long = "SECRET".repeat(50);
    for (const bad of [123, "", { a: long }]) {
      const result = await call("get-mail-message", { "message-id": bad });
      const env = result.structuredContent as Env;
      expect(env.error.code).toBe("invalid_input");
      expect(JSON.stringify(result)).not.toContain("SECRET");
    }
  });

  it("still dispatches valid input to the handler", async () => {
    const result = await call("get-mail-message", { "message-id": "abc" });
    expect((result.structuredContent as Env).error.code).toBe("not_implemented");
  });

  it("keeps catalog path params when a handler supplies its own inputSchema", async () => {
    const handlers = new Map([
      [
        "get-mail-message",
        { inputSchema: { select: z.string().optional() }, handler: () => successResult("ok") },
      ],
    ]);
    const c = await connect(createServer(catalog, handlers));
    const listed = (await c.listTools()).tools.find((t) => t.name === "get-mail-message");
    expect(listed?.inputSchema.required).toEqual(["message-id"]);
    expect(listed?.inputSchema.properties).toHaveProperty("select");
    const bad = await c.callTool({ name: "get-mail-message", arguments: { select: "id" } });
    expect((bad.structuredContent as Env).error.code).toBe("invalid_input");
    const good = await c.callTool({ name: "get-mail-message", arguments: { "message-id": "a" } });
    expect(good.isError).toBeFalsy();
    await c.close();
  });

  it("rejects unknown arguments without echoing values, and advertises additionalProperties:false", async () => {
    const result = await call("get-mail-message", { "message-id": "a", slect: "SECRETVALUE" });
    const env = result.structuredContent as Env;
    expect(result.isError).toBe(true);
    expect(env.error.code).toBe("invalid_input");
    expect(env.error.details).toMatchObject({ unexpectedCount: 1 });
    expect(JSON.stringify(result)).not.toContain("slect");
    expect(JSON.stringify(result)).not.toContain("SECRETVALUE");
    for (const tool of tools) expect(tool.inputSchema.additionalProperties).toBe(false);
  });

  it.each([
    ["a JWT", "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.SECRETKEYMATERIAL"],
    ["a short opaque token", "opaque-token"],
  ])("never echoes an unknown key name, even %s", async (_kind, secretKey) => {
    const result = await call("get-mail-message", { "message-id": "a", [secretKey]: 1 });
    const env = result.structuredContent as Env;
    expect(env.error.code).toBe("invalid_input");
    expect(env.error.details).toMatchObject({ unexpectedCount: 1 });
    expect(JSON.stringify(result)).not.toContain(secretKey);
  });

  it("passes known arguments through to the handler", async () => {
    let seen: unknown;
    const handlers = new Map([
      [
        "get-mail-message",
        {
          inputSchema: { select: z.string().optional() },
          handler: (a: unknown) => {
            seen = a;
            return successResult("ok");
          },
        },
      ],
    ]);
    const c = await connect(createServer(catalog, handlers));
    await c.callTool({ name: "get-mail-message", arguments: { "message-id": "a", select: "id" } });
    expect(seen).toEqual({ "message-id": "a", select: "id" });
    await c.close();
  });

  it("fails server construction when a handler exposes a reserved parameter", () => {
    for (const name of [...RESERVED_INPUT_NAMES, "URL"]) {
      const handlers = new Map([
        [
          "get-mail-message",
          { inputSchema: { [name]: z.string() }, handler: () => successResult("ok") },
        ],
      ]);
      expect(() => createServer(catalog, handlers)).toThrow(/get-mail-message.*"/);
    }
  });

  it("advertises required path params in tools/list", () => {
    const tool = tools.find((t) => t.name === "get-mail-message");
    expect(tool?.inputSchema.required).toEqual(["message-id"]);
    expect(tool?.inputSchema.properties).toHaveProperty("message-id");
    const site = tools.find((t) => t.name === "get-sharepoint-site-by-path");
    expect(site?.inputSchema.required).toContain("relative-path");
  });

  it("never lets a derived-schema tool omit required params in tools/list", () => {
    for (const entry of catalog) {
      const expected = Object.keys(derivePathParamSchema(entry.pathPattern));
      const tool = tools.find((t) => t.name === entry.toolName);
      expect([...(tool?.inputSchema.required ?? [])].sort()).toEqual([...expected].sort());
    }
  });

  it("describes only placeholders that are input properties", () => {
    for (const tool of tools) {
      const props = Object.keys(tool.inputSchema.properties ?? {});
      const refs = [...(tool.description ?? "").matchAll(/{([^{}]+)}/g)].map((m) => m[1]);
      for (const ref of refs) expect(props, `${tool.name} {${ref}}`).toContain(ref);
    }
  });
});

describe("stubs and errors", () => {
  it("returns a structured not_implemented error for unimplemented tools", async () => {
    const result = await client.callTool({
      name: "get-mail-message",
      arguments: { "message-id": "abc" },
    });
    expect(result.isError).toBe(true);
    const envelope = result.structuredContent as { ok: boolean; error: { code: string } };
    expect(envelope.ok).toBe(false);
    expect(envelope.error.code).toBe("not_implemented");
    const first = (result.content as { type: string; text: string }[])[0];
    expect(JSON.parse(first?.text ?? "{}")).toEqual(envelope);
  });

  it("errors cleanly for unknown tools without crashing", async () => {
    const outcome = await client
      .callTool({ name: "definitely-not-a-tool", arguments: {} })
      .then((r) => ({ isError: r.isError === true }))
      .catch(() => ({ isError: true }));
    expect(outcome.isError).toBe(true);
    expect((await client.listTools()).tools).toHaveLength(76);
  });

  it("converts handler exceptions into a sanitized graph_error", async () => {
    const secret = "Bearer eyJ-secret-token";
    const handlers = new Map([
      [
        "get-mail-message",
        {
          handler: () => {
            throw new Error(secret);
          },
        },
      ],
    ]);
    const c = await connect(createServer(catalog, handlers));
    const result = await c.callTool({ name: "get-mail-message", arguments: { "message-id": "x" } });
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result)).not.toContain(secret);
    expect((result.structuredContent as { error: { code: string } }).error.code).toBe(
      "graph_error",
    );
    await c.close();
  });

  it("defines the stable error-code set", () => {
    expect([...ERROR_CODES]).toEqual([
      "not_implemented",
      "invalid_input",
      "reauth_required",
      "capability_gap",
      "rate_limited",
      "graph_error",
      "restart_required",
    ]);
    expect(errorResult("rate_limited", "slow down").isError).toBe(true);
  });
});

describe("stdio entry point (built output)", () => {
  const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
  // Must be a direct child of the repo root: the compiled catalog loader resolves
  // ../endpoints.json relative to its own directory, and node_modules must be reachable.
  const outDir = resolve(root, "dist-smoke");

  beforeAll(() => {
    const tsc = createRequire(import.meta.url).resolve("typescript/bin/tsc");
    execFileSync(
      process.execPath,
      [tsc, "-p", resolve(root, "tsconfig.json"), "--outDir", outDir, "--declaration", "false"],
      { cwd: root, stdio: "pipe" },
    );
  }, 120_000);

  afterAll(() => {
    rmSync(outDir, { recursive: true, force: true });
  });

  it("serves tools/list from compiled index.js in an unrelated working directory", async () => {
    const transport = new StdioClientTransport({
      command: process.execPath,
      args: [resolve(outDir, "index.js")],
      cwd: tmpdir(),
      stderr: "pipe",
    });
    const c = new Client({ name: "smoke", version: "0.0.0" });
    try {
      await c.connect(transport);
      expect((await c.listTools()).tools).toHaveLength(76);
    } finally {
      await c.close();
    }
  }, 30_000);
});
