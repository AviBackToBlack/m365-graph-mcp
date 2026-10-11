import { readFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { describe, expect, it } from "vitest";
import { DEFAULT_CATALOG_PATH, loadEndpointCatalog } from "./catalog.js";

const EXPECTED_TOOL_COUNT = 76; // D21 removed list-chat-message-replies; D32 the two site items tools; D36 sites delta.
const EXPECTED_POST_TOOLS = new Set([
  "graph-batch",
  "get-schedule",
  "find-meeting-times",
  "search-query",
  "copilot-retrieve",
]);

describe("selected endpoint catalog", () => {
  it("parses the normative catalog and preserves the exact selected surface", async () => {
    const catalog = await loadEndpointCatalog();
    const names = catalog.map((entry) => entry.toolName);

    expect(catalog).toHaveLength(EXPECTED_TOOL_COUNT);
    expect(new Set(names).size).toBe(EXPECTED_TOOL_COUNT);
    expect(catalog.filter((entry) => entry.method === "get")).toHaveLength(71);
    expect(catalog.filter((entry) => entry.method === "post")).toHaveLength(5);
    expect(
      new Set(catalog.filter((entry) => entry.method === "post").map((entry) => entry.toolName)),
    ).toEqual(EXPECTED_POST_TOOLS);
  });

  it("keeps catalog-summary.json in step with the catalog", async () => {
    const catalog = await loadEndpointCatalog();
    const summary = JSON.parse(
      await readFile(join(dirname(DEFAULT_CATALOG_PATH), "catalog-summary.json"), "utf8"),
    ) as {
      tool_count: number;
      methods: { GET: number; POST: number };
      post_tools: { toolName: string }[];
    };
    expect(summary.tool_count).toBe(catalog.length);
    expect(summary.methods.GET).toBe(catalog.filter((e) => e.method === "get").length);
    expect(summary.methods.POST).toBe(catalog.filter((e) => e.method === "post").length);
    expect(new Set(summary.post_tools.map((t) => t.toolName))).toEqual(EXPECTED_POST_TOOLS);
  });

  it("contains no mutation HTTP verbs", async () => {
    const catalog = await loadEndpointCatalog();
    expect(catalog.every((entry) => entry.method === "get" || entry.method === "post")).toBe(true);
  });

  it("uses the accepted least-privilege scope for get-schedule", async () => {
    const catalog = await loadEndpointCatalog();
    const entry = catalog.find((item) => item.toolName === "get-schedule");
    expect(entry?.workScopes).toEqual(["Calendars.ReadBasic"]);
  });
});
