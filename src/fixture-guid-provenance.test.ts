import { existsSync, readdirSync, readFileSync } from "node:fs";
import { dirname, extname, join, posix, relative, sep, win32 } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const SOURCE_DIR = dirname(fileURLToPath(import.meta.url));
const REPO_DIR = join(SOURCE_DIR, "..");
const UUID_RE =
  /(?<![0-9a-f])[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}(?![0-9a-f])/gi;

/** Documented public Microsoft identifiers; never use personal tenant IDs here. */
const PUBLIC_REFERENCE_IDS = new Set([
  "00000003-0000-0000-c000-000000000000", // Graph application ID
  "00020329-0000-0000-c000-000000000046", // MAPI PS_PUBLIC_STRINGS
  "00062008-0000-0000-c000-000000000046", // MAPI appointment set
  "8eccc264-6880-4ebe-992f-8888d2eeaa1d", // Graph docs extended-property example
]);

const PATTERNED_TEST_SENTINELS = new Set([
  "11111111-2222-3333-4444-555555555555",
  "99999999-8888-7777-6666-555555555555",
  "99999999-2222-3333-4444-555555555555",
  "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
  "cccccccc-dddd-eeee-ffff-000000000000",
  "deadbeef-dead-beef-dead-beefdeadbeef",
  "123e4567-e89b-12d3-a456-426614174000",
  "223e4567-e89b-12d3-a456-426614174000",
  "123e4567-e89b-12d3-a456-426614174999",
]);

/** New deterministic synthetic fixtures authored for the D57 disclosure fix, not from any tenant. */
const ROTATED_FIXTURE_RE = /^f17e0000-0000-4000-8000-0000000000(?:0[1-9]|1[0-9]|20)$/;

const SCANNED_SUFFIXES = new Set([".ts", ".js", ".json", ".md", ".txt", ".yaml", ".yml", ".toml"]);
const EXCLUDED_DIRS = new Set([".git", "node_modules", "dist", "dist-smoke", "coverage"]);

function readableText(path: string): boolean {
  return SCANNED_SUFFIXES.has(extname(path)) || [".gitignore", ".prettierignore"].includes(path);
}

function walk(dir: string): string[] {
  const paths: string[] = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.isDirectory() && !EXCLUDED_DIRS.has(entry.name)) {
      paths.push(...walk(join(dir, entry.name)));
    } else if (entry.isFile() && readableText(entry.name)) {
      paths.push(join(dir, entry.name));
    }
  }
  return paths;
}

interface ManifestEntry {
  action: "KEEP" | "DROP" | "REWRITE";
  replacement?: string;
}

/**
 * In the private engineering repo, scan the prospective public export:
 * KEEP files and REWRITE payloads, not private operator ledgers or audit.
 * Also scan all src/ TS files, including files awaiting export-map re-pin.
 * Once migrated, the manifest is absent: scan checked-in product text.
 */
function publicTextFiles(): string[] {
  const map = join(REPO_DIR, "migration", "export-map.json");
  if (!existsSync(map)) return walk(REPO_DIR).sort();

  const data = JSON.parse(readFileSync(map, "utf8")) as {
    entries: Record<string, ManifestEntry>;
  };
  const files = new Set(walk(SOURCE_DIR));
  for (const [name, entry] of Object.entries(data.entries)) {
    if (entry.action === "DROP") continue;
    const source = entry.action === "REWRITE" ? entry.replacement : name;
    if (source === undefined) throw new Error("Missing replacement for " + name);
    if (readableText(source)) {
      const fullPath = join(REPO_DIR, source);
      if (!existsSync(fullPath)) throw new Error("Missing public export source: " + name);
      files.add(fullPath);
    }
  }
  return [...files].sort();
}

function invalidUuidLines(
  contents: string,
  isTestFile: boolean,
): {
  lines: number[];
  rotated: Set<string>;
} {
  const lines: number[] = [];
  const rotated = new Set<string>();
  for (const [index, line] of contents.split(/\r?\n/).entries()) {
    for (const match of line.matchAll(UUID_RE)) {
      const id = match[0].toLowerCase();
      const isPublicReference = PUBLIC_REFERENCE_IDS.has(id);
      const isSynthetic =
        isTestFile && (PATTERNED_TEST_SENTINELS.has(id) || ROTATED_FIXTURE_RE.test(id));
      if (!isPublicReference && !isSynthetic) lines.push(index + 1);
      if (isTestFile && ROTATED_FIXTURE_RE.test(id)) rotated.add(id);
    }
  }
  return { lines, rotated };
}

/**
 * Recognize test files in the source tree without assuming POSIX paths.
 * Explicit separator boundary also avoids src-other/ being mistaken for src/.
 */
function isSourceTestFile(path: string, root: string, separator: string): boolean {
  return path.startsWith(root + separator) && path.endsWith(".test.ts");
}

describe("public export UUID provenance guard", () => {
  it("recognizes POSIX and Windows test paths without matching sibling directories", () => {
    const unixRoot = posix.join("/", "workspace", "src");
    const windowsRoot = win32.join("C:\\", "workspace", "src");
    expect(isSourceTestFile(posix.join(unixRoot, "tools", "a.test.ts"), unixRoot, posix.sep)).toBe(
      true,
    );
    expect(
      isSourceTestFile(win32.join(windowsRoot, "tools", "a.test.ts"), windowsRoot, win32.sep),
    ).toBe(true);
    expect(
      isSourceTestFile(
        win32.join("C:\\", "workspace", "src-other", "a.test.ts"),
        windowsRoot,
        win32.sep,
      ),
    ).toBe(false);
    expect(isSourceTestFile(win32.join(windowsRoot, "tool.ts"), windowsRoot, win32.sep)).toBe(
      false,
    );
  });

  it("detects unfamiliar identifiers in exported catalog and documentation", () => {
    // Assemble invalid input without embedding a full unknown UUID in source.
    const unknown = ["58c2a109", "a471", "4d66", "b8a3", "100000000001"].join("-");
    for (const target of ["endpoints.json", "README.md"]) {
      expect(invalidUuidLines("Example: " + unknown, false).lines, target).toEqual([1]);
    }
  });

  it("scans the public export including rewritten docs, not private ledgers", () => {
    const files = publicTextFiles();
    expect(files).toContain(join(REPO_DIR, "endpoints.json"));

    const replacement = join(REPO_DIR, "migration", "public-replacements", "README.md");
    if (existsSync(join(REPO_DIR, "migration", "export-map.json"))) {
      expect(files).toContain(replacement);
      const entries = (
        JSON.parse(readFileSync(join(REPO_DIR, "migration", "export-map.json"), "utf8")) as {
          entries: Record<string, ManifestEntry>;
        }
      ).entries;
      // Each public rewrite uses an archive-only replacement payload whose
      // source path is DROP. These nine payloads are intentionally scanned.
      const replacementSources = new Set(
        Object.values(entries)
          .filter((entry) => entry.action === "REWRITE")
          .map((entry) => entry.replacement),
      );
      const excludedPaths = Object.entries(entries)
        .filter(([name, entry]) => entry.action === "DROP" && !replacementSources.has(name))
        .map(([name]) => join(REPO_DIR, name));
      expect(excludedPaths.length).toBeGreaterThan(0);
      for (const excluded of excludedPaths) expect(files).not.toContain(excluded);
    } else {
      expect(files).toContain(join(REPO_DIR, "README.md"));
    }
  });

  it("refuses unfamiliar UUID-shaped literals in every exported text file", () => {
    const violations: string[] = [];
    const rotated = new Set<string>();

    for (const path of publicTextFiles()) {
      const isTestFile = isSourceTestFile(path, SOURCE_DIR, sep);
      const found = invalidUuidLines(readFileSync(path, "utf8"), isTestFile);
      for (const line of found.lines) {
        // Never print possibly private identifiers in CI logs.
        violations.push(relative(REPO_DIR, path) + ":" + String(line));
      }
      for (const id of found.rotated) rotated.add(id);
    }

    expect(violations).toEqual([]);
    expect(rotated.size).toBe(20);
  });
});
