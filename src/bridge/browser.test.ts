import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import type { ChildProcess } from "node:child_process";
import { describe, expect, it, vi } from "vitest";
import {
  BrowserError,
  EARLY_EXIT_MS,
  MACOS_BROWSER_PATHS,
  assertProfileFree,
  browserArgs,
  discoverBrowser,
  ensureProfileDir,
  launchBrowser,
} from "./browser.js";
import type { SpawnFn } from "./browser.js";

const file = { isFile: () => true };
const code = (fn: () => unknown): string | undefined => {
  try {
    fn();
  } catch (e) {
    return e instanceof BrowserError ? e.code : "other";
  }
  return undefined;
};

describe("browser discovery", () => {
  it("uses a valid explicit path and never falls back from an invalid one", () => {
    const stat = vi.fn(() => file);
    expect(
      discoverBrowser({
        explicitPath: "/opt/chrome/chrome",
        platform: "darwin",
        statSync: stat,
        accessSync: () => undefined,
      }),
    ).toBe("/opt/chrome/chrome");
    for (const bad of ["relative/chrome", "chrome", "/x\0y"]) {
      expect(
        code(() => discoverBrowser({ explicitPath: bad, platform: "darwin", statSync: stat })),
      ).toBe("browser_path_invalid");
    }
    // Explicit path missing: no silent fallback to the fixed list.
    expect(
      code(() =>
        discoverBrowser({
          explicitPath: "/opt/missing",
          platform: "darwin",
          statSync: (p) => {
            if (p === "/opt/missing") throw new Error("ENOENT");
            return file;
          },
          accessSync: () => undefined,
        }),
      ),
    ).toBe("browser_path_invalid");
    // Directory, or not executable.
    expect(
      code(() =>
        discoverBrowser({
          explicitPath: "/opt/dir",
          platform: "darwin",
          statSync: () => ({ isFile: () => false }),
        }),
      ),
    ).toBe("browser_path_invalid");
    expect(
      code(() =>
        discoverBrowser({
          explicitPath: "/opt/noexec",
          platform: "darwin",
          statSync: () => file,
          accessSync: () => {
            throw new Error("EACCES");
          },
        }),
      ),
    ).toBe("browser_path_invalid");
  });

  it("falls back to the fixed macOS list in order (Chrome, then Edge)", () => {
    expect(MACOS_BROWSER_PATHS).toEqual([
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
      "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    ]);
    const only = (p: string) => (q: string) => {
      if (q !== p) throw new Error("ENOENT");
      return file;
    };
    const ok = { platform: "darwin" as const, accessSync: () => undefined };
    expect(discoverBrowser({ ...ok, statSync: only(MACOS_BROWSER_PATHS[1]!) })).toBe(
      MACOS_BROWSER_PATHS[1],
    );
    expect(discoverBrowser({ ...ok, statSync: () => file })).toBe(MACOS_BROWSER_PATHS[0]);
  });

  it("fails clearly when nothing is found, and never uses the macOS list elsewhere", () => {
    expect(
      code(() =>
        discoverBrowser({
          platform: "darwin",
          statSync: () => {
            throw new Error("ENOENT");
          },
        }),
      ),
    ).toBe("browser_not_found");
    expect(code(() => discoverBrowser({ platform: "linux", statSync: () => file }))).toBe(
      "browser_not_found",
    );
  });

  it("does not echo the configured path in errors", () => {
    try {
      discoverBrowser({ explicitPath: "relative/SECRET-CANARY" });
    } catch (e) {
      expect((e as Error).message).not.toContain("SECRET-CANARY");
    }
  });
});

describe("profile busy detection", () => {
  const link = (t: string) => ({
    readlinkSync: () => t,
  });
  it("reports profile_busy for a lock pointing at a live foreign process", () => {
    expect(code(() => assertProfileFree("/p", link("host-4242"), () => true))).toBe("profile_busy");
  });
  it("ignores stale locks, our own pid, missing locks and odd targets (never deletes)", () => {
    expect(code(() => assertProfileFree("/p", link("host-4242"), () => false))).toBeUndefined();
    expect(
      code(() => assertProfileFree("/p", link(`host-${process.pid}`), () => true)),
    ).toBeUndefined();
    expect(code(() => assertProfileFree("/p", link("garbage"), () => true))).toBeUndefined();
    expect(
      code(() =>
        assertProfileFree(
          "/p",
          {
            readlinkSync: () => {
              throw Object.assign(new Error("ENOENT"), { code: "ENOENT" });
            },
          },
          () => true,
        ),
      ),
    ).toBeUndefined();
  });
});

function fakeChild() {
  const ee = new EventEmitter() as EventEmitter & {
    stdio: unknown[];
    kill: ReturnType<typeof vi.fn>;
  };
  ee.stdio = [null, null, new PassThrough(), new PassThrough(), new PassThrough()];
  ee.kill = vi.fn();
  return ee;
}

describe("browser launch", () => {
  const opts = (spawnFn: SpawnFn, now?: () => number) => ({
    browserPath: "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    profileDir: "/Users/u/Library/Application Support/m365-bridge/profile",
    url: "https://developer.microsoft.com/graph/graph-explorer/",
    env: { PATH: "/usr/bin", HOME: "/Users/u", M365_GRAPH_TENANT_ID: "SECRET-CANARY" },
    spawnFn,
    ...(now ? { now } : {}),
  });

  it("spawns without a shell, with a pipe CDP fd 3/4 layout, a dedicated profile and no debug port", () => {
    const child = fakeChild();
    const spawnFn = vi.fn(() => child as unknown as ChildProcess);
    launchBrowser(opts(spawnFn as unknown as SpawnFn));
    const [command, args, options] = spawnFn.mock.calls[0] as unknown as [
      string,
      string[],
      { shell: boolean; stdio: unknown[]; env: Record<string, string> },
    ];
    expect(command).toContain("Google Chrome");
    expect(options.shell).toBe(false);
    expect(options.stdio).toEqual(["ignore", "ignore", "pipe", "pipe", "pipe"]);
    expect(args).toContain("--remote-debugging-pipe");
    expect(args).toContain(
      "--user-data-dir=/Users/u/Library/Application Support/m365-bridge/profile",
    );
    expect(args).toContain("--no-first-run");
    expect(args).toContain("--no-default-browser-check");
    expect(args.at(-1)).toBe("https://developer.microsoft.com/graph/graph-explorer/");
    expect(args.join(" ")).not.toMatch(/remote-debugging-port|remote-debugging-address/);
    // Bridge configuration is not leaked to the browser environment.
    expect(options.env.M365_GRAPH_TENANT_ID).toBeUndefined();
    expect(options.env.HOME).toBe("/Users/u");
  });

  it("kills the child when the CDP pipes are missing, so no browser keeps holding the profile", () => {
    const child = fakeChild();
    child.stdio = [null, null, new PassThrough(), null, null];
    const spawnFn = vi.fn(() => child as unknown as ChildProcess);
    expect(() => launchBrowser(opts(spawnFn as unknown as SpawnFn))).toThrow(/CDP pipes/);
    expect(child.kill).toHaveBeenCalled();
  });

  it("never emits a debugging-port flag from browserArgs", () => {
    expect(browserArgs("/p", "https://x/").some((a) => a.includes("remote-debugging-port"))).toBe(
      false,
    );
  });

  it("wires fd 3 as input and fd 4 as output, drains stderr", () => {
    const child = fakeChild();
    const proc = launchBrowser(opts((() => child) as unknown as SpawnFn));
    expect(proc.input).toBe(child.stdio[3]);
    expect(proc.output).toBe(child.stdio[4]);
    const stderr = child.stdio[2] as PassThrough;
    expect(stderr.readableFlowing).toBe(true);
  });

  it("flags an early exit (profile busy hand-off) versus a later exit", async () => {
    let t = 1_000;
    const child = fakeChild();
    const proc = launchBrowser(opts((() => child) as unknown as SpawnFn, () => t));
    t += EARLY_EXIT_MS - 1;
    child.emit("exit", 0, null);
    await expect(proc.exited).resolves.toMatchObject({ early: true });

    t = 1_000;
    const child2 = fakeChild();
    const proc2 = launchBrowser(opts((() => child2) as unknown as SpawnFn, () => t));
    t += EARLY_EXIT_MS + 1;
    child2.emit("exit", 0, null);
    await expect(proc2.exited).resolves.toMatchObject({ early: false });
    expect(proc2.hasExited).toBe(true);
  });

  it("reports an asynchronous child error as spawnError, not as a normal exit", async () => {
    const child = fakeChild();
    const proc = launchBrowser(opts((() => child) as unknown as SpawnFn, () => 1_000));
    child.emit("error", new Error("ENOENT"));
    await expect(proc.exited).resolves.toMatchObject({ spawnError: true });
    const child2 = fakeChild();
    const proc2 = launchBrowser(opts((() => child2) as unknown as SpawnFn));
    child2.emit("exit", 0, null);
    await expect(proc2.exited).resolves.toMatchObject({ spawnError: false });
  });

  it("maps spawn failures to spawn_failed without details", () => {
    const err = code(() =>
      launchBrowser(
        opts((() => {
          throw new Error("EACCES SECRET-CANARY");
        }) as unknown as SpawnFn),
      ),
    );
    expect(err).toBe("spawn_failed");
  });

  it("escalates termination SIGTERM then SIGKILL", async () => {
    vi.useFakeTimers();
    const child = fakeChild();
    const proc = launchBrowser(opts((() => child) as unknown as SpawnFn));
    const p = proc.terminate(100);
    await vi.advanceTimersByTimeAsync(250);
    await p;
    expect(child.kill.mock.calls.map((c) => c[0])).toEqual(["SIGTERM", "SIGKILL"]);
    vi.useRealTimers();
  });
});

describe("profile directory safety", () => {
  const pfs = (over: Partial<{ mode: number; uid: number; link: boolean; dir: boolean }> = {}) => ({
    mkdirSync: () => undefined,
    lstatSync: () => ({
      isDirectory: () => over.dir ?? true,
      isSymbolicLink: () => over.link ?? false,
      mode: over.mode ?? 0o40700,
      uid: over.uid ?? 501,
    }),
    readlinkSync: () => "",
  });
  const run = (over: Parameters<typeof pfs>[0], uid: number | undefined = 501) =>
    code(() => ensureProfileDir("/p", "darwin", pfs(over), uid));

  it("accepts a private directory owned by the current user", () => {
    expect(run({})).toBeUndefined();
  });

  it("refuses a directory owned by another user", () => {
    expect(run({ uid: 0 })).toBe("profile_unsafe");
  });

  it("still refuses symlinks, non-directories and group/other-accessible directories", () => {
    expect(run({ link: true })).toBe("profile_unsafe");
    expect(run({ dir: false })).toBe("profile_unsafe");
    expect(run({ mode: 0o40750 })).toBe("profile_unsafe");
    expect(
      code(() => ensureProfileDir("/p", "win32", pfs({ uid: 0, mode: 0o40777 }), 501)),
    ).toBeUndefined();
  });
});
