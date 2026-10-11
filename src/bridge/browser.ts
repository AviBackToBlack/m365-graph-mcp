import { spawn as nodeSpawn } from "node:child_process";
import type { ChildProcess } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import type { Readable, Writable } from "node:stream";

/**
 * Browser discovery and launch for the bridge (D42). The bridge owns the browser it launches: a
 * system Chrome/Edge with a dedicated non-default `--user-data-dir` and `--remote-debugging-pipe`
 * (CDP on inherited fd 3 = browser input, fd 4 = browser output). Never `--remote-debugging-port`,
 * never a shell, never attaching to an already-running regular browser.
 */

export const MACOS_BROWSER_PATHS: readonly string[] = [
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
];

/** A browser exiting sooner than this after launch is treated as "profile busy" (instance hand-off). */
export const EARLY_EXIT_MS = 5_000;

export type BrowserErrorCode =
  "browser_not_found" | "browser_path_invalid" | "profile_unsafe" | "profile_busy" | "spawn_failed";

export class BrowserError extends Error {
  readonly code: BrowserErrorCode;
  constructor(code: BrowserErrorCode, message: string) {
    super(message);
    this.name = "BrowserError";
    this.code = code;
  }
}

export interface DiscoverOptions {
  explicitPath?: string | undefined;
  platform?: NodeJS.Platform;
  statSync?: (p: string) => { isFile(): boolean };
  accessSync?: (p: string, mode: number) => void;
}

/**
 * Explicit path (absolute, existing file, executable on POSIX) wins and is never silently replaced
 * by the fixed list. Otherwise the first existing fixed macOS location is used.
 * @throws BrowserError with a message that never echoes configured paths.
 */
export function discoverBrowser(options: DiscoverOptions = {}): string {
  const platform = options.platform ?? process.platform;
  const stat = options.statSync ?? ((p: string) => fs.statSync(p));
  const access = options.accessSync ?? ((p: string, m: number) => fs.accessSync(p, m));
  const usable = (p: string): boolean => {
    try {
      if (!stat(p).isFile()) return false;
      if (platform !== "win32") access(p, fs.constants.X_OK);
      return true;
    } catch {
      return false;
    }
  };
  const explicit = options.explicitPath;
  if (explicit !== undefined) {
    if (!path.isAbsolute(explicit) || explicit.includes("\0") || !usable(explicit)) {
      throw new BrowserError(
        "browser_path_invalid",
        "Configured browser path is not an absolute path to an executable file.",
      );
    }
    return explicit;
  }
  if (platform === "darwin") {
    for (const candidate of MACOS_BROWSER_PATHS) if (usable(candidate)) return candidate;
  }
  throw new BrowserError(
    "browser_not_found",
    "No supported browser found (Google Chrome or Microsoft Edge); set M365_GRAPH_BROWSER_PATH.",
  );
}

export interface ProfileFs {
  mkdirSync(p: string, o: { recursive: true; mode: number }): unknown;
  lstatSync(p: string): {
    isDirectory(): boolean;
    isSymbolicLink(): boolean;
    mode: number;
    uid: number;
  };
  readlinkSync(p: string): string;
}

/**
 * Creates the dedicated profile directory (0700) and refuses symlinks/non-directories, directories
 * owned by another user and group/other-accessible directories (mirrors ensureRuntimeDir).
 */
export function ensureProfileDir(
  profileDir: string,
  platform: NodeJS.Platform = process.platform,
  pfs: ProfileFs = fs,
  uid: number | undefined = typeof process.getuid === "function" ? process.getuid() : undefined,
): void {
  pfs.mkdirSync(profileDir, { recursive: true, mode: 0o700 });
  const st = pfs.lstatSync(profileDir);
  if (st.isSymbolicLink() || !st.isDirectory()) {
    throw new BrowserError(
      "profile_unsafe",
      "Bridge browser profile path is not a plain directory.",
    );
  }
  if (platform !== "win32" && uid !== undefined && st.uid !== uid) {
    throw new BrowserError(
      "profile_unsafe",
      "Bridge browser profile directory is not owned by this user.",
    );
  }
  if (platform !== "win32" && (st.mode & 0o077) !== 0) {
    throw new BrowserError("profile_unsafe", "Bridge browser profile directory is not private.");
  }
}

export function isPidAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "EPERM";
  }
}

/**
 * Chrome on macOS/Linux records its owner in `<profile>/SingletonLock`, a symlink to
 * `<host>-<pid>`. A lock pointing to a live process means the profile is in use. The lock is never
 * deleted or bypassed; a stale lock is left for the browser itself to reclaim.
 * @throws BrowserError("profile_busy")
 */
export function assertProfileFree(
  profileDir: string,
  pfs: Pick<ProfileFs, "readlinkSync"> = fs,
  pidAlive: (pid: number) => boolean = isPidAlive,
): void {
  let target: string;
  try {
    target = pfs.readlinkSync(path.join(profileDir, "SingletonLock"));
  } catch {
    return; // no lock (or not a symlink): nothing to report
  }
  const m = /-(\d{1,10})$/.exec(target);
  const pid = m ? Number(m[1]) : undefined;
  if (pid !== undefined && pid > 0 && pid !== process.pid && pidAlive(pid)) {
    throw new BrowserError("profile_busy", "Bridge browser profile is in use by another process.");
  }
}

export function browserArgs(profileDir: string, url: string): string[] {
  return [
    `--user-data-dir=${profileDir}`,
    "--remote-debugging-pipe",
    "--no-first-run",
    "--no-default-browser-check",
    url,
  ];
}

export interface BrowserExit {
  code: number | null;
  signal: NodeJS.Signals | null;
  /** True when the browser exited within EARLY_EXIT_MS of launch (likely profile hand-off/busy). */
  early: boolean;
  /** True when the ChildProcess reported an asynchronous 'error' (spawn failure), not an exit. */
  spawnError: boolean;
}

export type SpawnFn = (
  command: string,
  args: readonly string[],
  options: {
    shell: false;
    stdio: ["ignore", "ignore", "pipe", "pipe", "pipe"];
    env: NodeJS.ProcessEnv;
    windowsHide: true;
  },
) => ChildProcess;

export interface LaunchOptions {
  browserPath: string;
  profileDir: string;
  url: string;
  env?: NodeJS.ProcessEnv;
  spawnFn?: SpawnFn;
  now?: () => number;
}

export class BrowserProcess {
  /** Browser's fd 3: write CDP commands here. */
  readonly input: Writable;
  /** Browser's fd 4: read CDP responses/events here. */
  readonly output: Readable;
  readonly exited: Promise<BrowserExit>;
  private exitInfo: BrowserExit | undefined;

  constructor(
    private readonly child: ChildProcess,
    now: () => number,
  ) {
    const started = now();
    this.input = child.stdio[3] as Writable;
    this.output = child.stdio[4] as Readable;
    // stderr is drained and discarded unread: browser diagnostics can contain URLs and are never
    // logged by the bridge (an unread pipe would also eventually block the browser).
    (child.stdio[2] as Readable | null)?.on("error", () => undefined);
    (child.stdio[2] as Readable | null)?.resume();
    this.exited = new Promise<BrowserExit>((resolve) => {
      const done = (
        code: number | null,
        signal: NodeJS.Signals | null,
        spawnError: boolean,
      ): void => {
        if (this.exitInfo !== undefined) return;
        this.exitInfo = { code, signal, early: now() - started < EARLY_EXIT_MS, spawnError };
        resolve(this.exitInfo);
      };
      child.once("exit", (code, signal) => done(code, signal, false));
      child.once("error", () => done(null, null, true));
    });
  }

  get hasExited(): boolean {
    return this.exitInfo !== undefined;
  }

  /** Escalating stop: caller should first try CDP `Browser.close`; then SIGTERM, then SIGKILL. */
  async terminate(graceMs = 3_000): Promise<void> {
    if (this.hasExited) return;
    const wait = (ms: number): Promise<"timeout"> =>
      new Promise((r) => setTimeout(() => r("timeout"), ms).unref());
    this.child.kill("SIGTERM");
    if ((await Promise.race([this.exited, wait(graceMs)])) === "timeout") {
      this.child.kill("SIGKILL");
      await Promise.race([this.exited, wait(graceMs)]);
    }
  }
}

/** Launches the browser. @throws BrowserError("spawn_failed") without echoing details. */
export function launchBrowser(options: LaunchOptions): BrowserProcess {
  const spawnFn: SpawnFn = options.spawnFn ?? (nodeSpawn as unknown as SpawnFn);
  // The bridge's own configuration is not handed to the browser.
  // Object.fromEntries avoids prototype setters for externally supplied environment keys.
  // Ignore object-meta keys as well as bridge secrets before handing the environment to Chrome.
  const env = Object.fromEntries(
    Object.entries(options.env ?? process.env).filter(
      ([key]) =>
        !key.startsWith("M365_GRAPH_") &&
        key !== "__proto__" &&
        key !== "constructor" &&
        key !== "prototype",
    ),
  ) as NodeJS.ProcessEnv;
  let child: ChildProcess;
  try {
    child = spawnFn(options.browserPath, browserArgs(options.profileDir, options.url), {
      shell: false,
      stdio: ["ignore", "ignore", "pipe", "pipe", "pipe"],
      env,
      windowsHide: true,
    });
  } catch {
    throw new BrowserError("spawn_failed", "Browser could not be started.");
  }
  if (!child.stdio[3] || !child.stdio[4]) {
    // Do not leave a browser holding the profile (it would read as profile_busy next run).
    try {
      child.kill("SIGKILL");
    } catch {
      // already gone
    }
    throw new BrowserError("spawn_failed", "Browser CDP pipes were not created.");
  }
  return new BrowserProcess(child, options.now ?? Date.now);
}
