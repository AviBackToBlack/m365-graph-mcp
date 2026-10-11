import {
  DISABLED_ACCEPTANCE_DIAGNOSTICS,
  createAcceptanceDiagnostics,
} from "./acceptance-diagnostics.js";
import { BridgeBroker, BrokerError, ensureRuntimeDir } from "./broker.js";
import {
  BrowserError,
  assertProfileFree,
  discoverBrowser,
  ensureProfileDir,
  launchBrowser,
} from "./browser.js";
import type { BrowserExit, BrowserProcess, SpawnFn } from "./browser.js";
import { CdpPipe } from "./cdp-pipe.js";
import { GRAPH_EXPLORER_URL, GraphExplorerCapture } from "./capture.js";
import type { CaptureDiagnostic } from "./capture.js";
import { BridgeConfigError, parseBridgeConfig } from "./config.js";
import type { BridgeConfigOptions } from "./config.js";
import path from "node:path";

/** Process exit codes of the bridge daemon. */
export const EXIT = {
  ok: 0,
  config: 2,
  browser: 3,
  profileBusy: 4,
  broker: 5,
  browserExited: 6,
  cdp: 7,
  /** The CDP pipe broke while the bridge was running (browser closed/killed, EOF, oversize frame). */
  cdpClosed: 8,
} as const;

/** After a non-local pipe close, wait this long for the browser's own exit to win the exit code. */
export const CDP_CLOSE_GRACE_MS = 250;

export const PROFILE_DIR_NAME = "profile";

export class BridgeStartError extends Error {
  readonly exitCode: number;
  readonly category: string;
  constructor(exitCode: number, category: string) {
    super(category);
    this.name = "BridgeStartError";
    this.exitCode = exitCode;
    this.category = category;
  }
}

export interface RunBridgeOptions extends BridgeConfigOptions {
  spawnFn?: SpawnFn;
  /** Receives fixed-category diagnostic lines only (never tokens, headers, claims or paths). */
  log?: (line: string) => void;
  now?: () => number;
  /**
   * Aborting (SIGINT/SIGTERM) at any point cancels startup or shuts a running bridge down cleanly:
   * everything created is torn down and the exit code is EXIT.ok.
   */
  signal?: AbortSignal;
}

export interface BridgeHandle {
  /** Resolves with the exit code once the bridge has shut down (browser exit or stop()). */
  readonly done: Promise<number>;
  stop(): Promise<void>;
}

function describe(d: CaptureDiagnostic): string {
  switch (d.kind) {
    case "token_accepted":
      return `token accepted (remaining ${d.remainingSeconds}s)`;
    case "token_rejected":
      return `token rejected (${d.reason})`;
    case "token_cleared":
      return `token cleared (${d.reason})`;
    default:
      return d.kind.replace("_", " ");
  }
}

/**
 * Starts the bridge: config, dirs, broker (so an already-running bridge is detected before another
 * browser is launched), browser, CDP capture. Everything started is torn down on failure.
 * @throws BridgeStartError with a fixed category.
 */
export async function startBridge(options: RunBridgeOptions = {}): Promise<BridgeHandle> {
  const log = options.log ?? (() => undefined);
  const platform = options.platform ?? process.platform;
  const env = options.env ?? process.env;

  let config;
  try {
    config = parseBridgeConfig(options);
  } catch (error) {
    if (error instanceof BridgeConfigError) throw new BridgeStartError(EXIT.config, error.message);
    throw new BridgeStartError(EXIT.config, "invalid configuration");
  }

  const profileDir = path.join(config.runtimeDir, PROFILE_DIR_NAME);
  let browserPath: string;
  try {
    ensureRuntimeDir(config.runtimeDir, platform);
    browserPath = discoverBrowser({ explicitPath: config.browserPath, platform });
    ensureProfileDir(profileDir, platform);
    assertProfileFree(profileDir);
  } catch (error) {
    if (error instanceof BrowserError) {
      throw new BridgeStartError(
        error.code === "profile_busy" ? EXIT.profileBusy : EXIT.browser,
        error.code,
      );
    }
    if (error instanceof BrokerError) throw new BridgeStartError(EXIT.broker, error.code);
    throw new BridgeStartError(EXIT.browser, "runtime_dir_failed");
  }

  const signal = options.signal;
  const interruptedHandle = (): BridgeHandle => ({
    done: Promise.resolve(EXIT.ok),
    stop: () => Promise.resolve(),
  });

  const holder: { capture?: GraphExplorerCapture } = {};
  const broker = new BridgeBroker({
    socketPath: config.socketPath,
    // The broker only ever sees the capture's TokenSource view; it is created lazily below.
    source: {
      current: () => holder.capture?.current() ?? { state: "unavailable" },
      clientPinEnabled: config.gate.clientPin !== undefined,
    },
    platform,
    ...(config.allowDevPipe ? { allowDevPipe: true } : {}),
    ...(options.now ? { now: options.now } : {}),
  });
  try {
    await broker.start();
  } catch (error) {
    if (signal?.aborted) return interruptedHandle();
    throw new BridgeStartError(
      EXIT.broker,
      error instanceof BrokerError ? error.code : "broker_start_failed",
    );
  }
  if (signal?.aborted) {
    await broker.stop(); // interrupted before any browser was launched
    return interruptedHandle();
  }

  let browser: BrowserProcess;
  try {
    browser = launchBrowser({
      browserPath,
      profileDir,
      url: GRAPH_EXPLORER_URL,
      env,
      ...(options.spawnFn ? { spawnFn: options.spawnFn } : {}),
      ...(options.now ? { now: options.now } : {}),
    });
  } catch (error) {
    await broker.stop();
    throw new BridgeStartError(
      EXIT.browser,
      error instanceof BrowserError ? error.code : "spawn_failed",
    );
  }

  // D48: enabled only by the exact opt-in; the sink is the bridge's existing stderr logger.
  const acceptance = config.acceptanceDiagnostics
    ? createAcceptanceDiagnostics(log)
    : DISABLED_ACCEPTANCE_DIAGNOSTICS;
  // Announced before capture starts, so it is always the first acceptance line.
  acceptance.modeEnabled(); // no-op unless enabled
  const cdp = new CdpPipe({ input: browser.input, output: browser.output });
  const cap = new GraphExplorerCapture({
    cdp,
    gate: config.gate,
    ...(options.now ? { now: options.now } : {}),
    onDiagnostic: (d) => log(describe(d)),
    acceptanceDiagnostics: acceptance,
  });
  holder.capture = cap;

  let stopping = false;
  let resolveDone!: (code: number) => void;
  const done = new Promise<number>((r) => (resolveDone = r));
  let exitCode: number = EXIT.ok;

  let finishing: Promise<void> | undefined;
  const doFinish = async (): Promise<void> => {
    cap.stop();
    if (!browser.hasExited && !cdp.isClosed) {
      // Clean shutdown if still connected. The command is not awaited on its own: a browser that
      // never answers must not hold shutdown for the full command timeout before escalation.
      void cdp.send("Browser.close").catch(() => undefined);
      await Promise.race([browser.exited, new Promise((r) => setTimeout(r, 2_000).unref())]);
    }
    cdp.close(); // reason "local": ignored by the close handler below
    await browser.terminate();
    await broker.stop();
    resolveDone(exitCode);
  };
  // Single shutdown path: every trigger shares one in-flight promise (no double shutdown).
  const finish = (): Promise<void> => (finishing ??= doFinish());

  const classifyExit = (exit: BrowserExit): number => {
    if (exit.spawnError) {
      log("browser failed");
      return EXIT.browser;
    }
    log(exit.early ? "browser exited early (profile busy?)" : "browser exited");
    return exit.early ? EXIT.profileBusy : EXIT.browserExited;
  };

  void browser.exited.then((exit) => {
    if (stopping) return;
    stopping = true;
    exitCode = classifyExit(exit);
    void finish();
  });

  const onAbort = (): void => {
    stopping = true;
    void finish();
  };
  signal?.addEventListener("abort", onAbort, { once: true });
  if (signal?.aborted) onAbort();

  let startFailed = false;
  try {
    await cap.start();
  } catch {
    startFailed = true;
  }
  if (signal?.aborted) {
    await finish(); // interrupted during startup: everything created is torn down, clean exit
    return { done, stop: () => Promise.resolve() };
  }
  if (startFailed) {
    stopping = true;
    // If the browser died during setup (e.g. a profile lock hand-off), its exit is the real cause:
    // report profile_busy / browser_exited / browser failure rather than a CDP setup failure.
    const browserExit = browser.hasExited ? await browser.exited : undefined;
    exitCode = browserExit !== undefined ? classifyExit(browserExit) : EXIT.cdp;
    await finish();
    throw new BridgeStartError(
      exitCode,
      browserExit !== undefined ? "browser_exited" : "cdp_setup_failed",
    );
  }

  // A broken pipe (browser closed/killed, EOF, oversize frame) ends the bridge. A deliberate local
  // close during shutdown is ignored. The browser's own exit, if imminent, decides the exit code.
  cdp.onClose((reason) => {
    if (reason === "local" || stopping) return;
    void Promise.race([
      browser.exited,
      new Promise((r) => setTimeout(r, CDP_CLOSE_GRACE_MS).unref()),
    ]).then(() => {
      if (stopping) return;
      stopping = true;
      exitCode = EXIT.cdpClosed;
      log("cdp pipe closed");
      void finish();
    });
  });
  log(`bridge ready (client pin ${cap.clientPinEnabled ? "on" : "off"})`);

  return {
    done,
    async stop() {
      stopping = true;
      await finish();
    },
  };
}
