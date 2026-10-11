import fs from "node:fs";
import net from "node:net";
import path from "node:path";
import { assertIpcPath, isDevPipe } from "./ipc-path.js";
import type { IpcPathOptions } from "./ipc-path.js";
import { MIN_REMAINING_MS } from "./token-gate.js";
import type { RejectReason } from "./token-gate.js";

/**
 * Local-only, user-restricted IPC broker (D43). Unix domain socket only; never
 * TCP. Filesystem ownership/permissions of the 0700 runtime directory and the
 * 0600 socket are the boundary. There is deliberately NO per-start shared
 * secret: a secret readable by the owning user would not protect against a
 * process already running as the same OS user, and an in-memory secret would
 * need a launch handoff to the MCP process that part A does not have. The
 * broker exposes no browser/CDP control and only the operations below.
 *
 * Protocol: one newline-terminated JSON request per connection, one
 * newline-terminated JSON response, then the connection closes.
 *   {"op":"getToken"} -> {"status":"ok","token":"...","expiresInSeconds":N}
 *                      | {"status":"reauth_required"|"unavailable"|
 *                         "unsupported_token_format","reason"?:"<category>"}
 *   {"op":"status"}   -> {"status":"ok","state":"ok"|<failure>,
 *                         "remainingSeconds"?:N,"clientPinEnabled":bool}
 *   anything else     -> {"status":"error","reason":"unknown_op"|"bad_request"|
 *                         "request_too_large"}
 */

export type BrokerFailureState = "reauth_required" | "unavailable" | "unsupported_token_format";

export type TokenSnapshot =
  | { state: "ok"; token: string; expMs: number }
  | { state: BrokerFailureState; reason?: RejectReason };

/** Supplies the current accepted token (part B backs this with CDP + the gate). */
export interface TokenSource {
  current(): TokenSnapshot;
  readonly clientPinEnabled: boolean;
}

/** In-memory source for tests and for part B to drive. */
export class InMemoryTokenSource implements TokenSource {
  private snapshot: TokenSnapshot = { state: "unavailable" };
  constructor(readonly clientPinEnabled = false) {}
  set(snapshot: TokenSnapshot): void {
    this.snapshot = snapshot;
  }
  current(): TokenSnapshot {
    return this.snapshot;
  }
}

export const MAX_REQUEST_BYTES = 1024;
export const DEFAULT_CONNECTION_TIMEOUT_MS = 5_000;
/**
 * Tokens with less than this remaining are not served. D44 requires at least five minutes before
 * serving, so the gate's acceptance floor is re-applied at serve time as the accepted token ages.
 */
export const SERVE_MIN_REMAINING_MS = MIN_REMAINING_MS;
const MAX_CONNECTIONS = 16;

export class BrokerError extends Error {
  readonly code: "already_running" | "unsafe_path";
  constructor(code: "already_running" | "unsafe_path", message: string) {
    super(message);
    this.name = "BrokerError";
    this.code = code;
  }
}

export interface BrokerOptions extends IpcPathOptions {
  socketPath: string;
  source: TokenSource;
  now?: () => number;
  connectionTimeoutMs?: number;
}

const isWin = (platform: NodeJS.Platform): boolean => platform === "win32";

/**
 * Creates the runtime directory (mode 0700) if needed and verifies it: a real
 * directory (not a symlink), owned by the current uid, with no group/other bits.
 */
export function ensureRuntimeDir(dir: string, platform: NodeJS.Platform = process.platform): void {
  fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  const st = fs.lstatSync(dir);
  if (st.isSymbolicLink() || !st.isDirectory()) {
    throw new BrokerError("unsafe_path", "Bridge runtime directory is not a plain directory.");
  }
  if (isWin(platform)) return; // POSIX ownership/mode do not apply (dev only).
  assertSafeAncestors(dir);
  if (typeof process.getuid === "function" && st.uid !== process.getuid()) {
    throw new BrokerError("unsafe_path", "Bridge runtime directory is not owned by this user.");
  }
  if ((st.mode & 0o077) !== 0) {
    throw new BrokerError(
      "unsafe_path",
      "Bridge runtime directory is accessible by group/others; refusing to use it.",
    );
  }
}

/**
 * Every ancestor of the runtime directory, along both the given path and its resolved real path,
 * must be owned by this user or root and must not be group/other-writable unless sticky (as /tmp).
 * Otherwise another user could swap an ancestor (for example a symlink) and relocate the socket.
 * Symlink entries themselves are skipped: the directory containing each is checked as the next
 * ancestor, and their targets are covered by the real-path chain.
 */
export function assertSafeAncestors(dir: string): void {
  const uid = typeof process.getuid === "function" ? process.getuid() : undefined;
  if (uid === undefined) return;
  let real: string;
  try {
    real = fs.realpathSync(dir);
  } catch {
    throw new BrokerError("unsafe_path", "Bridge runtime directory could not be resolved.");
  }
  for (const start of [path.resolve(dir), real]) {
    let current = path.dirname(start);
    for (;;) {
      let st: fs.Stats;
      try {
        st = fs.lstatSync(current);
      } catch {
        throw new BrokerError("unsafe_path", "A parent of the bridge runtime directory vanished.");
      }
      if (!st.isSymbolicLink()) {
        if (st.uid !== uid && st.uid !== 0) {
          throw new BrokerError(
            "unsafe_path",
            "A parent of the bridge runtime directory is owned by another user.",
          );
        }
        if ((st.mode & 0o022) !== 0 && (st.mode & 0o1000) === 0) {
          throw new BrokerError(
            "unsafe_path",
            "A parent of the bridge runtime directory is writable by group/others.",
          );
        }
      }
      const parent = path.dirname(current);
      if (parent === current) break;
      current = parent;
    }
  }
}

function probeSocket(socketPath: string): Promise<"live" | "stale"> {
  return new Promise((resolve, reject) => {
    const sock = net.createConnection({ path: socketPath });
    const timer = setTimeout(() => {
      sock.destroy();
      reject(new BrokerError("unsafe_path", "Existing bridge socket did not answer a probe."));
    }, 1_000);
    sock.once("connect", () => {
      clearTimeout(timer);
      sock.destroy();
      resolve("live");
    });
    sock.once("error", (err: NodeJS.ErrnoException) => {
      clearTimeout(timer);
      sock.destroy();
      if (err.code === "ECONNREFUSED" || err.code === "ENOENT") resolve("stale");
      else reject(new BrokerError("unsafe_path", "Existing bridge socket could not be probed."));
    });
  });
}

/** Refuses unsafe pre-existing endpoints; removes a stale socket; errors on a live one. */
async function prepareSocketPath(socketPath: string): Promise<void> {
  let st: fs.Stats;
  try {
    st = fs.lstatSync(socketPath);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
    throw new BrokerError("unsafe_path", "Bridge socket path could not be inspected.");
  }
  if (st.isSymbolicLink() || !st.isSocket()) {
    throw new BrokerError("unsafe_path", "Bridge socket path exists and is not a plain socket.");
  }
  if (typeof process.getuid === "function" && st.uid !== process.getuid()) {
    throw new BrokerError("unsafe_path", "Existing bridge socket is not owned by this user.");
  }
  if ((await probeSocket(socketPath)) === "live") {
    throw new BrokerError("already_running", "Token bridge is already running.");
  }
  try {
    fs.unlinkSync(socketPath);
  } catch (error) {
    // The stale socket may have been removed concurrently (e.g. by its dying owner).
    if ((error as NodeJS.ErrnoException).code !== "ENOENT") {
      throw new BrokerError("unsafe_path", "Stale bridge socket could not be removed.");
    }
  }
}

type ErrorReason = "unknown_op" | "bad_request" | "request_too_large";

export class BridgeBroker {
  private server: net.Server | undefined;
  private readonly sockets = new Set<net.Socket>();
  private readonly now: () => number;
  private readonly timeoutMs: number;
  private readonly fsMode: boolean;
  /** Identity of the socket this broker bound, so stop() never unlinks a successor's socket. */
  private boundIdentity: { dev: number; ino: number } | undefined;

  constructor(private readonly options: BrokerOptions) {
    this.now = options.now ?? Date.now;
    this.timeoutMs = options.connectionTimeoutMs ?? DEFAULT_CONNECTION_TIMEOUT_MS;
    this.fsMode = !isDevPipe(options.socketPath, options);
  }

  /** Starts listening. Rejects without binding on any unsafe condition. */
  async start(): Promise<void> {
    const { socketPath } = this.options;
    assertIpcPath(socketPath, this.options);
    const platform = this.options.platform ?? process.platform;
    if (this.fsMode) {
      ensureRuntimeDir(path.dirname(socketPath), platform);
      await prepareSocketPath(socketPath);
    }
    const server = net.createServer((socket) => this.handle(socket));
    server.maxConnections = MAX_CONNECTIONS;
    await new Promise<void>((resolve, reject) => {
      server.once("error", reject);
      // A string path only: this is a Unix domain socket, never a TCP port.
      server.listen(socketPath, () => {
        server.off("error", reject);
        resolve();
      });
    });
    this.server = server;
    if (this.fsMode && !isWin(platform)) {
      try {
        fs.chmodSync(socketPath, 0o600);
        const st = fs.lstatSync(socketPath);
        if (!st.isSocket() || (st.mode & 0o077) !== 0) {
          throw new BrokerError("unsafe_path", "Bridge socket permissions could not be secured.");
        }
        this.boundIdentity = { dev: st.dev, ino: st.ino };
      } catch (error) {
        await this.stop();
        throw error instanceof BrokerError
          ? error
          : new BrokerError("unsafe_path", "Bridge socket permissions could not be secured.");
      }
    }
  }

  async stop(): Promise<void> {
    const server = this.server;
    this.server = undefined;
    if (server === undefined) return;
    for (const s of this.sockets) s.destroy();
    await new Promise<void>((resolve) => server.close(() => resolve()));
    const identity = this.boundIdentity;
    this.boundIdentity = undefined;
    // Node/libuv normally removes the socket on close. Only remove a leftover pathname if it is
    // still the socket this broker bound; a replacement broker may already own the path.
    if (this.fsMode && identity !== undefined) {
      try {
        const st = fs.lstatSync(this.options.socketPath);
        if (st.isSocket() && st.dev === identity.dev && st.ino === identity.ino) {
          fs.unlinkSync(this.options.socketPath);
        }
      } catch {
        // Already removed by close().
      }
    }
  }

  private handle(socket: net.Socket): void {
    this.sockets.add(socket);
    socket.once("close", () => this.sockets.delete(socket));
    let buf = Buffer.alloc(0);
    let done = false;
    const respond = (body: unknown): void => {
      if (done) return;
      done = true;
      socket.end(`${JSON.stringify(body)}\n`);
    };
    socket.setTimeout(this.timeoutMs, () => {
      done = true;
      socket.destroy();
    });
    socket.on("error", () => socket.destroy());
    socket.on("data", (chunk: Buffer) => {
      if (done) return;
      // Bound before buffering: only the bytes up to the first newline count, and they are
      // checked against the limit before any concatenation, so a huge chunk is never copied.
      const nl = chunk.indexOf(0x0a);
      const part = nl === -1 ? chunk : chunk.subarray(0, nl);
      if (buf.length + part.length > MAX_REQUEST_BYTES) {
        respond({ status: "error", reason: "request_too_large" satisfies ErrorReason });
        return;
      }
      buf = Buffer.concat([buf, part]);
      if (nl === -1) return;
      let body: unknown;
      try {
        body = this.dispatch(buf.toString("utf8"));
      } catch {
        // A throwing token source (or clock) must never crash the bridge daemon; report a fixed,
        // token-free unavailable state instead.
        body = { status: "unavailable" satisfies BrokerFailureState };
      }
      try {
        respond(body);
      } catch {
        socket.destroy(); // e.g. an unserializable body; never crash the daemon
      }
    });
  }

  private dispatch(line: string): unknown {
    let req: unknown;
    try {
      req = JSON.parse(line);
    } catch {
      return { status: "error", reason: "bad_request" satisfies ErrorReason };
    }
    if (req === null || typeof req !== "object" || Array.isArray(req)) {
      return { status: "error", reason: "bad_request" satisfies ErrorReason };
    }
    const keys = Object.keys(req);
    const op = (req as { op?: unknown }).op;
    if (typeof op !== "string" || keys.length !== 1) {
      return { status: "error", reason: "bad_request" satisfies ErrorReason };
    }
    if (op === "getToken") return this.getToken();
    if (op === "status") return this.status();
    return { status: "error", reason: "unknown_op" satisfies ErrorReason };
  }

  /** Current snapshot, downgraded to reauth_required when too close to expiry. */
  private effective(): TokenSnapshot & { remainingMs?: number } {
    const snap = this.options.source.current();
    if (snap.state !== "ok") return snap;
    // Defend against a source that violates TokenSnapshot at runtime: never serve a non-string
    // token or a non-numeric expiry.
    if (typeof snap.token !== "string" || typeof snap.expMs !== "number") {
      return { state: "unavailable" };
    }
    const remainingMs = snap.expMs - this.now();
    if (!Number.isFinite(remainingMs) || remainingMs < SERVE_MIN_REMAINING_MS) {
      return { state: "reauth_required", reason: "lifetime" };
    }
    return { ...snap, remainingMs };
  }

  private getToken(): unknown {
    const snap = this.effective();
    if (snap.state === "ok") {
      return {
        status: "ok",
        token: snap.token,
        expiresInSeconds: Math.floor((snap.remainingMs ?? 0) / 1000),
      };
    }
    return { status: snap.state, ...(snap.reason ? { reason: snap.reason } : {}) };
  }

  private status(): unknown {
    const snap = this.effective();
    return {
      status: "ok",
      state: snap.state,
      ...(snap.state === "ok"
        ? { remainingSeconds: Math.floor((snap.remainingMs ?? 0) / 1000) }
        : {}),
      ...(snap.state !== "ok" && snap.reason ? { reason: snap.reason } : {}),
      clientPinEnabled: this.options.source.clientPinEnabled,
    };
  }
}
