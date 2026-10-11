import type { Readable, Writable } from "node:stream";

/**
 * Tiny Chrome DevTools Protocol client over `--remote-debugging-pipe` (D42).
 *
 * Wire format (Chromium pipe handler; the same convention other pipe clients use): each message is
 * a JSON object terminated by a single NUL (0x00) byte. The browser reads commands from inherited
 * fd 3 and writes responses/events to inherited fd 4; the caller passes those two streams in.
 *
 * Security shape:
 * - There is NO generic command API. `send` only accepts methods from a fixed allowlist, enforced
 *   before anything is written (cookie, storage, runtime, DOM and page-scraping methods are absent).
 * - Only a fixed set of event methods is ever dispatched; every other event is dropped on receipt.
 * - CDP messages are sensitive transient data. Errors carry only fixed category text and the method
 *   name, never message content; nothing here logs or persists message bodies.
 */

type Scope = "browser" | "session";

/** Allowed commands and whether they are browser-level or must carry a sessionId. */
const ALLOWED_METHODS: ReadonlyMap<string, Scope> = new Map<string, Scope>([
  ["Target.setDiscoverTargets", "browser"],
  ["Target.getTargets", "browser"],
  ["Target.attachToTarget", "browser"],
  ["Target.detachFromTarget", "browser"],
  // Clean shutdown of the bridge-owned browser (avoids profile damage from a hard kill).
  ["Browser.close", "browser"],
  ["Network.enable", "session"],
]);

export const CDP_ALLOWED_METHODS: readonly string[] = [...ALLOWED_METHODS.keys()];

/** Event methods that are dispatched to handlers; anything else is discarded unseen. */
const ALLOWED_EVENTS: ReadonlySet<string> = new Set([
  "Target.targetCreated",
  "Target.targetInfoChanged",
  "Target.targetDestroyed",
  "Target.attachedToTarget",
  "Target.detachedFromTarget",
  "Network.requestWillBeSent",
  "Network.requestWillBeSentExtraInfo",
  "Network.responseReceived",
  "Network.loadingFailed",
]);

/**
 * Maximum bytes of one NUL-terminated message. A Graph Explorer Network event is far below 1 MiB;
 * 8 MiB leaves headroom for Target.getTargets while bounding memory against a runaway peer.
 * Exceeding it closes the pipe (no resynchronisation attempt).
 */
export const DEFAULT_MAX_MESSAGE_BYTES = 8 * 1024 * 1024;
export const DEFAULT_COMMAND_TIMEOUT_MS = 10_000;
export const MAX_IN_FLIGHT_COMMANDS = 64;

export type CdpErrorCode =
  | "forbidden_method"
  | "bad_arguments"
  | "closed"
  | "timeout"
  | "too_many_in_flight"
  | "write_failed"
  | "command_failed";

export class CdpError extends Error {
  readonly code: CdpErrorCode;
  constructor(code: CdpErrorCode, message: string) {
    super(message);
    this.name = "CdpError";
    this.code = code;
  }
}

export type CdpCloseReason = "eof" | "stream_error" | "oversize" | "local";

export interface CdpEvent {
  method: string;
  params: Record<string, unknown>;
  sessionId?: string;
}

export interface CdpClient {
  send(method: string, params?: Record<string, unknown>, sessionId?: string): Promise<unknown>;
  onEvent(handler: (event: CdpEvent) => void): () => void;
}

export interface CdpPipeOptions {
  /** Writable end: the browser's fd 3 (its command input). */
  input: Writable;
  /** Readable end: the browser's fd 4 (its output). */
  output: Readable;
  maxMessageBytes?: number;
  commandTimeoutMs?: number;
}

interface Pending {
  method: string;
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  timer: NodeJS.Timeout;
}

const isRecord = (v: unknown): v is Record<string, unknown> =>
  v !== null && typeof v === "object" && !Array.isArray(v);

export class CdpPipe implements CdpClient {
  private readonly input: Writable;
  private readonly output: Readable;
  private readonly maxBytes: number;
  private readonly timeoutMs: number;
  private nextId = 1;
  private readonly pending = new Map<number, Pending>();
  private readonly eventHandlers = new Set<(event: CdpEvent) => void>();
  private readonly closeHandlers = new Set<(reason: CdpCloseReason) => void>();
  private chunks: Buffer[] = [];
  private chunkBytes = 0;
  private closedReason: CdpCloseReason | undefined;

  constructor(options: CdpPipeOptions) {
    this.input = options.input;
    this.output = options.output;
    this.maxBytes = options.maxMessageBytes ?? DEFAULT_MAX_MESSAGE_BYTES;
    this.timeoutMs = options.commandTimeoutMs ?? DEFAULT_COMMAND_TIMEOUT_MS;
    this.output.on("data", (chunk: Buffer | string) =>
      this.onData(typeof chunk === "string" ? Buffer.from(chunk, "utf8") : chunk),
    );
    this.output.on("end", () => this.fail("eof"));
    this.output.on("close", () => this.fail("eof"));
    this.output.on("error", () => this.fail("stream_error"));
    this.input.on("error", () => this.fail("stream_error"));
    this.input.on("close", () => this.fail("eof"));
  }

  get isClosed(): boolean {
    return this.closedReason !== undefined;
  }

  onEvent(handler: (event: CdpEvent) => void): () => void {
    this.eventHandlers.add(handler);
    return () => this.eventHandlers.delete(handler);
  }

  onClose(handler: (reason: CdpCloseReason) => void): void {
    if (this.closedReason !== undefined) handler(this.closedReason);
    else this.closeHandlers.add(handler);
  }

  /** Sends an allowlisted command. Forbidden methods reject locally; nothing is written. */
  send(method: string, params?: Record<string, unknown>, sessionId?: string): Promise<unknown> {
    const scope = ALLOWED_METHODS.get(method);
    if (scope === undefined) {
      return Promise.reject(new CdpError("forbidden_method", "CDP method is not allowlisted."));
    }
    if (scope === "session") {
      if (typeof sessionId !== "string" || sessionId.length === 0 || params !== undefined) {
        return Promise.reject(
          new CdpError("bad_arguments", `${method} requires a sessionId and takes no parameters.`),
        );
      }
    } else if (sessionId !== undefined) {
      return Promise.reject(new CdpError("bad_arguments", `${method} is a browser-level command.`));
    }
    if (this.closedReason !== undefined) {
      return Promise.reject(new CdpError("closed", "CDP pipe is closed."));
    }
    if (this.pending.size >= MAX_IN_FLIGHT_COMMANDS) {
      return Promise.reject(new CdpError("too_many_in_flight", "Too many CDP commands in flight."));
    }
    const id = this.nextId++;
    const message: Record<string, unknown> = { id, method };
    if (params !== undefined) message.params = params;
    if (sessionId !== undefined) message.sessionId = sessionId;
    let wire: string;
    try {
      wire = JSON.stringify(message);
    } catch {
      return Promise.reject(new CdpError("bad_arguments", "CDP parameters are not serializable."));
    }
    return new Promise<unknown>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new CdpError("timeout", `${method} timed out.`));
      }, this.timeoutMs);
      this.pending.set(id, { method, resolve, reject, timer });
      try {
        this.input.write(`${wire}\0`, (err) => {
          if (err) this.settle(id, new CdpError("write_failed", "CDP write failed."));
        });
      } catch {
        this.settle(id, new CdpError("write_failed", "CDP write failed."));
      }
    });
  }

  /** Locally closes the pipe, rejecting pending commands. Does not stop the browser. */
  close(): void {
    this.fail("local");
  }

  private settle(id: number, error: Error | undefined, value?: unknown): void {
    const p = this.pending.get(id);
    if (p === undefined) return;
    this.pending.delete(id);
    clearTimeout(p.timer);
    if (error) p.reject(error);
    else p.resolve(value);
  }

  private fail(reason: CdpCloseReason): void {
    if (this.closedReason !== undefined) return;
    this.closedReason = reason;
    this.chunks = [];
    this.chunkBytes = 0;
    for (const [id, p] of [...this.pending]) {
      this.pending.delete(id);
      clearTimeout(p.timer);
      p.reject(new CdpError("closed", `CDP pipe closed before ${p.method} completed.`));
    }
    if (reason === "oversize" || reason === "local") {
      this.output.destroy();
      this.input.destroy();
    }
    const handlers = [...this.closeHandlers];
    this.closeHandlers.clear();
    this.eventHandlers.clear();
    for (const h of handlers) {
      try {
        h(reason);
      } catch {
        // A faulty handler must not break teardown.
      }
    }
  }

  private append(part: Buffer): boolean {
    if (this.chunkBytes + part.length > this.maxBytes) {
      this.fail("oversize");
      return false;
    }
    if (part.length > 0) {
      this.chunks.push(part);
      this.chunkBytes += part.length;
    }
    return true;
  }

  private onData(chunk: Buffer): void {
    let start = 0;
    while (this.closedReason === undefined) {
      const nul = chunk.indexOf(0, start);
      if (nul === -1) {
        this.append(chunk.subarray(start));
        return;
      }
      if (!this.append(chunk.subarray(start, nul))) return;
      const frame = this.chunks.length === 1 ? this.chunks[0]! : Buffer.concat(this.chunks);
      this.chunks = [];
      this.chunkBytes = 0;
      start = nul + 1;
      if (frame.length > 0) this.handleFrame(frame);
    }
  }

  private handleFrame(frame: Buffer): void {
    let msg: unknown;
    try {
      msg = JSON.parse(frame.toString("utf8"));
    } catch {
      return; // Malformed JSON is dropped; its content is never surfaced.
    }
    if (!isRecord(msg)) return;
    if (typeof msg.id === "number") {
      if (!this.pending.has(msg.id)) return; // unknown or late id
      if (isRecord(msg.error)) {
        const code = typeof msg.error.code === "number" ? ` (code ${msg.error.code})` : "";
        this.settle(msg.id, new CdpError("command_failed", `CDP command failed${code}.`));
      } else {
        this.settle(msg.id, undefined, isRecord(msg.result) ? msg.result : {});
      }
      return;
    }
    const method = msg.method;
    if (typeof method !== "string" || !ALLOWED_EVENTS.has(method)) return;
    const event: CdpEvent = {
      method,
      params: isRecord(msg.params) ? msg.params : {},
      ...(typeof msg.sessionId === "string" ? { sessionId: msg.sessionId } : {}),
    };
    for (const h of [...this.eventHandlers]) {
      try {
        h(event);
      } catch {
        // Handler faults must not break the read loop.
      }
    }
  }
}
