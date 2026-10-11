import fs from "node:fs";
import net from "node:net";
import path from "node:path";
import { assertSafeAncestors } from "../bridge/broker.js";
import { assertIpcPath } from "../bridge/ipc-path.js";
import type { IpcPathOptions } from "../bridge/ipc-path.js";
import { HEADER_SAFE_RE, MAX_TOKEN_LENGTH } from "./jwt-claims.js";
import { TokenError } from "./token-provider.js";
import type { TokenProvider } from "./token-provider.js";

export const DEFAULT_BRIDGE_TIMEOUT_MS = 3_000;
/** Largest accepted response line: a maximal token plus JSON envelope. */
export const MAX_RESPONSE_BYTES = MAX_TOKEN_LENGTH + 1024;

/** Fixed, token-free categories that may appear in error messages. */
const SERVER_CATEGORIES: ReadonlySet<string> = new Set([
  "reauth_required",
  "unavailable",
  "unsupported_token_format",
  "error",
]);
type LocalCategory =
  "connect_failed" | "timeout" | "bad_response" | "response_too_large" | "unsafe_endpoint";

export interface BridgeTokenProviderOptions extends IpcPathOptions {
  socketPath: string;
  timeoutMs?: number;
}

/**
 * TokenProvider backed by the local Graph Explorer token bridge broker. Each
 * call opens a fresh connection and asks for the current token; nothing is
 * cached here (callers already call per request, and the broker owns lifetime).
 * Every failure maps to `reauth_required` with a fixed message carrying only a
 * category; the token and server-supplied text never appear in errors.
 */
export class BridgeTokenProvider implements TokenProvider {
  private readonly socketPath: string;
  private readonly timeoutMs: number;
  private readonly platform: NodeJS.Platform;

  constructor(options: BridgeTokenProviderOptions) {
    try {
      assertIpcPath(options.socketPath, options);
    } catch {
      throw new TokenError("auth_misconfigured", "Token bridge socket path is invalid.");
    }
    this.socketPath = options.socketPath;
    this.timeoutMs = options.timeoutMs ?? DEFAULT_BRIDGE_TIMEOUT_MS;
    this.platform = options.platform ?? process.platform;
  }

  async getAccessToken(): Promise<string> {
    let line: string;
    try {
      this.verifyEndpoint();
      line = await this.request();
    } catch (category) {
      throw this.fail(typeof category === "string" ? category : "connect_failed");
    }
    let body: unknown;
    try {
      body = JSON.parse(line);
    } catch {
      throw this.fail("bad_response");
    }
    if (body === null || typeof body !== "object" || Array.isArray(body)) {
      throw this.fail("bad_response");
    }
    const { status, token } = body as { status?: unknown; token?: unknown };
    if (status === "ok") {
      if (
        typeof token === "string" &&
        token.length > 0 &&
        token.length <= MAX_TOKEN_LENGTH &&
        HEADER_SAFE_RE.test(token)
      ) {
        return token;
      }
      throw this.fail("bad_response");
    }
    throw this.fail(
      typeof status === "string" && SERVER_CATEGORIES.has(status) ? status : "bad_response",
    );
  }

  /**
   * Mirrors the broker's boundary before connecting, so a socket planted by another user under a
   * modifiable directory is never trusted to supply a token: the runtime directory must be a
   * plain 0700 directory owned by this user with safe ancestors, and the socket must be a socket
   * owned by this user. Throws a category string (caught by getAccessToken). POSIX only; the
   * Windows dev pipe has no filesystem boundary.
   */
  private verifyEndpoint(): void {
    if (this.platform === "win32") return;
    const uid = typeof process.getuid === "function" ? process.getuid() : undefined;
    if (uid === undefined) return;
    const dir = path.dirname(this.socketPath);
    let socketStat: fs.Stats;
    try {
      const dirStat = fs.lstatSync(dir);
      if (
        dirStat.isSymbolicLink() ||
        !dirStat.isDirectory() ||
        dirStat.uid !== uid ||
        (dirStat.mode & 0o077) !== 0
      ) {
        throw "unsafe_endpoint" satisfies LocalCategory;
      }
      socketStat = fs.lstatSync(this.socketPath);
    } catch (error) {
      if (error === "unsafe_endpoint") throw error;
      throw "connect_failed" satisfies LocalCategory; // bridge not running / dir missing
    }
    if (!socketStat.isSocket() || socketStat.uid !== uid) {
      throw "unsafe_endpoint" satisfies LocalCategory;
    }
    try {
      assertSafeAncestors(dir);
    } catch {
      throw "unsafe_endpoint" satisfies LocalCategory;
    }
  }

  private fail(category: LocalCategory | string): TokenError {
    return new TokenError(
      "reauth_required",
      `Graph token bridge did not provide a token (${category}); sign in through the bridge's Graph Explorer session and retry.`,
    );
  }

  private request(): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      const chunks: Buffer[] = [];
      let size = 0;
      let settled = false;
      const socket = net.createConnection({ path: this.socketPath });
      const finish = (err: LocalCategory | undefined, value?: string): void => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        socket.destroy();
        if (err) reject(err);
        else resolve(value ?? "");
      };
      const timer = setTimeout(() => finish("timeout"), this.timeoutMs);
      socket.once("connect", () => socket.write('{"op":"getToken"}\n'));
      socket.on("error", () => finish("connect_failed"));
      socket.on("data", (chunk: Buffer) => {
        chunks.push(chunk);
        size += chunk.length;
        const all = Buffer.concat(chunks);
        const nl = all.indexOf(0x0a);
        if (nl !== -1) {
          if (nl > MAX_RESPONSE_BYTES) finish("response_too_large");
          else finish(undefined, all.subarray(0, nl).toString("utf8"));
        } else if (size > MAX_RESPONSE_BYTES) {
          finish("response_too_large");
        }
      });
      socket.on("end", () => finish("bad_response"));
    });
  }
}
