import path from "node:path";

/**
 * Node documents roughly 103 usable bytes for a macOS `sockaddr_un.sun_path`
 * (104 including the terminating NUL). Longer paths fail explicitly.
 */
export const MAX_SOCKET_PATH_BYTES = 103;
/** Windows named-pipe prefix; development/test only, never production. */
export const DEV_PIPE_PREFIX = "\\\\.\\pipe\\";

export class IpcPathError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "IpcPathError";
  }
}

export interface IpcPathOptions {
  /** Windows development only: permit `\\.\pipe\` names. Non-production. */
  allowDevPipe?: boolean;
  /** Platform override for tests; defaults to `process.platform`. */
  platform?: NodeJS.Platform;
}

/** True when `p` is an accepted dev named-pipe name under the given options. */
export function isDevPipe(p: string, options: IpcPathOptions = {}): boolean {
  const platform = options.platform ?? process.platform;
  return options.allowDevPipe === true && platform === "win32" && p.startsWith(DEV_PIPE_PREFIX);
}

/**
 * Asserts that `p` is something the bridge may listen on or connect to: an
 * absolute filesystem path of acceptable length. Ports, host:port strings,
 * relative paths and file descriptors are refused so the bridge can never end
 * up on TCP. Messages never echo the path.
 */
export function assertIpcPath(p: unknown, options: IpcPathOptions = {}): asserts p is string {
  if (typeof p !== "string" || p.length === 0 || p.includes("\0")) {
    throw new IpcPathError("IPC endpoint must be a non-empty filesystem path string.");
  }
  if (isDevPipe(p, options)) return;
  // On Windows, net.Server.listen(<any path>) binds a named pipe, so only the explicitly enabled
  // development pipe is acceptable there; production is a macOS Unix domain socket (D43).
  if ((options.platform ?? process.platform) === "win32") {
    throw new IpcPathError(
      "On Windows only an explicitly enabled development named pipe is supported.",
    );
  }
  if (!path.posix.isAbsolute(p) || p.startsWith("\\\\")) {
    throw new IpcPathError(
      "IPC endpoint must be an absolute filesystem path (never a port or host).",
    );
  }
  if (Buffer.byteLength(p, "utf8") > MAX_SOCKET_PATH_BYTES) {
    throw new IpcPathError(
      `IPC socket path exceeds ${MAX_SOCKET_PATH_BYTES} bytes; use a shorter M365_GRAPH_BRIDGE_DIR.`,
    );
  }
}
