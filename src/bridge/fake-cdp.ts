import { PassThrough } from "node:stream";

/**
 * Test-only fake CDP peer: stands in for the browser end of `--remote-debugging-pipe`. `input`
 * receives what the bridge writes (browser fd 3); `output` is what the bridge reads (fd 4).
 * Not used by production code.
 */
export interface Received {
  id: number;
  method: string;
  params?: unknown;
  sessionId?: string;
}

export type PeerHandler = (msg: Received) => unknown;

export class FakeCdpPeer {
  readonly input = new PassThrough();
  readonly output = new PassThrough();
  readonly received: Received[] = [];
  readonly handlers = new Map<string, PeerHandler>();
  /** Raw bytes written by the bridge, for "nothing was written" assertions. */
  rawBytes = 0;
  private buf = Buffer.alloc(0);

  constructor() {
    this.input.on("data", (chunk: Buffer) => {
      this.rawBytes += chunk.length;
      this.buf = Buffer.concat([this.buf, chunk]);
      for (;;) {
        const i = this.buf.indexOf(0);
        if (i === -1) return;
        const frame = this.buf.subarray(0, i).toString("utf8");
        this.buf = this.buf.subarray(i + 1);
        const msg = JSON.parse(frame) as Received;
        this.received.push(msg);
        const h = this.handlers.get(msg.method);
        if (h) {
          const result = h(msg);
          this.respond(msg.id, result ?? {}, msg.sessionId);
        }
      }
    });
  }

  on(method: string, handler: PeerHandler): this {
    this.handlers.set(method, handler);
    return this;
  }

  respond(id: number, result: unknown, sessionId?: string): void {
    this.writeFrame({ id, result, ...(sessionId ? { sessionId } : {}) });
  }

  event(method: string, params: unknown, sessionId?: string): void {
    this.writeFrame({ method, params, ...(sessionId ? { sessionId } : {}) });
  }

  writeFrame(obj: unknown): void {
    this.output.write(`${JSON.stringify(obj)}\0`);
  }

  writeRaw(data: Buffer | string): void {
    this.output.write(data);
  }

  eof(): void {
    this.output.end();
  }
}
