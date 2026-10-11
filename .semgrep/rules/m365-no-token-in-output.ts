declare const token: string;
declare const accessToken: string;
declare const ctx: { bearer: string; auth: { secret: string }; count: number };
declare const clientSecret: string;
declare const authorization: string;
declare const status: string;
declare const line: string;

function leaks(): void {
  // ruleid: m365-no-token-in-output
  console.error(token);
  // ruleid: m365-no-token-in-output
  console.warn("bridge failed", accessToken);
  // ruleid: m365-no-token-in-output
  console.log(`got ${ctx.bearer}`);
  // ruleid: m365-no-token-in-output
  console.error({ detail: ctx.auth.secret });
  // ruleid: m365-no-token-in-output
  console.error("secret is " + clientSecret);
  // ruleid: m365-no-token-in-output
  process.stderr.write(`Authorization: ${authorization}\n`);
  // ruleid: m365-no-token-in-output
  process.stdout.write(token);
  // ruleid: m365-no-token-in-output
  console.error(token.trim());
  // ruleid: m365-no-token-in-output
  console.error({ token });
  // ruleid: m365-no-token-in-output
  log(`captured ${token}`);
  // ruleid: m365-no-token-in-output
  logger.info(accessToken);
  // ruleid: m365-no-token-in-output
  console.error(headers.get("Authorization"));
  // ruleid: m365-no-token-in-output
  console.error(record["token"]);
}

declare function log(line: string): void;
declare const logger: { info(x: unknown): void };
declare const headers: Headers;
declare const record: Record<string, string>;
declare const socket: { write(s: string): void };

function safe(): void {
  // ok: m365-no-token-in-output
  console.error("token unavailable");
  // ok: m365-no-token-in-output
  console.error(`bearer rejected: ${status}`);
  // ok: m365-no-token-in-output
  console.error("Authorization header missing", ctx.count);
  // ok: m365-no-token-in-output
  process.stderr.write(`m365-graph-bridge: ${line}\n`);
  // ok: m365-no-token-in-output
  process.stdout.write("ready\n");
  // ok: m365-no-token-in-output
  log("bridge ready (client pin on)");
  // ok: m365-no-token-in-output
  console.error(headers.get("ConsistencyLevel"));
  // ok: m365-no-token-in-output
  socket.write(JSON.stringify({ ok: true, token }));
  // ok: m365-no-token-in-output
  const stored = token;
  void stored;
}
leaks();
safe();
