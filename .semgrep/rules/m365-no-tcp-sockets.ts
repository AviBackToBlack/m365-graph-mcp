import net from "node:net";
import { createConnection, connect } from "node:net";

declare const socketPath: string;
declare const port: number;
declare const options: { socketPath: string };

function sockets(): void {
  // ruleid: m365-no-tcp-sockets
  net.createConnection({ port: 8080 });
  // ruleid: m365-no-tcp-sockets
  net.createConnection({ host: "127.0.0.1", port });
  // ruleid: m365-no-tcp-sockets
  net.connect({ port, host: "localhost" });
  // ruleid: m365-no-tcp-sockets
  net.createConnection(8080, "127.0.0.1");
  // ruleid: m365-no-tcp-sockets
  net.connect(9222);
  // ruleid: m365-no-tcp-sockets
  createConnection(8080, "example.com");
  // ruleid: m365-no-tcp-sockets
  connect({ port: 443, host: "example.com" });
  // ruleid: m365-no-tcp-sockets
  new net.Socket().connect(443, "example.com");
  const sock = new net.Socket();
  // ruleid: m365-no-tcp-sockets
  sock.connect({ port: 443 });

  const server = net.createServer();
  // ruleid: m365-no-tcp-sockets
  server.listen(8080);
  // ruleid: m365-no-tcp-sockets
  server.listen(0, "127.0.0.1");
  // ruleid: m365-no-tcp-sockets
  server.listen({ port, host: "::1" });
  // ruleid: m365-no-tcp-sockets
  server.listen({ host: "0.0.0.0" });
  // ruleid: m365-no-tcp-sockets
  server.listen(port, "0.0.0.0", () => undefined);
  // ruleid: m365-no-tcp-sockets
  server.listen(port);
  // ruleid: m365-no-tcp-sockets
  server.listen();

  // ok: m365-no-tcp-sockets
  net.createConnection({ path: socketPath });
  // ok: m365-no-tcp-sockets
  net.createConnection({ path: socketPath, timeout: 1000 });
  // ok: m365-no-tcp-sockets
  createConnection({ path: socketPath });
  // ok: m365-no-tcp-sockets
  sock.connect({ path: socketPath });
  // ok: m365-no-tcp-sockets
  server.listen(socketPath, () => undefined);
  // ok: m365-no-tcp-sockets
  server.listen(options.socketPath);
  // ok: m365-no-tcp-sockets
  server.listen({ path: socketPath });
}
sockets();
