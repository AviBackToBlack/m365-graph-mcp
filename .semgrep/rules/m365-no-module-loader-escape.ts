// ruleid: m365-no-module-loader-escape
import { createRequire } from "node:module";
// ruleid: m365-no-module-loader-escape
import * as nodeModule from "module";

declare const name: string;

async function escapes(): Promise<void> {
  // ruleid: m365-no-module-loader-escape
  const net = process.getBuiltinModule("net");
  // ruleid: m365-no-module-loader-escape
  const req = createRequire(import.meta.url);
  // ruleid: m365-no-module-loader-escape
  const req2 = nodeModule.createRequire(import.meta.url);
  // Calls through the returned require are not matched; its createRequire() above is, which fails.
  const cp = req("child_process");
  // ruleid: m365-no-module-loader-escape
  const vm = require("node:vm");
  // ruleid: m365-no-module-loader-escape
  const dyn = await import(name);
  // ruleid: m365-no-module-loader-escape
  const dyn2 = await import(`node:${name}`);
  // ruleid: m365-no-module-loader-escape
  const raw = process.binding("tcp_wrap");

  // ok: m365-no-module-loader-escape
  const fixed = await import("./server.js");
  // ok: m365-no-module-loader-escape
  const text = "require(x) and process.getBuiltinModule are mentioned in a string only";
  void [net, req2, cp, vm, dyn, dyn2, raw, fixed, text];
}
void escapes();
