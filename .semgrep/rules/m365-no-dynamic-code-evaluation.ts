// ruleid: m365-no-dynamic-code-evaluation
import vm from "node:vm";
// ruleid: m365-no-dynamic-code-evaluation
import { runInNewContext } from "vm";
// ruleid: m365-no-dynamic-code-evaluation
import { Worker } from "node:worker_threads";
// ruleid: m365-no-dynamic-code-evaluation
import cluster from "cluster";
// ruleid: m365-no-dynamic-code-evaluation
import * as repl from "node:repl";

declare const code: string;
declare const lookup: { evaluate(expression: string): unknown };

function cases(): void {
  // ruleid: m365-no-dynamic-code-evaluation
  eval(code);
  // ruleid: m365-no-dynamic-code-evaluation
  globalThis.eval(code);
  // ruleid: m365-no-dynamic-code-evaluation
  process.dlopen({ exports: {} }, code);
  // ruleid: m365-no-dynamic-code-evaluation
  const w = require("node:worker_threads");
  void [Worker, cluster, repl, w];
  // ruleid: m365-no-dynamic-code-evaluation
  const a = new Function("a", code);
  // ruleid: m365-no-dynamic-code-evaluation
  const b = Function(code);
  // ruleid: m365-no-dynamic-code-evaluation
  const c = require("node:vm");

  // ok: m365-no-dynamic-code-evaluation
  lookup.evaluate(code);
  // ok: m365-no-dynamic-code-evaluation
  const d = JSON.parse(code);
  // ok: m365-no-dynamic-code-evaluation
  const e = new RegExp(code);
  // ok: m365-no-dynamic-code-evaluation
  const f = "eval(code) as prose in a string";
  void [a, b, c, d, e, f, vm, runInNewContext];
}
cases();
