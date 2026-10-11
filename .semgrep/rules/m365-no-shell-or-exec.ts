import { spawn } from "node:child_process";
// ruleid: m365-no-shell-or-exec
import { exec } from "node:child_process";
// ruleid: m365-no-shell-or-exec
import { execFileSync as run } from "child_process";
import * as nodeCp from "node:child_process";
import processes from "child_process";

declare const cp: typeof import("node:child_process");
declare const child_process: typeof import("node:child_process");
declare const target: string;

function cases(command: string): void {
  // ruleid: m365-no-shell-or-exec
  exec(command);
  // ruleid: m365-no-shell-or-exec
  execSync(command);
  // ruleid: m365-no-shell-or-exec
  cp.execFile(command);
  // ruleid: m365-no-shell-or-exec
  child_process.spawnSync(command);
  // ruleid: m365-no-shell-or-exec
  spawn(command, [], { shell: true });
  // ruleid: m365-no-shell-or-exec
  spawn(command, [], { stdio: "ignore", shell: command });
  // ruleid: m365-no-shell-or-exec
  const options = { shell: "/bin/sh" };
  // ruleid: m365-no-shell-or-exec
  run(command, options); // an aliased import of execFileSync is resolved
  // ruleid: m365-no-shell-or-exec
  nodeCp.exec(command);
  // ruleid: m365-no-shell-or-exec
  processes.execSync(command);

  // ok: m365-no-shell-or-exec
  nodeCp.spawn(command, [], { shell: false });

  // ok: m365-no-shell-or-exec
  spawn(command, [], { shell: false, stdio: ["ignore", "pipe", "pipe", "pipe", "pipe"] });
  // ok: m365-no-shell-or-exec
  spawn(command, [], { stdio: "ignore" });
  // ok: m365-no-shell-or-exec
  const m = /-(\d+)$/.exec(target);
  // ok: m365-no-shell-or-exec
  const n = new RegExp("x").exec(target);
  void m;
  void n;
}
cases("x");
