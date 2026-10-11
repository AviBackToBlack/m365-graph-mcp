// ruleid: m365-no-child-process-outside-browser
import { spawn } from "node:child_process";
// ruleid: m365-no-child-process-outside-browser
import cp from "child_process";
// ruleid: m365-no-child-process-outside-browser
import * as cp2 from "node:child_process";
// ruleid: m365-no-child-process-outside-browser
import type { ChildProcess } from "node:child_process";
// ruleid: m365-no-child-process-outside-browser
import "node:child_process";

// ruleid: m365-no-child-process-outside-browser
const viaRequire = require("node:child_process");
async function viaDynamicImport() {
  // ruleid: m365-no-child-process-outside-browser
  return await import("node:child_process");
}

// ok: m365-no-child-process-outside-browser
import { readFile } from "node:fs/promises";
// ok: m365-no-child-process-outside-browser
import net from "node:net";
// ok: m365-no-child-process-outside-browser
const note = "child_process is only mentioned in a string";
