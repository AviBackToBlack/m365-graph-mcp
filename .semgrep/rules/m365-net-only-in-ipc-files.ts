// ruleid: m365-net-only-in-ipc-files
import net from "node:net";
// ruleid: m365-net-only-in-ipc-files
import { createConnection } from "net";
// ruleid: m365-net-only-in-ipc-files
import * as netNs from "node:net";
// ruleid: m365-net-only-in-ipc-files
const viaRequire = require("node:net");
async function lazy(): Promise<unknown> {
  // ruleid: m365-net-only-in-ipc-files
  return await import("node:net");
}

// ok: m365-net-only-in-ipc-files
import fs from "node:fs";
// ok: m365-net-only-in-ipc-files
import { netStats } from "./netstats.js";
// ok: m365-net-only-in-ipc-files
const text = 'import net from "node:net" as prose in a string';
export { net, createConnection, netNs, viaRequire, lazy, fs, netStats, text };
