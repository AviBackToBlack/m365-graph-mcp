import fs from "node:fs";
import fsp from "node:fs/promises";
import { readFile } from "node:fs/promises";
// ruleid: m365-no-file-writes-outside-bridge
import { writeFile, appendFileSync as aliasedAppend } from "node:fs/promises";
// ruleid: m365-no-file-writes-outside-bridge
import { writeSync, openSync } from "node:fs";

declare const target: string;
declare const data: string;
declare const fd: number;
declare const stream: { write(s: string): void };

async function cases(): Promise<void> {
  // ruleid: m365-no-file-writes-outside-bridge
  await writeFile(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.writeFileSync(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.appendFileSync(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.createWriteStream(target);
  // ruleid: m365-no-file-writes-outside-bridge
  await fs.promises.writeFile(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.copyFileSync(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  aliasedAppend(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.writeSync(fd, data);
  // ruleid: m365-no-file-writes-outside-bridge
  writeSync(fd, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.openSync(target, "w");
  // Generic name: the bare call is not matched, but its fs import (above) is, which fails the gate.
  openSync(target, "w");
  // ruleid: m365-no-file-writes-outside-bridge
  await fsp.open(target, "w");
  // ruleid: m365-no-file-writes-outside-bridge
  fs.rmSync(target, { recursive: true });
  // ruleid: m365-no-file-writes-outside-bridge
  fs.unlinkSync(target);
  // ruleid: m365-no-file-writes-outside-bridge
  await fsp.rename(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.mkdirSync(target, { recursive: true });
  // ruleid: m365-no-file-writes-outside-bridge
  fs.mkdtempSync(target);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.truncateSync(target);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.cpSync(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  await fsp.cp(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.symlinkSync(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.linkSync(target, data);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.chmodSync(target, 0o777);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.chownSync(target, 0, 0);
  // ruleid: m365-no-file-writes-outside-bridge
  fs.utimesSync(target, 0, 0);

  // ok: m365-no-file-writes-outside-bridge
  await readFile(target, "utf8");
  // ok: m365-no-file-writes-outside-bridge
  fs.readFileSync(target, "utf8");
  // ok: m365-no-file-writes-outside-bridge
  fs.existsSync(target);
  // ok: m365-no-file-writes-outside-bridge
  fs.lstatSync(target);
  // ok: m365-no-file-writes-outside-bridge
  process.stdout.write(data);
  // ok: m365-no-file-writes-outside-bridge
  stream.write(data);
}
void cases();
