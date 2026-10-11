declare const port: number;

// ruleid: m365-no-remote-debugging-port
const a = "--remote-debugging-port=9222";
// ruleid: m365-no-remote-debugging-port
const b = '--remote-debugging-port=0';
// ruleid: m365-no-remote-debugging-port
const c = `--remote-debugging-port=${port}`;
// ruleid: m365-no-remote-debugging-port
const d = ["--user-data-dir=/x", "-remote-debugging-port=0"];
// ruleid: m365-no-remote-debugging-port
const e = "--remote-debugging-address=0.0.0.0";

// ok: m365-no-remote-debugging-port
const f = "--remote-debugging-pipe";
// ok: m365-no-remote-debugging-port
const g = ["--user-data-dir=/x", "--remote-debugging-pipe", "--no-first-run"];
// Never `--remote-debugging-port`: a comment that merely documents the rule is not a finding.
// ok: m365-no-remote-debugging-port
const h = "--no-first-run";
export { a, b, c, d, e, f, g, h };
