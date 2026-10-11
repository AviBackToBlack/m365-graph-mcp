#!/usr/bin/env node
import { BridgeStartError, startBridge } from "./run.js";

// Token bridge daemon entry point (bin: m365-graph-bridge). Diagnostics are fixed categories on
// stderr only; never tokens, headers, claims, tenant ids or browser output.
const log = (line: string): void => {
  process.stderr.write(`m365-graph-bridge: ${line}\n`);
};

// Exit codes: 0 clean/interrupted, 2 config, 3 browser, 4 profile busy, 5 broker, 6 browser exited,
// 7 CDP setup, 8 CDP pipe closed (see ARCHITECTURE.md).
async function main(): Promise<number> {
  // Handlers are installed before async startup so a signal during startup cancels it and cleans up
  // whatever was created. `once`: a second signal falls back to the default (immediate) behavior.
  const abort = new AbortController();
  const onSignal = (): void => abort.abort();
  process.once("SIGINT", onSignal);
  process.once("SIGTERM", onSignal);
  const handle = await startBridge({ log, signal: abort.signal });
  return handle.done;
}

main().then(
  (code) => {
    // Not process.exit(): let stderr flush and the event loop drain naturally.
    process.exitCode = code;
  },
  (error: unknown) => {
    log(
      error instanceof BridgeStartError ? `failed: ${error.category}` : "failed: unexpected error",
    );
    process.exitCode = error instanceof BridgeStartError ? error.exitCode : 1;
  },
);
