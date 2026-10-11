import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";

/** Stable, machine-readable error codes shared by every tool. Do not rename. */
export const ERROR_CODES = [
  "not_implemented",
  "invalid_input",
  "reauth_required",
  "capability_gap",
  "rate_limited",
  "graph_error",
  "restart_required",
] as const;

export type ErrorCode = (typeof ERROR_CODES)[number];

export interface SuccessEnvelope {
  ok: true;
  data: unknown;
  meta?: Record<string, unknown>;
}

export interface ErrorEnvelope {
  ok: false;
  error: {
    code: ErrorCode;
    message: string;
    details?: Record<string, unknown>;
  };
}

export type Envelope = SuccessEnvelope | ErrorEnvelope;

function toResult(envelope: Envelope, isError: boolean): CallToolResult {
  return {
    content: [{ type: "text", text: JSON.stringify(envelope) }],
    structuredContent: envelope as unknown as Record<string, unknown>,
    isError,
  };
}

export function successResult(data: unknown, meta?: Record<string, unknown>): CallToolResult {
  const envelope: SuccessEnvelope =
    meta === undefined ? { ok: true, data } : { ok: true, data, meta };
  return toResult(envelope, false);
}

/**
 * Build a structured tool error. `message` and `details` must never contain
 * access tokens or Authorization headers; callers pass only sanitized text.
 */
export function errorResult(
  code: ErrorCode,
  message: string,
  details?: Record<string, unknown>,
): CallToolResult {
  const envelope: ErrorEnvelope = {
    ok: false,
    error: details === undefined ? { code, message } : { code, message, details },
  };
  return toResult(envelope, true);
}
