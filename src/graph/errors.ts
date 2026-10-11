import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { errorResult, type ErrorCode } from "../envelope.js";

/**
 * Structured GraphClient failure. `message` and `details` are built only from fixed
 * text and bounded, sanitized values; they never contain tokens, Authorization
 * headers, or raw response bodies. No `cause` is ever attached.
 */
export class GraphClientError extends Error {
  readonly code: ErrorCode;
  readonly details: Record<string, unknown> | undefined;

  constructor(code: ErrorCode, message: string, details?: Record<string, unknown>) {
    super(message);
    this.name = "GraphClientError";
    this.code = code;
    this.details = details;
  }

  toJSON(): { name: string; code: ErrorCode; message: string; details?: Record<string, unknown> } {
    return this.details === undefined
      ? { name: this.name, code: this.code, message: this.message }
      : { name: this.name, code: this.code, message: this.message, details: this.details };
  }

  toResult(): CallToolResult {
    return errorResult(this.code, this.message, this.details);
  }
}

const MAX_TEXT = 200;

/**
 * Strips control characters, redacts bearer-looking material and JWT-shaped strings,
 * and truncates. Applied to every piece of Graph-supplied text before it is surfaced.
 */
export function sanitizeGraphText(value: unknown, max = MAX_TEXT): string | undefined {
  if (typeof value !== "string") return undefined;
  const cleaned = value
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u001f\u007f]+/g, " ")
    .replace(/Bearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [redacted]")
    .replace(/eyJ[A-Za-z0-9_-]{5,}(?:\.[A-Za-z0-9_-]+){0,2}/g, "[redacted]")
    .trim();
  if (cleaned.length === 0) return undefined;
  return cleaned.length > max ? `${cleaned.slice(0, max)}...` : cleaned;
}

/** Removes any occurrence of the live token from text (defense in depth). */
export function redactToken(text: string, token: string | undefined): string {
  if (token === undefined || token.length === 0) return text;
  return text.split(token).join("[redacted]");
}
