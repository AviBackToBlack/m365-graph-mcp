import type { CdpClient, CdpEvent } from "./cdp-pipe.js";
import type { TokenSnapshot, TokenSource } from "./broker.js";
import {
  MIN_REMAINING_MS,
  classifyGraphOrigin,
  evaluateCandidate,
  isExactGraphOrigin,
} from "./token-gate.js";
import { DISABLED_ACCEPTANCE_DIAGNOSTICS } from "./acceptance-diagnostics.js";
import type {
  AcceptanceDiagnostics,
  AuthorizationObservation,
  GraphExplorerPageShape,
} from "./acceptance-diagnostics.js";
import type { RejectReason, TokenGateConfig } from "./token-gate.js";

/**
 * Observes the bridge-owned Graph Explorer page over CDP and turns outbound Graph requests into an
 * accepted bearer (a TokenSource for the broker). It only observes: browser sign-in, MFA and token
 * renewal stay browser-owned.
 *
 * RENEWAL LIMITATION: a renewed token is captured only when Graph Explorer itself issues a Graph
 * request (a user action, or Graph Explorer's own calls). The bridge never triggers requests. When
 * the held token drops below the 5-minute serve floor the broker reports `reauth_required` until a
 * fresh acceptable token is observed.
 *
 * Sensitive data: events are transient. Only the Bearer value is extracted from an eligible
 * request; headers, events and the bearer are never logged, persisted, or put into errors or
 * diagnostics (diagnostics are fixed categories).
 */

/**
 * Graph Explorer page paths (query/fragment allowed): `/graph/graph-explorer` or
 * `/<locale>/graph/graph-explorer`, optional trailing slash. A locale is exactly two ASCII letters,
 * a hyphen and two ASCII letters (case-insensitive; e.g. `en-us`, `de-DE`). Nothing else is
 * accepted: no extra segments, no suffixes. Host, scheme, port and credentials are checked
 * separately and stay exact.
 */
const GRAPH_EXPLORER_PATH_RE = /^(?:\/([a-z]{2}-[a-z]{2}))?\/graph\/graph-explorer\/?$/i;

/** URL opened in the bridge-owned browser (Microsoft Learn's documented Graph Explorer link). */
export const GRAPH_EXPLORER_URL = "https://developer.microsoft.com/graph/graph-explorer/";

/**
 * The single Graph Explorer page matcher. Returns the page shape (`canonical`, or `localized` when
 * the path carries an `xx-xx` locale segment) or undefined when the URL is not an eligible page.
 * The locale itself is deliberately not returned. `isGraphExplorerUrl` is derived from this, and the
 * D48 acceptance diagnostics print only the shape.
 */
export function classifyGraphExplorerPage(raw: unknown): GraphExplorerPageShape | undefined {
  if (typeof raw !== "string") return undefined;
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    return undefined;
  }
  if (
    url.protocol !== "https:" ||
    url.hostname !== "developer.microsoft.com" ||
    url.port !== "" ||
    url.username !== "" ||
    url.password !== ""
  ) {
    return undefined;
  }
  const match = GRAPH_EXPLORER_PATH_RE.exec(url.pathname);
  if (match === null) return undefined;
  return match[1] === undefined ? "canonical" : "localized";
}

export function isGraphExplorerUrl(raw: unknown): boolean {
  return classifyGraphExplorerPage(raw) !== undefined;
}

export const DEFAULT_PENDING_TTL_MS = 15_000;
/** Delay before rescanning after a failed (non-stale) attach; bounds retry frequency. */
export const DEFAULT_ATTACH_RETRY_MS = 2_000;
export const MAX_PENDING_REQUESTS = 64;

export type CaptureDiagnostic =
  | { kind: "target_attached" }
  | { kind: "target_lost" }
  | { kind: "attach_failed" }
  | { kind: "token_accepted"; remainingSeconds: number }
  | { kind: "token_rejected"; reason: RejectReason }
  | { kind: "token_cleared"; reason: "http_401" | "lifetime" };

export interface GraphExplorerCaptureOptions {
  cdp: CdpClient;
  gate: TokenGateConfig;
  now?: () => number;
  pendingTtlMs?: number;
  /** Delay before retrying after a failed attach (default 2 s). */
  attachRetryMs?: number;
  maxPending?: number;
  /** Fixed-category diagnostics only (no token, header or claim content). */
  onDiagnostic?: (d: CaptureDiagnostic) => void;
  /** D48 shape-only acceptance diagnostics; absent means disabled (no observation work). */
  acceptanceDiagnostics?: AcceptanceDiagnostics;
}

interface PendingRequest {
  url: string;
  bearer: string | undefined;
  status: number | undefined;
  createdMs: number;
  /**
   * True when this request id reached its Graph URL through a redirect hop. ExtraInfo carries no
   * URL or hop marker, so it cannot be tied to the final hop: for such requests only the bearer in
   * the final hop's own requestWillBeSent headers counts.
   */
  redirected: boolean;
  /** D48 observation only: where the bearer was first seen. Never read by capture logic. */
  authObservation?: AuthorizationObservation;
  /** D48 observation only: a non-redirected ExtraInfo event carried a bearer. Never read by capture. */
  extraInfoBearer?: boolean;
}

/** An ExtraInfo bearer candidate seen before its requestWillBeSent (only the bearer is kept). */
interface BufferedExtra {
  bearer: string;
  createdMs: number;
}

interface AttachAttempt {
  targetId: string;
  /** Latest URL seen for the target while attaching; undefined once it is destroyed. */
  url: string | undefined;
}

const isRecord = (v: unknown): v is Record<string, unknown> =>
  v !== null && typeof v === "object" && !Array.isArray(v);

const BEARER_RE = /^Bearer ([^\s]+)$/i;

/**
 * Returns the bearer value iff exactly one header is named Authorization (case-insensitive) and
 * its value is a single `Bearer <token>`. Anything else (none, several, other scheme) yields
 * undefined. No other header is read.
 */
export function extractBearer(headers: unknown): string | undefined {
  if (!isRecord(headers)) return undefined;
  let value: unknown;
  let count = 0;
  for (const key of Object.keys(headers)) {
    if (key.toLowerCase() === "authorization") {
      count++;
      value = headers[key];
    }
  }
  if (count !== 1 || typeof value !== "string") return undefined;
  return BEARER_RE.exec(value)?.[1];
}

export class GraphExplorerCapture implements TokenSource {
  readonly clientPinEnabled: boolean;
  private readonly cdp: CdpClient;
  private readonly gate: TokenGateConfig;
  private readonly now: () => number;
  private readonly ttlMs: number;
  private readonly maxPending: number;
  private readonly diag: (d: CaptureDiagnostic) => void;
  private readonly acc: AcceptanceDiagnostics;
  private unsubscribe: (() => void) | undefined;

  private targetId: string | undefined;
  private targetUrl: string | undefined;
  private attaching: AttachAttempt | undefined;
  private sessionId: string | undefined;
  private stopped = false;
  private scanning = false;
  private rescanAgain = false;
  private retryTimer: ReturnType<typeof setTimeout> | undefined;
  private readonly attachRetryMs: number;
  private readonly requests = new Map<string, PendingRequest>();
  private readonly extras = new Map<string, BufferedExtra>();

  private held: { token: string; expMs: number } | undefined;
  private failure: TokenSnapshot = { state: "unavailable" };

  constructor(options: GraphExplorerCaptureOptions) {
    this.cdp = options.cdp;
    this.gate = options.gate;
    this.clientPinEnabled = options.gate.clientPin !== undefined;
    this.now = options.now ?? Date.now;
    this.ttlMs = options.pendingTtlMs ?? DEFAULT_PENDING_TTL_MS;
    this.maxPending = options.maxPending ?? MAX_PENDING_REQUESTS;
    this.diag = options.onDiagnostic ?? (() => undefined);
    this.attachRetryMs = options.attachRetryMs ?? DEFAULT_ATTACH_RETRY_MS;
    this.acc = options.acceptanceDiagnostics ?? DISABLED_ACCEPTANCE_DIAGNOSTICS;
  }

  async start(): Promise<void> {
    this.unsubscribe = this.cdp.onEvent((e) => this.onEvent(e));
    await this.cdp.send("Target.setDiscoverTargets", { discover: true });
    await this.scanTargets();
  }

  private async scanTargets(): Promise<void> {
    const res = await this.cdp.send("Target.getTargets");
    const infos = isRecord(res) && Array.isArray(res.targetInfos) ? res.targetInfos : [];
    for (const info of infos) this.considerTarget(info);
  }

  /** Looks for another eligible page after the claim was released or lost. Errors are swallowed. */
  private async rescan(): Promise<void> {
    if (this.stopped) return;
    if (this.scanning) {
      this.rescanAgain = true;
      return;
    }
    this.scanning = true;
    try {
      do {
        this.rescanAgain = false;
        await this.scanTargets();
      } while (this.rescanAgain && !this.stopped);
    } catch {
      // The command failed or timed out. An idle signed-in page may emit no further target event,
      // so schedule the same bounded retry as a failed attach (a closed pipe ends the bridge).
      this.scheduleRetry();
    } finally {
      this.scanning = false;
    }
  }

  /** Stops observing and drops all held material (browser exit or shutdown). */
  stop(): void {
    this.stopped = true;
    if (this.retryTimer !== undefined) clearTimeout(this.retryTimer);
    this.retryTimer = undefined;
    this.unsubscribe?.();
    this.unsubscribe = undefined;
    this.markUnavailable();
  }

  markUnavailable(): void {
    this.held = undefined;
    this.requests.clear();
    this.extras.clear();
    this.recentlyEvaluated.clear();
    this.sessionId = undefined;
    this.targetId = undefined;
    this.targetUrl = undefined;
    this.failure = { state: "unavailable" };
  }

  current(): TokenSnapshot {
    this.sweep();
    const held = this.held;
    if (held !== undefined) {
      if (held.expMs - this.now() >= MIN_REMAINING_MS) {
        return { state: "ok", token: held.token, expMs: held.expMs };
      }
      this.held = undefined;
      this.diag({ kind: "token_cleared", reason: "lifetime" });
      this.failure = { state: "reauth_required", reason: "lifetime" };
    }
    return this.failure;
  }

  // ---- target handling ----

  private onEvent(e: CdpEvent): void {
    switch (e.method) {
      case "Target.targetCreated":
      case "Target.targetInfoChanged":
        this.considerTarget(e.params.targetInfo);
        return;
      case "Target.targetDestroyed":
        if (this.attaching !== undefined && e.params.targetId === this.attaching.targetId) {
          this.attaching.url = undefined;
        }
        if (this.targetId !== undefined && e.params.targetId === this.targetId) this.targetLost();
        return;
      case "Target.detachedFromTarget":
        if (
          e.sessionId === undefined &&
          this.sessionId !== undefined &&
          e.params.sessionId === this.sessionId
        ) {
          this.targetLost();
        }
        return;
      case "Target.attachedToTarget":
        return; // the attach command result is authoritative
      default:
        break;
    }
    // Network events: only from the single session attached to the bridge's Graph Explorer page.
    if (this.sessionId === undefined || e.sessionId !== this.sessionId) return;
    this.sweep();
    switch (e.method) {
      case "Network.requestWillBeSent":
        this.onRequest(e.params, true);
        return;
      case "Network.requestWillBeSentExtraInfo":
        this.onRequest(e.params, false);
        return;
      case "Network.responseReceived":
        this.onResponse(e.params);
        return;
      case "Network.loadingFailed":
        if (typeof e.params.requestId === "string") {
          this.requests.delete(e.params.requestId);
          this.extras.delete(e.params.requestId);
        }
        return;
    }
  }

  private considerTarget(info: unknown): void {
    if (this.stopped || !isRecord(info) || typeof info.targetId !== "string") return;
    const url = typeof info.url === "string" ? info.url : undefined;
    if (info.targetId === this.targetId) {
      if (url !== undefined) {
        // The claimed page left Graph Explorer: release the claim so another page can be used.
        if (!isGraphExplorerUrl(url)) this.release();
        else {
          if (this.acc.enabled) this.observePageShape(url, this.targetUrl);
          this.targetUrl = url;
        }
      }
      return;
    }
    if (this.attaching !== undefined) {
      if (info.targetId === this.attaching.targetId && url !== undefined) this.attaching.url = url;
      return;
    }
    if (this.targetId !== undefined) return;
    if (info.type !== "page" || !isGraphExplorerUrl(url)) return;
    void this.attach(info.targetId, url as string);
  }

  /** D48: reports the claimed page's shape (never its URL or locale) when it is new or changed. */
  private observePageShape(url: string | undefined, previousUrl: string | undefined): void {
    const shape = classifyGraphExplorerPage(url);
    if (shape !== undefined && shape !== classifyGraphExplorerPage(previousUrl)) {
      this.acc.graphExplorerPage(shape);
    }
  }

  private detach(sessionId: string): void {
    void this.cdp.send("Target.detachFromTarget", { sessionId }).catch(() => undefined);
  }

  private async attach(targetId: string, url: string): Promise<void> {
    const attempt: AttachAttempt = { targetId, url };
    this.attaching = attempt;
    let sessionId: string | undefined;
    let stale = false;
    try {
      const res = await this.cdp.send("Target.attachToTarget", { targetId, flatten: true });
      sessionId = isRecord(res) && typeof res.sessionId === "string" ? res.sessionId : undefined;
      if (sessionId === undefined) throw new Error("no session");
      await this.cdp.send("Network.enable", undefined, sessionId);
      if (this.stopped || !isGraphExplorerUrl(attempt.url)) {
        // The page navigated away (or went away) while we were attaching: do not claim it.
        stale = true;
        this.detach(sessionId);
        this.diag({ kind: "attach_failed" });
        return;
      }
      this.targetId = targetId;
      this.targetUrl = attempt.url;
      this.sessionId = sessionId;
      if (this.acc.enabled) this.observePageShape(attempt.url, undefined);
      if (this.held === undefined && this.failure.state === "unavailable") {
        this.failure = { state: "reauth_required" };
      }
      this.diag({ kind: "target_attached" });
    } catch {
      this.diag({ kind: "attach_failed" });
      if (sessionId !== undefined) this.detach(sessionId);
    } finally {
      this.attaching = undefined;
      // Without a committed claim, look again: events for other targets were ignored during the
      // attach window, and an idle signed-in page may never emit another event. A stale page is
      // rescanned at once; a failed attach is retried after a bounded delay (no tight loop).
      if (!this.stopped && this.targetId === undefined) {
        if (stale) void this.rescan();
        else this.scheduleRetry();
      }
    }
  }

  private scheduleRetry(): void {
    if (this.retryTimer !== undefined || this.stopped) return;
    this.retryTimer = setTimeout(() => {
      this.retryTimer = undefined;
      if (!this.stopped && this.targetId === undefined && this.attaching === undefined) {
        void this.rescan();
      }
    }, this.attachRetryMs);
    this.retryTimer.unref?.();
  }

  /** Releases the claim without the target going away (navigated off Graph Explorer). */
  private release(): void {
    const sessionId = this.sessionId;
    this.targetLost(false);
    if (sessionId !== undefined) this.detach(sessionId);
    void this.rescan();
  }

  private targetLost(rescan = true): void {
    this.targetId = undefined;
    this.targetUrl = undefined;
    this.sessionId = undefined;
    this.requests.clear();
    this.extras.clear();
    this.recentlyEvaluated.clear();
    this.diag({ kind: "target_lost" });
    if (rescan) void this.rescan();
  }

  // ---- request correlation ----

  private sweep(): void {
    const cutoff = this.now() - this.ttlMs;
    for (const [id, r] of this.requests) {
      if (r.createdMs < cutoff) this.requests.delete(id);
    }
    for (const [id, x] of this.extras) {
      if (x.createdMs < cutoff) this.extras.delete(id);
    }
  }

  private track(requestId: string, url: string): PendingRequest {
    let r = this.requests.get(requestId);
    if (r === undefined) {
      while (this.requests.size >= this.maxPending) {
        const oldest = this.requests.keys().next();
        if (oldest.done) break;
        this.requests.delete(oldest.value);
      }
      r = { url, bearer: undefined, status: undefined, createdMs: this.now(), redirected: false };
      this.requests.set(requestId, r);
    }
    return r;
  }

  private bufferExtra(requestId: string, bearer: string): void {
    this.extras.delete(requestId);
    while (this.extras.size >= this.maxPending) {
      const oldest = this.extras.keys().next();
      if (oldest.done) break;
      this.extras.delete(oldest.value);
    }
    this.extras.set(requestId, { bearer, createdMs: this.now() });
  }

  private onRequest(params: Record<string, unknown>, withUrl: boolean): void {
    const requestId = params.requestId;
    if (typeof requestId !== "string") return;
    let r: PendingRequest | undefined;
    if (withUrl) {
      const request = params.request;
      const url = isRecord(request) ? request.url : undefined;
      const buffered = this.extras.get(requestId);
      const wasTracked = this.requests.has(requestId);
      this.extras.delete(requestId);
      this.requests.delete(requestId); // a redirect hop reuses the id; always start fresh
      const redirected = isRecord(params.redirectResponse);
      if (typeof url !== "string" || !isExactGraphOrigin(url)) {
        // D48 observation only: a tracked Graph request redirected away from the Graph origin.
        if (redirected && wasTracked && this.acc.enabled) {
          this.acc.redirectSeen();
          if (buffered !== undefined) this.acc.redirectExtraInfoDiscarded();
        }
        return;
      }
      r = this.track(requestId, url);
      r.redirected = redirected;
      r.bearer = extractBearer(isRecord(request) ? request.headers : undefined);
      if (this.acc.enabled) {
        // D48 observation only. The label records where a bearer was FIRST seen (arrival order),
        // independently of which copy capture uses below.
        if (!redirected && buffered !== undefined) {
          r.extraInfoBearer = true;
          r.authObservation = "extra_info_before_main";
        } else if (r.bearer !== undefined) {
          r.authObservation = "main_event";
        }
        if (redirected) {
          this.acc.redirectSeen();
          if (buffered !== undefined) this.acc.redirectExtraInfoDiscarded();
        }
      }
      // A buffered ExtraInfo belongs to this request only when no redirect is involved: after a
      // redirect it may belong to an earlier hop (possibly to a non-Graph origin), so it is dropped.
      if (r.bearer === undefined && !redirected) r.bearer = buffered?.bearer;
    } else {
      const bearer = extractBearer(params.headers);
      if (bearer === undefined) return;
      r = this.requests.get(requestId);
      if (r === undefined) {
        // D48 observation only: ExtraInfo for a request already evaluated (its response came
        // first). Reported once; capture handling below is unchanged.
        if (this.acc.enabled && this.takeRecentlyEvaluated(requestId)) {
          this.acc.extraInfoAuthorization("present_after_response");
        }
        // ExtraInfo may precede requestWillBeSent: keep only the bearer candidate (bounded, TTL);
        // the exact-Graph-origin check on the main event still gates it.
        this.bufferExtra(requestId, bearer);
        return;
      }
      if (r.redirected) {
        if (this.acc.enabled) this.acc.redirectExtraInfoDiscarded();
        return; // cannot be tied to the final hop
      }
      if (this.acc.enabled) r.extraInfoBearer = true;
      if (r.bearer === undefined) {
        r.bearer = bearer;
        if (this.acc.enabled) r.authObservation = "extra_info_after_main";
      }
    }
    this.maybeEvaluate(requestId, r);
  }

  private onResponse(params: Record<string, unknown>): void {
    const requestId = params.requestId;
    if (typeof requestId !== "string") return;
    const r = this.requests.get(requestId);
    if (r === undefined) return;
    const response = params.response;
    const status = isRecord(response) ? response.status : undefined;
    const url = isRecord(response) ? response.url : undefined;
    if (
      typeof status !== "number" ||
      !Number.isInteger(status) ||
      !isExactGraphOrigin(String(url))
    ) {
      this.requests.delete(requestId);
      return;
    }
    r.status = status;
    this.maybeEvaluate(requestId, r);
  }

  private maybeEvaluate(requestId: string, r: PendingRequest): void {
    if (r.bearer === undefined || r.status === undefined) return;
    this.requests.delete(requestId);
    if (this.acc.enabled) {
      // D48 observation only: categories derived from the already-matched request, never its URL.
      this.acc.graphRequestOrigin(classifyGraphOrigin(r.url));
      if (r.authObservation !== undefined) this.acc.authorizationObservation(r.authObservation);
      this.acc.extraInfoAuthorization(r.extraInfoBearer === true ? "present" : "not_observed");
      // A response was received for this hop, so it is the final one of any redirect chain.
      if (r.redirected && r.authObservation === "main_event") {
        this.acc.redirectFinalHopAuthorizationFromMainEvent();
      }
      if (!r.redirected) this.rememberEvaluated(requestId);
    }
    this.accept(r.bearer, r.url, r.status);
  }

  /** D48 observation only: ids of recently evaluated, non-redirected requests (bounded, TTL). */
  private readonly recentlyEvaluated = new Map<string, number>();

  private rememberEvaluated(requestId: string): void {
    this.recentlyEvaluated.delete(requestId);
    while (this.recentlyEvaluated.size >= this.maxPending) {
      const oldest = this.recentlyEvaluated.keys().next();
      if (oldest.done) break;
      this.recentlyEvaluated.delete(oldest.value);
    }
    this.recentlyEvaluated.set(requestId, this.now());
  }

  /** Removes and reports whether `requestId` was evaluated within the pending TTL. */
  private takeRecentlyEvaluated(requestId: string): boolean {
    const at = this.recentlyEvaluated.get(requestId);
    if (at === undefined) return false;
    this.recentlyEvaluated.delete(requestId);
    return this.now() - at <= this.ttlMs;
  }

  // ---- gate integration ----

  private accept(bearer: string, url: string, status: number): void {
    const nowMs = this.now();
    if (status === 401 && this.held?.token === bearer) {
      // Graph rejected the very token we hold: stop serving it.
      this.held = undefined;
      this.failure = { state: "reauth_required", reason: "http_401" };
      this.diag({ kind: "token_cleared", reason: "http_401" });
      return;
    }
    const sourceOk =
      this.sessionId !== undefined &&
      this.targetUrl !== undefined &&
      isGraphExplorerUrl(this.targetUrl);
    const result = evaluateCandidate(
      { bearer, requestUrl: url, responseStatus: status, sourceIsBridgeTarget: sourceOk },
      this.gate,
      nowMs,
      this.acc,
    );
    if (result.ok) {
      const held = this.held;
      const heldUsable = held !== undefined && held.expMs - nowMs >= MIN_REMAINING_MS;
      if (
        !heldUsable ||
        (result.bearer !== held.token && result.expMs >= held.expMs) // equal expiry may be a renewal
      ) {
        this.held = { token: result.bearer, expMs: result.expMs };
      }
      this.diag({
        kind: "token_accepted",
        remainingSeconds: Math.floor(result.remainingMs / 1000),
      });
      return;
    }
    this.diag({ kind: "token_rejected", reason: result.reason });
    const held = this.held;
    if (held !== undefined && held.expMs - nowMs >= MIN_REMAINING_MS) return; // keep the good one
    this.held = undefined;
    this.failure =
      result.reason === "unsupported_token_format"
        ? { state: "unsupported_token_format", reason: result.reason }
        : { state: "reauth_required", reason: result.reason };
  }
}
