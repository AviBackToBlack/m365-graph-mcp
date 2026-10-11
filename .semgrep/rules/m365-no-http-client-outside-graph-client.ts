// ruleid: m365-no-http-client-outside-graph-client
import http from "node:http";
// ruleid: m365-no-http-client-outside-graph-client
import { request } from "https";
// ruleid: m365-no-http-client-outside-graph-client
import * as tls from "node:tls";
// ruleid: m365-no-http-client-outside-graph-client
import axios from "axios";
// ruleid: m365-no-http-client-outside-graph-client
import { lookup } from "node:dns/promises";
// ruleid: m365-no-http-client-outside-graph-client
import "undici";
// ruleid: m365-no-http-client-outside-graph-client
export { Agent } from "node:http2";
// ruleid: m365-no-http-client-outside-graph-client
import {
  get,
  type RequestOptions,
} from "node:https";

// ruleid: m365-no-http-client-outside-graph-client
const viaRequire = require("node:https");
async function other(url: string): Promise<void> {
  // ruleid: m365-no-http-client-outside-graph-client
  await import("node:http");
  // ruleid: m365-no-http-client-outside-graph-client
  await fetch(url);
  // ruleid: m365-no-http-client-outside-graph-client
  await globalThis.fetch(url);
  // ruleid: m365-no-http-client-outside-graph-client
  new WebSocket(url);
}

// ok: m365-no-http-client-outside-graph-client
import net from "node:net";
// ok: m365-no-http-client-outside-graph-client
import { readFile } from "node:fs/promises";
// ok: m365-no-http-client-outside-graph-client
import { token } from "./http-token.js";
// ok: m365-no-http-client-outside-graph-client
import type { GraphClient } from "../graph/client.js";
function aliases(): void {
  // ruleid: m365-no-http-client-outside-graph-client
  const f = fetch;
  // ruleid: m365-no-http-client-outside-graph-client
  const g = globalThis.fetch;
  // ruleid: m365-no-http-client-outside-graph-client
  register(fetch);
  // ruleid: m365-no-http-client-outside-graph-client
  const opts = { impl: fetch };
  // ruleid: m365-no-http-client-outside-graph-client
  new globalThis.WebSocket("wss://example.invalid");
  // ruleid: m365-no-http-client-outside-graph-client
  const thunk = () => fetch;
  // ruleid: m365-no-http-client-outside-graph-client
  const list = [fetch];
  // ruleid: m365-no-http-client-outside-graph-client
  const { fetch: picked } = globalThis;
  void [f, g, opts, thunk, list, picked];
}
function giveFetch(): typeof fetch {
  // ruleid: m365-no-http-client-outside-graph-client
  return fetch;
}
giveFetch();
function destructure(): void {
  // ruleid: m365-no-http-client-outside-graph-client
  const { fetch } = globalThis;
  void fetch;
}
destructure();
function letDouble(): void {
  let fetch = (u: string) => u;
  fetch = (u: string) => u.trim();
  // ok: m365-no-http-client-outside-graph-client
  register({ fetch });
}
letDouble();
declare function register(x: unknown): void;
aliases();
function localDouble(): void {
  const fetch = (u: string) => u;
  // ok: m365-no-http-client-outside-graph-client
  const opts = { fetch };
  void opts;
}
localDouble();
// ok: m365-no-http-client-outside-graph-client
const helper = { fetchImpl: (u: string) => u };
helper.fetchImpl("x");
// ok: m365-no-http-client-outside-graph-client
const text = 'import http from "node:http" mentioned in a string only';
// ok: m365-no-http-client-outside-graph-client
export type Scheme = "https" | "other";
// ok: m365-no-http-client-outside-graph-client
export const SCHEMES = ["http", "https"] as const;
// import http from "node:http" in a comment is documentation, not an import.
export { get, http, request, tls, axios, lookup, viaRequire, other, net, readFile, token, helper, text };
export type { GraphClient };
