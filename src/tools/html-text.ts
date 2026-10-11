/**
 * Dependency-free HTML to plain-text normalizer for Teams message bodies (SM-11 part B).
 *
 * The input is UNTRUSTED data and is never rendered, executed or followed: it is scanned once,
 * left to right, and turned into text. Properties the callers and tests rely on:
 *
 * - Linear time. Every loop step advances the cursor; every `indexOf` search starts at the cursor
 *   and either consumes what it found or (on failure) is never repeated (an unterminated tag
 *   turns the remaining `<` into literal text; an unterminated comment, `<script>` or `<style>`
 *   drops the rest of the input). There is no regex over the remainder and no backtracking.
 * - Bounded output. At most `maxChars` characters are produced (never splitting a surrogate pair)
 *   and the input is capped at `MAX_HTML_INPUT_CHARS`; `truncated` says either cap was hit.
 * - Never throws on malformed markup. Unknown tags are dropped and their text kept.
 * - Entities are decoded exactly once, on text and attribute values only. Numeric references to
 *   NUL, surrogates or code points above U+10FFFF become U+FFFD; unknown named entities stay
 *   literal. C0 and C1 control characters and bidi override characters are removed from the
 *   result.
 * - Nothing is fetched. `<img>` becomes a placeholder; a hosted-content id is surfaced only for a
 *   Graph URL of the SAME chat and message (so a later unit can fetch it deliberately). Link
 *   targets are printed as text only for http(s) and mailto URLs.
 *
 * Known simplification: a tag ends at the first `>` (a `>` inside a quoted attribute value ends the
 * tag early and the remainder is kept as text). Teams escapes `>` in attribute values.
 */

/** Longest HTML input processed; the rest is ignored and reported as truncated. */
export const MAX_HTML_INPUT_CHARS = 500_000;
/** Appended (by the caller) after text that hit a cap. */
export const TRUNCATION_MARKER = "\n[... truncated]";

export interface HtmlTextResult {
  text: string;
  truncated: boolean;
}

export interface HtmlTextContext {
  /**
   * The message's scope and id: a hosted-content `<img>` must point at exactly these. A chat
   * message sets `chatId`; a channel message sets `teamId` and `channelId` (and `replyToId` when it
   * is a reply). Never both: a context with both or neither scope never yields an id.
   */
  chatId?: string;
  teamId?: string;
  channelId?: string;
  messageId?: string;
  replyToId?: string;
}

const NAMED_ENTITIES: ReadonlyMap<string, string> = new Map([
  ["amp", "&"],
  ["lt", "<"],
  ["gt", ">"],
  ["quot", '"'],
  ["apos", "'"],
  ["nbsp", " "],
  ["ndash", "–"],
  ["mdash", "—"],
  ["hellip", "…"],
  ["lsquo", "‘"],
  ["rsquo", "’"],
  ["ldquo", "“"],
  ["rdquo", "”"],
  ["laquo", "«"],
  ["raquo", "»"],
  ["bull", "•"],
  ["middot", "·"],
  ["copy", "©"],
  ["reg", "®"],
  ["trade", "™"],
  ["euro", "€"],
  ["times", "×"],
  ["deg", "°"],
]);

/** Longest `&...;` body looked at; bounds the work done per `&`. */
const MAX_ENTITY_BODY = 10;

function decodeEntityBody(body: string): string | undefined {
  if (body.charCodeAt(0) === 35 /* # */) {
    const hex = body.charCodeAt(1) === 120 /* x */ || body.charCodeAt(1) === 88; /* X */
    const digits = body.slice(hex ? 2 : 1);
    if (!(hex ? /^[0-9a-fA-F]{1,6}$/ : /^[0-9]{1,7}$/).test(digits)) return undefined;
    const cp = Number.parseInt(digits, hex ? 16 : 10);
    if (cp === 0 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff)) return "�";
    return String.fromCodePoint(cp);
  }
  return NAMED_ENTITIES.get(body);
}

/** Decodes standard named and numeric entities once; anything else stays literal. Linear. */
export function decodeEntities(text: string): string {
  let amp = text.indexOf("&");
  if (amp === -1) return text;
  const out: string[] = [];
  let last = 0;
  while (amp !== -1) {
    let semi = -1;
    const limit = Math.min(text.length, amp + MAX_ENTITY_BODY + 2);
    for (let j = amp + 2; j < limit; j += 1) {
      if (text.charCodeAt(j) === 59 /* ; */) {
        semi = j;
        break;
      }
    }
    const decoded = semi === -1 ? undefined : decodeEntityBody(text.slice(amp + 1, semi));
    if (decoded === undefined) {
      amp = text.indexOf("&", amp + 1);
      continue;
    }
    out.push(text.slice(last, amp), decoded);
    last = semi + 1;
    amp = text.indexOf("&", last);
  }
  out.push(text.slice(last));
  return out.join("");
}

// C0 controls (except tab, LF, CR), DEL, C1 controls and bidi override/isolate characters.
// eslint-disable-next-line no-control-regex
const SCRUB_RE = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f‪-‮⁦-⁩]/g;
const scrub = (text: string): string => text.replace(SCRUB_RE, "");

/** `text` cut to at most `max` UTF-16 units without splitting a surrogate pair. */
function safeSlice(text: string, max: number): string {
  if (text.length <= max) return text;
  let end = max;
  const last = text.charCodeAt(end - 1);
  if (end > 0 && last >= 0xd800 && last <= 0xdbff) end -= 1;
  return text.slice(0, Math.max(0, end));
}

/** Plain-text bodies: no interpretation at all, only the control-character scrub and the cap. */
export function capPlainText(text: string, maxChars: number): HtmlTextResult {
  const clipped = text.length > MAX_HTML_INPUT_CHARS ? safeSlice(text, MAX_HTML_INPUT_CHARS) : text;
  const cleaned = scrub(clipped);
  const capped = safeSlice(cleaned, maxChars);
  return {
    text: capped,
    truncated: clipped.length < text.length || capped.length < cleaned.length,
  };
}

const isWs = (c: number): boolean =>
  c === 32 || c === 160 || (c >= 9 && c <= 13); /* tab, LF, VT, FF, CR */
const isLetter = (c: number): boolean => (c >= 65 && c <= 90) || (c >= 97 && c <= 122);

const BLOCK_TAGS: ReadonlySet<string> = new Set([
  "div",
  "p",
  "table",
  "thead",
  "tbody",
  "tfoot",
  "tr",
  "blockquote",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "section",
  "article",
  "header",
  "footer",
  "aside",
  "nav",
  "main",
  "dl",
  "dt",
  "dd",
  "figure",
  "figcaption",
  "details",
  "summary",
  "address",
  "form",
  "fieldset",
]);
/** Tags whose attributes are read; every other tag is only classified by name. */
const ATTR_TAGS: ReadonlySet<string> = new Set(["a", "img", "attachment", "emoji", "customemoji"]);
const WANTED_ATTRS: ReadonlySet<string> = new Set(["href", "src", "alt", "title", "id"]);
/** Longest tag body whose attributes are parsed. */
const MAX_TAG_ATTR_CHARS = 8192;
const MAX_ALT_CHARS = 100;
const MAX_HREF_CHARS = 500;
const MAX_HOSTED_ID_CHARS = 1024;
const MAX_INDENT_LEVELS = 4;
/** Pending line prefix (list bullets, mention `@`) is bounded so its handling stays O(1). */
const MAX_PREFIX_CHARS = 64;

// eslint-disable-next-line no-control-regex
const SAFE_HREF_RE = /^(?:https?:\/\/|mailto:)[^\s<>"'\\\u0000-\u001f\u007f]+$/i;
const ATTACHMENT_ID_RE = /^[A-Za-z0-9._-]{1,128}$/;
const HOSTED_ID_RE = /^[A-Za-z0-9_\-=+.]+$/;
const DIGITS_RE = /^[0-9]{1,20}$/;
const GRAPH_PREFIX = "https://graph.microsoft.com/";

interface Attrs {
  get(name: string): string | undefined;
}

function parseAttrs(s: string, from: number): Attrs {
  const out = new Map<string, string>();
  const n = s.length;
  let i = from;
  while (i < n) {
    while (i < n && (isWs(s.charCodeAt(i)) || s.charCodeAt(i) === 47) /* / */) i += 1;
    if (i >= n) break;
    const nameStart = i;
    while (i < n) {
      const c = s.charCodeAt(i);
      if (isWs(c) || c === 61 /* = */ || c === 47) break;
      i += 1;
    }
    const name = s.slice(nameStart, i).toLowerCase();
    while (i < n && isWs(s.charCodeAt(i))) i += 1;
    let value = "";
    if (i < n && s.charCodeAt(i) === 61) {
      i += 1;
      while (i < n && isWs(s.charCodeAt(i))) i += 1;
      const q = s.charCodeAt(i);
      if (q === 34 || q === 39) {
        const close = s.indexOf(s[i] as string, i + 1);
        value = s.slice(i + 1, close === -1 ? n : close);
        i = close === -1 ? n : close + 1;
      } else {
        const start = i;
        while (i < n && !isWs(s.charCodeAt(i))) i += 1;
        value = s.slice(start, i);
      }
    }
    if (WANTED_ATTRS.has(name) && !out.has(name)) out.set(name, decodeEntities(value));
  }
  return out;
}

/**
 * The hosted-content id of an `<img src>` on `https://graph.microsoft.com/{v1.0|1.0}` whose path is
 * exactly one of (never fetched either way; undefined for anything else):
 * - chat: `/chats/{chatId}/messages/{messageId}/hostedContents/{id}/$value`;
 * - channel message: `/teams/{teamId}/channels/{channelId}/messages/{messageId}/hostedContents/{id}/$value`
 *   (not for a reply);
 * - channel reply: `/teams/{teamId}/channels/{channelId}/messages/{replyToId}/replies/{messageId}/hostedContents/{id}/$value`.
 * The team id is a GUID and compared case-insensitively; chat and channel ids are opaque and exact.
 */
function hostedContentId(src: string, ctx: HtmlTextContext): string | undefined {
  if (ctx.messageId === undefined) return undefined;
  const chatMode = ctx.chatId !== undefined;
  const channelMode = ctx.teamId !== undefined && ctx.channelId !== undefined;
  // Exactly one scope: a chat message never matches channel URLs and vice versa.
  if (chatMode === channelMode) return undefined;
  if (!src.startsWith(GRAPH_PREFIX) || src.length > 4096) return undefined;
  const cut = src.slice(GRAPH_PREFIX.length).split(/[?#]/, 1)[0] ?? "";
  const seg = cut.split("/");
  if (seg[0] !== "v1.0" && seg[0] !== "1.0") return undefined;
  let id: string | undefined;
  if (chatMode) {
    if (seg.length !== 8) return undefined;
    if (seg[1] !== "chats" || seg[3] !== "messages" || seg[5] !== "hostedContents") {
      return undefined;
    }
    if (seg[7] !== "$value") return undefined;
    if (!sameDecoded(seg[2], ctx.chatId)) return undefined;
    if (seg[4] !== ctx.messageId || !DIGITS_RE.test(seg[4] as string)) return undefined;
    id = seg[6];
  } else {
    // .../teams/{team}/channels/{channel}/messages/{message}/hostedContents/{id}/$value
    // .../teams/{team}/channels/{channel}/messages/{parent}/replies/{reply}/hostedContents/{id}/$value
    if (seg[1] !== "teams" || seg[3] !== "channels" || seg[5] !== "messages") return undefined;
    if (!sameDecoded(seg[2], ctx.teamId, true) || !sameDecoded(seg[4], ctx.channelId)) {
      return undefined;
    }
    const message = seg[6] as string;
    if (!DIGITS_RE.test(message)) return undefined;
    if (seg.length === 10 && seg[7] === "hostedContents" && seg[9] === "$value") {
      // A reply's images use the `/replies/{reply}` form below, which also pins the parent.
      if (ctx.replyToId !== undefined || message !== ctx.messageId) return undefined;
      id = seg[8];
    } else if (
      seg.length === 12 &&
      seg[7] === "replies" &&
      seg[9] === "hostedContents" &&
      seg[11] === "$value"
    ) {
      // A reply's own image: the reply is this message and its parent is the message's replyToId.
      if (seg[8] !== ctx.messageId || !DIGITS_RE.test(seg[8] as string)) return undefined;
      if (ctx.replyToId === undefined || message !== ctx.replyToId) return undefined;
      id = seg[10];
    } else {
      return undefined;
    }
  }
  return id !== undefined &&
    id.length > 0 &&
    id.length <= MAX_HOSTED_ID_CHARS &&
    HOSTED_ID_RE.test(id)
    ? id
    : undefined;
}

/**
 * Whether the percent-decoded URL segment equals `expected`: exactly, or ignoring case when
 * `ignoreCase` is set (used for the team GUID). Malformed escapes: false.
 */
function sameDecoded(
  segment: string | undefined,
  expected: string | undefined,
  ignoreCase = false,
): boolean {
  if (segment === undefined || expected === undefined) return false;
  try {
    const decoded = decodeURIComponent(segment);
    return ignoreCase ? decoded.toLowerCase() === expected.toLowerCase() : decoded === expected;
  } catch {
    return false;
  }
}

const compareForm = (text: string): string =>
  text
    .trim()
    .toLowerCase()
    .replace(/^(?:https?:\/\/|mailto:)/, "")
    .replace(/\/+$/, "");

/** Index just past the `</name ...>` that closes a raw-text element, or -1 if there is none. */
function findRawTextEnd(html: string, name: string, from: number): number {
  let p = from;
  for (;;) {
    const k = html.indexOf("</", p);
    if (k === -1) return -1;
    let match = true;
    for (let t = 0; t < name.length; t += 1) {
      if ((html.charCodeAt(k + 2 + t) | 32) !== name.charCodeAt(t)) {
        match = false;
        break;
      }
    }
    if (match) {
      const after = html.charCodeAt(k + 2 + name.length);
      if (Number.isNaN(after) || isWs(after) || after === 62 /* > */ || after === 47) {
        const gt = html.indexOf(">", k);
        return gt === -1 ? html.length : gt + 1;
      }
    }
    p = k + 2;
  }
}

/**
 * Converts an HTML fragment to plain text of at most `maxChars` characters (see the file header
 * for the contract). Block elements and `<br>` become line breaks (at most one blank line),
 * list items get `- ` markers (indented per nesting level), `<pre>` keeps its whitespace,
 * `<a href>` becomes `text (url)` when the URL differs from the text, `<at>` becomes `@Name`,
 * `<attachment id>` becomes `[attachment: id]`, `<emoji>` its glyph and `<img>` a placeholder.
 */
export function htmlToText(
  html: string,
  maxChars: number,
  ctx: HtmlTextContext = {},
): HtmlTextResult {
  const inputClipped = html.length > MAX_HTML_INPUT_CHARS;
  const src = inputClipped ? safeSlice(html, MAX_HTML_INPUT_CHARS) : html;
  const len = src.length;

  const parts: string[] = [];
  let outLen = 0;
  let truncated = inputClipped;
  let done = false;

  let pendingBreak = 0;
  let pendingSpace = false;
  let pendingPrefix = "";
  // Bounded so `endsWith`/`slice` on it stay O(1): crafted `<li></at>` runs cannot make it
  // grow with the input (which would make `</at>` handling quadratic).
  const addPrefix = (str: string): void => {
    if (pendingPrefix.length + str.length <= MAX_PREFIX_CHARS) pendingPrefix += str;
  };
  let preDepth = 0;
  let listDepth = 0;
  let anchor: { href: string; text: string } | undefined;

  const push = (str: string): void => {
    if (done || str.length === 0) return;
    const room = maxChars - outLen;
    let piece = str;
    if (piece.length > room) {
      piece = safeSlice(piece, Math.max(0, room));
      truncated = true;
      done = true;
    }
    parts.push(piece);
    outLen += piece.length;
  };

  const flushPending = (): void => {
    if (outLen > 0) {
      if (pendingBreak > 0) push("\n".repeat(pendingBreak));
      else if (pendingSpace) push(" ");
    }
    pendingBreak = 0;
    pendingSpace = false;
    if (pendingPrefix !== "") {
      const prefix = pendingPrefix;
      pendingPrefix = "";
      push(prefix);
    }
  };

  const breakAfter = (n: number): void => {
    pendingBreak = Math.max(pendingBreak, n);
  };

  const word = (w: string): void => {
    flushPending();
    push(w);
    if (anchor !== undefined && anchor.text.length < 300) anchor.text += w;
  };

  const emitText = (raw: string): void => {
    if (done || raw.length === 0) return;
    if (preDepth > 0) {
      flushPending();
      push(raw.replace(/\r\n?/g, "\n"));
      return;
    }
    const n = raw.length;
    let i = 0;
    while (i < n && !done) {
      if (isWs(raw.charCodeAt(i))) {
        pendingSpace = true;
        i += 1;
        continue;
      }
      let j = i + 1;
      while (j < n && !isWs(raw.charCodeAt(j))) j += 1;
      word(raw.slice(i, j));
      i = j;
    }
  };

  const closeAnchor = (): void => {
    const a = anchor;
    anchor = undefined;
    if (a === undefined || done) return;
    const text = a.text.trim();
    if (text === "" || compareForm(text) === compareForm(a.href)) return;
    pendingSpace = false;
    push(` (${a.href})`);
  };

  let noTagEnd = false;
  let i = 0;
  while (i < len && !done) {
    const lt = src.indexOf("<", i);
    const textEnd = lt === -1 ? len : lt;
    if (textEnd > i) emitText(decodeEntities(src.slice(i, textEnd)));
    if (lt === -1) break;
    i = lt;
    const c1 = src.charCodeAt(i + 1);

    if (c1 === 33 /* ! */ || c1 === 63 /* ? */) {
      if (src.startsWith("<!--", i)) {
        const end = src.indexOf("-->", i + 4);
        i = end === -1 ? len : end + 3;
      } else {
        const gt = src.indexOf(">", i + 2);
        i = gt === -1 ? len : gt + 1;
      }
      continue;
    }

    const closing = c1 === 47; /* / */
    const nameStart = i + (closing ? 2 : 1);
    if (noTagEnd || !isLetter(src.charCodeAt(nameStart))) {
      emitText("<");
      i += 1;
      continue;
    }
    const gt = src.indexOf(">", nameStart);
    if (gt === -1) {
      noTagEnd = true;
      emitText("<");
      i += 1;
      continue;
    }
    let nameEnd = nameStart;
    while (nameEnd < gt) {
      const c = src.charCodeAt(nameEnd);
      if (isWs(c) || c === 47 || c === 62) break;
      nameEnd += 1;
    }
    const name = src.slice(nameStart, nameEnd).toLowerCase();
    i = gt + 1;

    if (name === "script" || name === "style") {
      if (!closing) {
        const end = findRawTextEnd(src, name, i);
        i = end === -1 ? len : end;
      }
      continue;
    }

    const attrs: Attrs | undefined =
      !closing && ATTR_TAGS.has(name) && gt - nameEnd <= MAX_TAG_ATTR_CHARS
        ? parseAttrs(src.slice(nameEnd, gt), 0)
        : undefined;

    switch (name) {
      case "br":
        if (!closing) pendingBreak = Math.min(2, pendingBreak + 1);
        break;
      case "hr":
        breakAfter(1);
        break;
      case "ul":
      case "ol":
        breakAfter(1);
        listDepth = closing ? Math.max(0, listDepth - 1) : listDepth + 1;
        break;
      case "li":
        breakAfter(1);
        if (!closing) {
          addPrefix(`${"  ".repeat(Math.min(Math.max(0, listDepth - 1), MAX_INDENT_LEVELS))}- `);
        }
        break;
      case "pre":
        breakAfter(1);
        preDepth = closing ? Math.max(0, preDepth - 1) : preDepth + 1;
        break;
      case "td":
      case "th":
        pendingSpace = true;
        break;
      case "a":
        closeAnchor();
        if (!closing) {
          const href = attrs?.get("href")?.trim();
          if (href !== undefined && href.length <= MAX_HREF_CHARS && SAFE_HREF_RE.test(href)) {
            anchor = { href, text: "" };
          }
        }
        break;
      case "at":
        if (closing) {
          if (pendingPrefix.endsWith("@")) pendingPrefix = pendingPrefix.slice(0, -1);
        } else {
          addPrefix("@");
        }
        break;
      case "attachment": {
        if (closing) break;
        const id = attrs?.get("id");
        emitText(
          id !== undefined && ATTACHMENT_ID_RE.test(id) ? `[attachment: ${id}]` : "[attachment]",
        );
        break;
      }
      case "emoji": {
        if (closing) break;
        const glyph = (attrs?.get("alt") ?? attrs?.get("title") ?? "").trim();
        if (glyph !== "") emitText(safeSlice(glyph, MAX_ALT_CHARS));
        break;
      }
      case "customemoji": {
        if (closing) break;
        const alt = (attrs?.get("alt") ?? "").trim();
        if (alt !== "") emitText(`:${safeSlice(alt, MAX_ALT_CHARS)}:`);
        break;
      }
      case "img": {
        if (closing) break;
        // The genuine marker comes first, and alt text cannot contain a look-alike marker.
        const alt = safeSlice((attrs?.get("alt") ?? "").trim(), MAX_ALT_CHARS).replace(
          /hostedContentId\s*=/gi,
          "hostedContentId:",
        );
        const id = hostedContentId(attrs?.get("src") ?? "", ctx);
        emitText(
          `[image${id === undefined ? "" : ` hostedContentId=${id}`}${alt === "" ? "" : `: ${alt}`}]`,
        );
        break;
      }
      default:
        if (BLOCK_TAGS.has(name)) breakAfter(1);
    }
  }
  closeAnchor();

  const text = scrub(parts.join(""));
  return { text, truncated };
}
