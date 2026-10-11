/**
 * Fail-closed validation of catalog path-placeholder values.
 *
 * A captured placeholder value is DATA. Because the catalog allowlists exact path
 * patterns, a value such as `delta`, `$count`, `microsoft.graph.foo()` or `getByIds`
 * must never be able to turn `/groups/{group-id}` into a different (uncataloged) Graph
 * route. Every placeholder therefore has an allowlisted value SHAPE (below), and a set of
 * universal rules is applied on top to the value and to every further percent-decoding of
 * it. A placeholder name without a shape is rejected.
 *
 * Shapes are verified against Microsoft Graph documentation (site-get, chat-get,
 * mailfolder-get, well-known mail folder names, Exchange REST ids, driveItem/listItem ids).
 */

const MAX_DECODE_ROUNDS = 4;

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** `local@domain` UPN (guest `#EXT#` form included); exactly one "@" by construction. */
const UPN_RE = /^[A-Za-z0-9._+'#~-]{1,256}@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$/;
/** Exchange REST ids: base64url-ish (`AAMk...=`), never purely alphabetic. */
const EXCHANGE_ID_RE = /^[A-Za-z0-9_\-=+]{1,1024}$/;
/** SharePoint/OneDrive ids (`b!...`, `01ABC...`); never purely alphabetic. */
const DRIVE_ID_RE = /^[A-Za-z0-9!_\-=+.]{1,512}$/;
const DIGITS_RE = /^[0-9]{1,20}$/;
/** Teams thread ids: `19:...@thread.v2|tacv2|skype|unq.gbl.spaces`, plus the `48:notes` chat. */
const TEAMS_ID_RE = /^19:[A-Za-z0-9_.\-=]{1,512}@(?:thread\.(?:v2|tacv2|skype)|unq\.gbl\.spaces)$/;
const HOSTNAME_RE = /^[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$/;

const WELL_KNOWN_FOLDERS: ReadonlySet<string> = new Set([
  "archive",
  "clutter",
  "conflicts",
  "conversationhistory",
  "deleteditems",
  "drafts",
  "inbox",
  "junkemail",
  "localfailures",
  "msgfolderroot",
  "outbox",
  "recoverableitemsdeletions",
  "scheduled",
  "searchfolders",
  "sentitems",
  "serverfailures",
  "syncissues",
]);

/** Words that address OData/Graph routes rather than entities (compared case-insensitively). */
const RESERVED_WORDS: ReadonlySet<string> = new Set([
  "delta",
  "count",
  "ref",
  "value",
  "metadata",
  "batch",
  "$count",
  "$ref",
  "$value",
  "$metadata",
  "$batch",
  "$entity",
  "$all",
  "$each",
  "getbyids",
  "getavailableextensionproperties",
  "checkmembergroups",
  "checkmemberobjects",
  "getmembergroups",
  "getmemberobjects",
  "validateproperties",
  "restore",
]);

/** An id that is not just letters, so that a bare function/route name can never be one. */
function hasNonLetter(value: string): boolean {
  return /[^A-Za-z]/.test(value);
}

function isOpaque(value: string, re: RegExp): boolean {
  return re.test(value) && hasNonLetter(value);
}

const SHAPES: Readonly<Record<string, (value: string) => boolean>> = {
  "group-id": (v) => GUID_RE.test(v),
  "team-id": (v) => GUID_RE.test(v),
  "list-id": (v) => GUID_RE.test(v),
  "columnDefinition-id": (v) => GUID_RE.test(v),
  "user-id": (v) => GUID_RE.test(v) || UPN_RE.test(v),
  "listItem-id": (v) => DIGITS_RE.test(v),
  "chatMessage-id": (v) => DIGITS_RE.test(v),
  "chat-id": (v) => TEAMS_ID_RE.test(v) || v === "48:notes",
  "channel-id": (v) => TEAMS_ID_RE.test(v),
  "message-id": (v) => isOpaque(v, EXCHANGE_ID_RE),
  "event-id": (v) => isOpaque(v, EXCHANGE_ID_RE),
  "attachment-id": (v) => isOpaque(v, EXCHANGE_ID_RE),
  "calendar-id": (v) => isOpaque(v, EXCHANGE_ID_RE),
  "mailFolder-id": (v) => WELL_KNOWN_FOLDERS.has(v.toLowerCase()) || isOpaque(v, EXCHANGE_ID_RE),
  "drive-id": (v) => isOpaque(v, DRIVE_ID_RE),
  "site-id": (v) => {
    if (v.toLowerCase() === "root") return true;
    if (GUID_RE.test(v)) return true;
    const parts = v.split(",");
    if (parts.length === 2) return GUID_RE.test(parts[0] ?? "") && GUID_RE.test(parts[1] ?? "");
    if (parts.length === 3) {
      return (
        HOSTNAME_RE.test(parts[0] ?? "") &&
        GUID_RE.test(parts[1] ?? "") &&
        GUID_RE.test(parts[2] ?? "")
      );
    }
    return HOSTNAME_RE.test(v);
  },
};

/**
 * Pattern-scoped shapes, applied IN ADDITION to the shared shape of the same placeholder name
 * (key: `<pathPattern>|<name>`). `site-id` is shared by every `/sites/{site-id}/...` tool, but
 * Microsoft documents only `GET /sites/{hostname}:/{server-relative-path}` for the colon form
 * (site-getbypath), so there it must be a plain hostname: not `root`, not a composite id.
 * `get-user-presence` (D39): Learn's presence-get says "replace {id} with the user's GUID", so the
 * shared GUID-or-UPN `user-id` shape (other `/users/{user-id}/...` routes accept a UPN) is narrowed
 * for this one route to a GUID. One definition serves both the single path and graph-batch.
 */
const SCOPED_SHAPES: Readonly<Record<string, (value: string) => boolean>> = {
  "/sites/{site-id}:/{path}|site-id": (v) => HOSTNAME_RE.test(v),
  "/users/{user-id}/presence|user-id": (v) => GUID_RE.test(v),
};

/**
 * `value`, then each successive percent-decoding of it, until stable. Undefined when an
 * escape is malformed or the value still changes after the decode budget (fail closed).
 * Deliberate tradeoff: a literal "%" (sent as `%25`) decodes to a lone "%" whose next
 * decode is malformed, so it is inexpressible in any placeholder. No current id shape or
 * SharePoint site path admits "%"; revisit if a future skipEncoding param needs it.
 */
export function decodeForms(value: string): string[] | undefined {
  const forms = [value];
  let current = value;
  for (let i = 0; i < MAX_DECODE_ROUNDS; i += 1) {
    let decoded: string;
    try {
      decoded = decodeURIComponent(current);
    } catch {
      return undefined;
    }
    if (decoded === current) return forms;
    forms.push(decoded);
    current = decoded;
  }
  return undefined;
}

function universalViolation(form: string, reserved: boolean): string | undefined {
  const trimmed = form.trim();
  if (trimmed === "." || trimmed === "..") return "must not be a dot segment";
  if (trimmed.startsWith("$")) return "must not start with '$'";
  if (/[()]/.test(form)) return "must not contain parentheses";
  if (/^microsoft\.graph\./i.test(trimmed)) return "must not be a microsoft.graph function or type";
  if (reserved && RESERVED_WORDS.has(trimmed.toLowerCase()))
    return "is a reserved OData/Graph route word";
  return undefined;
}

/**
 * Why `actual` (the value Graph will see for the single path segment) may not be bound to
 * placeholder `name`, or undefined if acceptable.
 */
export function placeholderViolation(
  name: string,
  actual: string,
  pathPattern?: string,
): string | undefined {
  const forms = decodeForms(actual);
  if (forms === undefined) return "has malformed or excessive percent-encoding";
  for (const form of forms) {
    const reason = universalViolation(form, true);
    if (reason !== undefined) return reason;
  }
  const shape = SHAPES[name];
  if (shape === undefined) return "has no allowed value shape";
  if (!shape(actual)) return "does not have an allowed value shape";
  const scoped = pathPattern === undefined ? undefined : SCOPED_SHAPES[`${pathPattern}|${name}`];
  if (scoped !== undefined && !scoped(actual)) return "does not have an allowed value shape";
  return undefined;
}

/**
 * Same idea for a multi-segment `skipEncoding` value (site-relative path). It may not
 * contain ":" in any decoding (it would close Graph's path-based addressing early), and no
 * segment may be a dot segment, start with "$", contain parentheses, or start with
 * `microsoft.graph.`. Bare reserved words (`delta`, `count`, ...) are allowed here on
 * purpose: inside `:/{path}` they are ordinary SharePoint site names, not routes.
 */
export function pathValueViolation(value: string): string | undefined {
  const forms = decodeForms(value);
  if (forms === undefined) return "has malformed or excessive percent-encoding";
  const segments = value.split("/").length;
  for (const form of forms) {
    if (form.includes(":")) return "must not contain ':'";
    // Only literal "/" separates segments: an encoded one (%2F, %252F, ...) would surface as a
    // new segment after Graph decodes it, and the batch URL grammar rejects it too.
    if (form.split("/").length !== segments) return "must not contain an encoded '/'";
    // Never data in a SharePoint server-relative path; refused here so the single-request path
    // agrees with the stricter batch URL grammar instead of merely percent-encoding them.
    // eslint-disable-next-line no-control-regex
    if (/[?#\\\u0000-\u001f\u007f]/.test(form)) return "contains a forbidden character";
    for (const segment of form.split("/")) {
      const reason = universalViolation(segment, false);
      if (reason !== undefined) return reason;
    }
  }
  return undefined;
}
