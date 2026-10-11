/**
 * Strict timestamp validation shared by every tool that takes a date/time as data.
 *
 * Strict UTC (`utcTimestampMs` and friends, used by delta `receivedAfter`): only the UTC `Z`
 * form is accepted (`2024-01-31T00:00:00Z`, optional 1-7 fraction digits); offsets and
 * offset-less local times are rejected. Calendar `startDateTime`/`endDateTime` use the
 * offset-aware `instantTicks`/`canonicalInstant` at the end of this file instead.
 */
const ISO_UTC_RE = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d{1,7}))?Z$/;

/**
 * Milliseconds since the epoch of a real UTC instant, or undefined. `Date.parse` normalizes
 * impossible dates (2024-02-30 -> March 1), so the parsed components must round-trip to
 * exactly what the caller wrote. Fraction digits beyond the whole second are ignored.
 */
export function utcTimestampMs(value: string): number | undefined {
  const m = ISO_UTC_RE.exec(value);
  if (m === null) return undefined;
  const [y, mo, d, h, mi, s] = m.slice(1, 7).map(Number) as [
    number,
    number,
    number,
    number,
    number,
    number,
  ];
  const t = new Date(Date.UTC(y, mo - 1, d, h, mi, s));
  const real =
    t.getUTCFullYear() === y &&
    t.getUTCMonth() === mo - 1 &&
    t.getUTCDate() === d &&
    t.getUTCHours() === h &&
    t.getUTCMinutes() === mi &&
    t.getUTCSeconds() === s;
  return real ? t.getTime() : undefined;
}

/** True for a real UTC instant in the strict `Z` form (see `utcTimestampMs`). */
export function isRealUtcTimestamp(value: string): boolean {
  return utcTimestampMs(value) !== undefined;
}

/**
 * Offset-aware instants (calendar `startDateTime`/`endDateTime`, tool boundary only).
 *
 * Accepts strict ISO-8601 with `Z` or an explicit `+HH:MM` / `-HH:MM` offset and the same
 * optional 1-7 fraction digits as above. Offset-less (local) timestamps are rejected. Offsets
 * are bounded to +/-14:00 (the real-world range), minutes 00-59; `-00:00` is rejected because
 * RFC 3339 gives it the meaning "unknown local offset" (write `Z` or `+00:00`). The wall-clock
 * fields must be a real date (no 02-30 or hour 24), and the resulting UTC instant must fall in
 * years 0100-9999 so it can be written back in four-digit form.
 *
 * The instant is exact (BigInt 100 ns ticks) and is re-emitted by `canonicalInstant` as UTC `Z`
 * so Graph never has to interpret an offset (the instances docs only show offset-less values)
 * and a raw `+` never reaches a query string.
 */
const ISO_INSTANT_RE =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d{1,7}))?(?:Z|([+-])(\d{2}):(\d{2}))$/;
const MAX_OFFSET_MINUTES = 14 * 60;
const TICKS_PER_SECOND = 10_000_000n;

/** Exact UTC instant in 100 ns ticks since the epoch, or undefined (see above). */
export function instantTicks(value: string): bigint | undefined {
  const m = ISO_INSTANT_RE.exec(value);
  if (m === null) return undefined;
  const [y, mo, d, h, mi, s] = m.slice(1, 7).map(Number) as [
    number,
    number,
    number,
    number,
    number,
    number,
  ];
  const wall = new Date(Date.UTC(y, mo - 1, d, h, mi, s));
  const real =
    wall.getUTCFullYear() === y &&
    wall.getUTCMonth() === mo - 1 &&
    wall.getUTCDate() === d &&
    wall.getUTCHours() === h &&
    wall.getUTCMinutes() === mi &&
    wall.getUTCSeconds() === s;
  if (!real) return undefined;
  let offsetMinutes = 0;
  if (m[8] !== undefined) {
    const oh = Number(m[9]);
    const om = Number(m[10]);
    const total = oh * 60 + om;
    if (om > 59 || total > MAX_OFFSET_MINUTES) return undefined;
    if (m[8] === "-" && total === 0) return undefined;
    offsetMinutes = m[8] === "-" ? -total : total;
  }
  const utcMs = wall.getTime() - offsetMinutes * 60_000;
  const year = new Date(utcMs).getUTCFullYear();
  if (year < 100 || year > 9999) return undefined;
  const fraction = (m[7] ?? "").padEnd(7, "0");
  return BigInt(utcMs / 1000) * TICKS_PER_SECOND + BigInt(fraction);
}

/**
 * Canonical UTC form of an instant: `YYYY-MM-DDTHH:MM:SS[.f]Z`, fraction digits trimmed of
 * trailing zeros and omitted when zero (so `...:00Z` stays byte-identical and `.0000000Z`
 * becomes `...:00Z`). Undefined when `value` is not a valid instant per `instantTicks`.
 */
export function canonicalInstant(value: string): string | undefined {
  const ticks = instantTicks(value);
  if (ticks === undefined) return undefined;
  // Floor division: BigInt `/` and `%` truncate toward zero, which for pre-1970 instants would
  // round the second up and leave a negative fraction (e.g. "00:00:00.-5Z").
  let seconds = ticks / TICKS_PER_SECOND;
  let remainder = ticks % TICKS_PER_SECOND;
  if (remainder < 0n) {
    seconds -= 1n;
    remainder += TICKS_PER_SECOND;
  }
  const iso = new Date(Number(seconds) * 1000).toISOString(); // years 0100-9999: 4 digits
  const fraction = remainder.toString().padStart(7, "0").replace(/0+$/, "");
  return `${iso.slice(0, 19)}${fraction === "" ? "" : `.${fraction}`}Z`;
}
