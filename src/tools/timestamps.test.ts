import { describe, expect, it } from "vitest";
import { canonicalInstant, isRealUtcTimestamp, utcTimestampMs } from "./timestamps.js";
import { canonicalTimeZone } from "./timezones.js";

describe("isRealUtcTimestamp", () => {
  it("accepts real UTC instants with optional fraction", () => {
    for (const v of [
      "2024-01-31T00:00:00Z",
      "2024-02-29T23:59:59Z",
      "2024-01-01T00:00:00.1Z",
      "2024-01-01T00:00:00.0000000Z",
    ]) {
      expect(isRealUtcTimestamp(v), v).toBe(true);
    }
    expect(utcTimestampMs("1970-01-01T00:00:01Z")).toBe(1000);
  });

  it("rejects impossible dates, offsets, local times and other shapes", () => {
    for (const v of [
      "2024-02-30T00:00:00Z",
      "2023-02-29T00:00:00Z",
      "2024-04-31T00:00:00Z",
      "2024-01-01T24:00:00Z",
      "2024-01-01T00:60:00Z",
      "2024-01-01T00:00:60Z",
      "0024-01-01T00:00:00Z",
      "2024-01-01T00:00:00",
      "2024-01-01T00:00:00+00:00",
      "2024-01-01T00:00:00-08:00",
      "2024-01-01T00:00:00z",
      "2024-01-01t00:00:00Z",
      "2024-01-01 00:00:00Z",
      "2024-01-01T00:00Z",
      "2024-01-01T00:00:00.12345678Z",
      "2024-01-01",
      " 2024-01-01T00:00:00Z",
      "2024-01-01T00:00:00Z\n",
      "",
    ]) {
      expect(isRealUtcTimestamp(v), JSON.stringify(v)).toBe(false);
    }
  });
});

describe("canonicalTimeZone", () => {
  it("resolves IANA, Windows and UTC names to their canonical spelling", () => {
    expect(canonicalTimeZone("Europe/London")).toBe("Europe/London");
    expect(canonicalTimeZone("europe/london")).toBe("Europe/London");
    expect(canonicalTimeZone("utc")).toBe("UTC");
    expect(canonicalTimeZone("pacific standard time")).toBe("Pacific Standard Time");
    expect(canonicalTimeZone("Central Standard Time")).toBe("Central Standard Time");
    // Documented IANA names that some ICU builds lack in supportedValuesOf.
    expect(canonicalTimeZone("Asia/Kolkata")).toBe("Asia/Kolkata");
    expect(canonicalTimeZone("Europe/Kyiv")).toBe("Europe/Kyiv");
  });

  it("rejects everything outside the allowlist, including header-breaking text", () => {
    for (const v of [
      "",
      "Nowhere/Land",
      'UTC"',
      "UTC, outlook.allow-unsafe-html",
      "UTC\r\nX: y",
      "UTC;q=1",
      "Pacific Standard Time ",
      "Türkiye Standard Time",
      "__proto__",
      "constructor",
    ]) {
      expect(canonicalTimeZone(v), JSON.stringify(v)).toBeUndefined();
    }
  });

  it("only yields printable-ASCII names without quotes or commas", () => {
    for (const v of ["Europe/Berlin", "Central Standard Time (Mexico)", "UTC+12", "Etc/GMT+12"]) {
      const c = canonicalTimeZone(v) as string;
      expect(c).toBe(v);
      expect(/^[\x20-\x7e]+$/.test(c) && !/["',;]/.test(c)).toBe(true);
    }
  });
});

describe("canonicalInstant", () => {
  it("floors pre-1970 fractional instants instead of emitting a negative fraction", () => {
    expect(canonicalInstant("1969-12-31T23:59:59.5Z")).toBe("1969-12-31T23:59:59.5Z");
    expect(canonicalInstant("1969-12-31T23:59:59.0000001Z")).toBe("1969-12-31T23:59:59.0000001Z");
    expect(canonicalInstant("1970-01-01T00:59:59.25+01:00")).toBe("1969-12-31T23:59:59.25Z");
    expect(canonicalInstant("1950-06-15T12:00:00.1234567Z")).toBe("1950-06-15T12:00:00.1234567Z");
  });
});
