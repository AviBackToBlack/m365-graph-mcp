import { describe, expect, it } from "vitest";
import { isExtendedPropertyId } from "./extended-properties.js";

const GUID = "{00020329-0000-0000-C000-000000000046}";

describe("isExtendedPropertyId numeric ranges", () => {
  it("accepts GUID-backed Id values only in 0x8000-0xFFFE", () => {
    expect(isExtendedPropertyId(`String ${GUID} Id 0x8000`)).toBe(true);
    expect(isExtendedPropertyId(`String ${GUID} Id 0xFFFE`)).toBe(true);
    for (const id of ["0x1", "0x7FFF", "0xFFFF", "0x10000", "0x0"]) {
      expect(isExtendedPropertyId(`String ${GUID} Id ${id}`), id).toBe(false);
    }
  });

  it("accepts bare proptags only in 0x0001-0x7FFF", () => {
    expect(isExtendedPropertyId("String 0x0001")).toBe(true);
    expect(isExtendedPropertyId("String 0x7FFF")).toBe(true);
    for (const id of ["0x0", "0x8000", "0xFFFE", "0x10000"]) {
      expect(isExtendedPropertyId(`String ${id}`), id).toBe(false);
    }
  });

  it("leaves the Name form unaffected", () => {
    expect(isExtendedPropertyId(`String ${GUID} Name MyProp`)).toBe(true);
    // A name that merely ends in " Id 0x.." is still a Name-form id (no numeric range).
    expect(isExtendedPropertyId(`String ${GUID} Name Foo Id 0x10`)).toBe(true);
  });
});
