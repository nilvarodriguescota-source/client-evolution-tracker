import { describe, it, expect } from "vitest";
import {
  ExcelValidator,
  colIndexToLetter,
  cellRef,
  preciseRound,
  preciseMultiply,
  preciseSum,
  toDisplay,
} from "../lib/excel-validator";

// ---------------------------------------------------------------------------
// Utility functions
// ---------------------------------------------------------------------------

describe("colIndexToLetter", () => {
  it("converts single-letter columns", () => {
    expect(colIndexToLetter(0)).toBe("A");
    expect(colIndexToLetter(25)).toBe("Z");
  });

  it("converts multi-letter columns", () => {
    expect(colIndexToLetter(26)).toBe("AA");
    expect(colIndexToLetter(27)).toBe("AB");
  });
});

describe("cellRef", () => {
  it("builds Excel-style references", () => {
    expect(cellRef("Sheet1", 5, 0)).toBe("Sheet1!A5");
    expect(cellRef("Jan", 12, 3)).toBe("Jan!D12");
  });
});

describe("preciseRound", () => {
  it("rounds to specified decimal places", () => {
    expect(preciseRound(1.23456, 2)).toBe(1.23);
    expect(preciseRound(1.23556, 2)).toBe(1.24);
    expect(preciseRound(1.23456, 4)).toBe(1.2346);
  });

  it("handles edge cases", () => {
    expect(preciseRound(0, 2)).toBe(0);
    expect(preciseRound(-1.555, 2)).toBe(-1.55);
  });
});

describe("preciseMultiply", () => {
  it("multiplies with 4-digit precision", () => {
    expect(preciseMultiply(10.5, 3.2)).toBe(33.6);
    expect(preciseMultiply(0.1, 0.2)).toBe(0.02);
  });
});

describe("preciseSum", () => {
  it("sums with precision", () => {
    expect(preciseSum([0.1, 0.2, 0.3])).toBe(0.6);
    expect(preciseSum([100.1234, 200.5678])).toBe(300.6912);
  });
});

describe("toDisplay", () => {
  it("rounds to 2 decimal places for presentation", () => {
    expect(toDisplay(1.2345)).toBe(1.23);
    expect(toDisplay(1.2355)).toBe(1.24);
    expect(toDisplay(100)).toBe(100);
  });
});

// ---------------------------------------------------------------------------
// Layer 1: Existence
// ---------------------------------------------------------------------------

describe("Layer 1: Existence", () => {
  it("passes for present values", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkExistence("hello", 1, 0, "name")).toBe(true);
    expect(v.checkExistence(42, 1, 1, "qty")).toBe(true);
    expect(v.checkExistence(0, 1, 2, "zero")).toBe(true);
  });

  it("fails for null/undefined/empty", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkExistence(null, 1, 0, "field1")).toBe(false);
    expect(v.checkExistence(undefined, 2, 0, "field2")).toBe(false);
    expect(v.checkExistence("", 3, 0, "field3")).toBe(false);
    expect(v.checkExistence("  ", 4, 0, "field4")).toBe(false);

    const report = v.getReport();
    expect(report.failedChecks).toBe(4);
    expect(report.diagnostics).toHaveLength(4);
    expect(report.diagnostics[0].layer).toBe("existence");
    expect(report.diagnostics[0].severity).toBe("error");
  });
});

// ---------------------------------------------------------------------------
// Layer 2: Type
// ---------------------------------------------------------------------------

describe("Layer 2: Type", () => {
  it("validates number type", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkType(42, "number", 1, 0, "qty")).toBe(true);
    expect(v.checkType("123", "number", 1, 1, "price")).toBe(true);
    expect(v.checkType("abc", "number", 1, 2, "bad")).toBe(false);
    expect(v.checkType(null, "number", 1, 3, "null")).toBe(false);
  });

  it("validates string type", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkType("hello", "string", 1, 0, "name")).toBe(true);
    expect(v.checkType(123, "string", 1, 1, "code")).toBe(true); // numbers coerce
    expect(v.checkType(null, "string", 1, 2, "bad")).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Layer 3: Integrity
// ---------------------------------------------------------------------------

describe("Layer 3: Integrity", () => {
  it("validates numeric integrity", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkNumericIntegrity(100, 1, 0, "val")).toBe(true);
    expect(v.checkNumericIntegrity(Infinity, 1, 0, "bad")).toBe(false);
    expect(v.checkNumericIntegrity(NaN, 1, 0, "bad")).toBe(false);
    expect(v.checkNumericIntegrity(-5, 1, 0, "neg")).toBe(false);
    expect(v.checkNumericIntegrity(-5, 1, 0, "neg", { allowNegative: true })).toBe(true);
  });

  it("validates min/max ranges", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkNumericIntegrity(5, 1, 0, "val", { min: 10 })).toBe(false);
    expect(v.checkNumericIntegrity(5, 1, 0, "val", { max: 3 })).toBe(false);
    expect(v.checkNumericIntegrity(5, 1, 0, "val", { min: 1, max: 10 })).toBe(true);
  });

  it("validates string integrity", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkStringIntegrity("OK", 1, 0, "name")).toBe(true);
    expect(v.checkStringIntegrity("", 1, 0, "name")).toBe(false);
    expect(v.checkStringIntegrity("x".repeat(600), 1, 0, "name")).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Layer 4: Structure
// ---------------------------------------------------------------------------

describe("Layer 4: Structure", () => {
  it("validates row has minimum columns", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkStructure(["a", "b", "c", 1, 2, 3], 1, 6, "product")).toBe(true);
    expect(v.checkStructure(["a", "b"], 2, 6, "product")).toBe(false);
  });

  it("fails for non-array input", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkStructure("not array" as any, 1, 6, "product")).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Layer 5: Traceability
// ---------------------------------------------------------------------------

describe("Layer 5: Traceability", () => {
  it("creates traced values with origin", () => {
    const v = new ExcelValidator("Jan2024");
    const traced = v.traceValue(42.5, 10, 3);
    expect(traced.value).toBe(42.5);
    expect(traced.origin.sheet).toBe("Jan2024");
    expect(traced.origin.row).toBe(10);
    expect(traced.origin.col).toBe(3);
    expect(traced.origin.colLetter).toBe("D");
    expect(traced.origin.cellRef).toBe("Jan2024!D10");
    expect(traced.origin.rawValue).toBe(42.5);
  });
});

// ---------------------------------------------------------------------------
// Cross-cutting: Math consistency
// ---------------------------------------------------------------------------

describe("Math consistency", () => {
  it("passes when qty × price ≈ total", () => {
    const v = new ExcelValidator("Sheet1");
    const result = v.checkMathConsistency(10, 5.5, 55, 1, "Product A");
    expect(result.valid).toBe(true);
    expect(result.calculated).toBe(55);
    expect(result.difference).toBe(0);
  });

  it("fails when values diverge beyond tolerance", () => {
    const v = new ExcelValidator("Sheet1");
    const result = v.checkMathConsistency(10, 5.5, 60, 1, "Product A", 0.01);
    expect(result.valid).toBe(false);
    expect(result.difference).toBe(5);
  });

  it("handles small floating-point differences gracefully", () => {
    const v = new ExcelValidator("Sheet1");
    const result = v.checkMathConsistency(3, 1.1, 3.3, 1, "Product", 0.01);
    expect(result.valid).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Aggregation integrity
// ---------------------------------------------------------------------------

describe("Aggregation integrity", () => {
  it("passes when sum matches total", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkAggregation([10, 20, 30], 60, 1, "Client")).toBe(true);
  });

  it("fails when sum does not match total", () => {
    const v = new ExcelValidator("Sheet1");
    expect(v.checkAggregation([10, 20, 30], 65, 1, "Client")).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Duplicate detection
// ---------------------------------------------------------------------------

describe("Duplicate detection", () => {
  it("returns empty for unique items", () => {
    const v = new ExcelValidator("Sheet1");
    const dups = v.checkDuplicates(
      [
        { key: "A", label: "Item A", row: 1 },
        { key: "B", label: "Item B", row: 2 },
      ],
      "Test",
    );
    expect(dups).toHaveLength(0);
  });

  it("detects duplicates", () => {
    const v = new ExcelValidator("Sheet1");
    const dups = v.checkDuplicates(
      [
        { key: "A", label: "Item A", row: 1 },
        { key: "A", label: "Item A", row: 5 },
      ],
      "Test",
    );
    expect(dups).toHaveLength(1);
    expect(dups[0]).toBe("A");

    const report = v.getReport();
    expect(report.diagnostics.some((d) => d.layer === "duplicate")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Report generation
// ---------------------------------------------------------------------------

describe("Validation report", () => {
  it("generates complete report", () => {
    const v = new ExcelValidator("Sheet1");
    v.checkExistence("OK", 1, 0, "f1");
    v.checkExistence(null, 2, 0, "f2");
    v.checkType(42, "number", 3, 0, "f3");

    const report = v.getReport();
    expect(report.totalChecks).toBe(3);
    expect(report.passedChecks).toBe(2);
    expect(report.failedChecks).toBe(1);
    expect(report.valid).toBe(false); // has 1 error
    expect(report.layerSummary.existence).toEqual({ passed: 1, failed: 1 });
    expect(report.layerSummary.type).toEqual({ passed: 1, failed: 0 });
  });

  it("merges reports", () => {
    const v1 = new ExcelValidator("Sheet1");
    v1.checkExistence("OK", 1, 0, "f1");

    const v2 = new ExcelValidator("Sheet2");
    v2.checkExistence(null, 1, 0, "f2");

    v1.merge(v2);
    const report = v1.getReport();
    expect(report.totalChecks).toBe(2);
    expect(report.passedChecks).toBe(1);
    expect(report.diagnostics).toHaveLength(1);
  });
});
