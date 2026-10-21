import { describe, expect, it } from "vitest";
import {
  checkFormula,
  formulaReferences,
  copyRange,
  fillWrites,
  normalizeFormula,
  pasteWrites,
  type Range,
  evaluateSheet,
  formatValue,
  formulaRanges,
  FormulaSyntaxError,
  parseRange,
  rangeName,
  shiftFormula,
} from "../src/engine/sheet.ts";

describe("evaluateSheet", () => {
  const data = { A1: "1", A2: "2", A3: "2", A4: "5", A5: "text", B1: "3.1" };

  it("computes the statistics functions the assignment uses", () => {
    const values = evaluateSheet({
      ...data,
      C1: "=AVERAGE(A1:A5)",
      C2: "=MEDIAN(A1:A4)",
      C3: "=MODE(A1:A4)",
      C4: "=MAX(A1:A4)-MIN(A1:A4)",
      C5: "=STDEV.S(A1:A4)",
      C6: "=COUNTIF(A1:A4,A2)",
      C7: "=COUNT(A1:A5)",
      C8: "=SUM(A1:A4)*2+1",
    });
    expect(values.C1).toBe(2.5);
    expect(values.C2).toBe(2);
    expect(values.C3).toBe(2);
    expect(values.C4).toBe(4);
    expect(values.C5).toBeCloseTo(1.7320508, 6);
    expect(values.C6).toBe(2);
    expect(values.C7).toBe(4);
    expect(values.C8).toBe(21);
  });

  it("returns Excel errors", () => {
    const values = evaluateSheet({
      A1: "1",
      A2: "2",
      B1: "=MODE(A1:A2)",
      B2: "=A1/0",
      B3: "=NOPE(A1)",
      B4: "=AVERAGE(Z1:Z3)",
    });
    expect(values.B1).toEqual({ error: "#N/A" });
    expect(values.B2).toEqual({ error: "#DIV/0!" });
    expect(values.B3).toEqual({ error: "#NAME?" });
    expect(values.B4).toEqual({ error: "#DIV/0!" });
  });

  it("follows chains and survives circular references", () => {
    const values = evaluateSheet({ A1: "=B1+1", B1: "4", C1: "=C1+1" });
    expect(values.A1).toBe(5);
    expect(values.C1).toBe(1);
  });
});

describe("shiftFormula", () => {
  it("moves relative references and keeps anchored ones", () => {
    expect(shiftFormula("=AVERAGE(D3:D37)", 2, 0)).toBe("=AVERAGE(F3:F37)");
    expect(shiftFormula("=COUNTIF($G$3:$G$37,L13)", 0, 2)).toBe("=COUNTIF($G$3:$G$37,L15)");
    expect(shiftFormula("=A$1+$B2", 1, 1)).toBe("=B$1+$B3");
    expect(shiftFormula("=A1", -1, 0)).toBe("=#REF!");
  });

  it("does not mistake function names for references", () => {
    expect(shiftFormula("=STDEV.S(D3:D37)", 1, 0)).toBe("=STDEV.S(E3:E37)");
  });
});

describe("parsing helpers", () => {
  it("rejects formulas Excel would refuse", () => {
    expect(() => checkFormula("=AVERAGE(D3:D37")).toThrow(FormulaSyntaxError);
    expect(() => checkFormula("=1+")).toThrow(FormulaSyntaxError);
    expect(() => checkFormula("=MAX(D3:D37)-MIN(D3:D37)")).not.toThrow();
  });

  it("lists the ranges a formula points at", () => {
    expect(formulaRanges("=MAX(D3:D37)-MIN(E3:E37)+A1").map(rangeName)).toEqual(["D3:D37", "E3:E37", "A1"]);
  });

  it("normalizes ranges typed backwards", () => {
    const range = parseRange("D37:D3");
    expect(range && rangeName(range)).toBe("D3:D37");
  });
});

describe("formatValue", () => {
  it("applies the assignment number formats", () => {
    expect(formatValue(3.21857, "0.00", 65)).toBe("3.22");
    expect(formatValue(2618.857, "$#,##0.00", 80)).toBe("$2,618.86");
    expect(formatValue(26.4, "0", 65)).toBe("26");
  });

  it("fits General numbers to the column and overflows fixed formats with #", () => {
    expect(formatValue(2.948571428571, "General", 65)).toBe("2.948571");
    expect(formatValue(42, "General", 65)).toBe("42");
    expect(formatValue(123456.78, "$#,##0.00", 40)).toMatch(/^#+$/);
  });
});

describe("editing helpers", () => {
  it("tidies a formula on Enter", () => {
    expect(normalizeFormula("=average(d3:d37")).toBe("=AVERAGE(D3:D37)");
    expect(normalizeFormula("=max(d3:d37)-min(d3:d37")).toBe("=MAX(D3:D37)-MIN(D3:D37)");
    expect(normalizeFormula("hello")).toBe("hello");
  });

  it("fills right and down like the fill handle", () => {
    const inputs = { M4: "=AVERAGE(D3:D37)", M13: "=COUNTIF($G$3:$G$37,L13)" };
    const right = fillWrites(inputs, parseRange("M4") as Range, parseRange("M4:O4") as Range);
    expect(right).toEqual({ N4: "=AVERAGE(E3:E37)", O4: "=AVERAGE(F3:F37)" });
    const down = fillWrites(inputs, parseRange("M13") as Range, parseRange("M13:M15") as Range);
    expect(down.M15).toBe("=COUNTIF($G$3:$G$37,L15)");
  });

  it("pastes copies with shifted references and moves cuts", () => {
    const inputs = { M4: "=AVERAGE(D3:D37)", A1: "x" };
    expect(pasteWrites(copyRange(inputs, parseRange("M4") as Range), { col: 13, row: 3 })).toEqual({
      N4: "=AVERAGE(E3:E37)",
    });
    expect(pasteWrites(copyRange(inputs, parseRange("M4") as Range, true), { col: 13, row: 3 })).toEqual({
      M4: "",
      N4: "=AVERAGE(D3:D37)",
    });
  });
});

describe("assignment formats and references", () => {
  it("formats accounting and percent cells like Excel", () => {
    expect(formatValue(350.99, "accounting", 107)).toBe("$ 350.99 ");
    expect(formatValue(0, "accounting", 107)).toBe("$ -   ");
    expect(formatValue(-12.5, "accounting", 107)).toBe("$ (12.50)");
    expect(formatValue(0.067515, "0.00%", 100)).toBe("6.75%");
  });

  it("lists references with their $ signs and ranges", () => {
    expect(formulaReferences("=C6*$C$2+SUM(E6:E12)").map((r) => r.text)).toEqual(["C6", "$C$2", "E6", "E12"]);
    expect(rangeName(formulaReferences("=SUM(E6:E12)")[0].range)).toBe("E6:E12");
  });
});
