import { describe, expect, it } from "vitest";
import {
  checkFormula,
  evaluateBook,
  evaluateExpression,
  evaluateSheet,
  rangeName,
  shiftFormula,
} from "../src/engine/sheet.ts";

const column = { A1: "1", A2: "2", A3: "3" };

describe("dynamic arrays", () => {
  it("spills a range into the cells below, filling them with its values", () => {
    const { values, spills, owner } = evaluateBook({ ...column, C1: "=A1:A3" });
    expect([values.C1, values.C2, values.C3]).toEqual([1, 2, 3]);
    expect(rangeName(spills.C1)).toBe("C1:C3");
    expect(owner).toEqual({ C2: "C1", C3: "C1" });
  });

  it("lifts operators and scalar functions over arrays", () => {
    const values = evaluateSheet({ ...column, C1: "=A1:A3*2", D1: "=-A1:A3", E1: "=ABS(A1:A3-3)" });
    expect([values.C1, values.C2, values.C3]).toEqual([2, 4, 6]);
    expect([values.D1, values.D2, values.D3]).toEqual([-1, -2, -3]);
    expect([values.E1, values.E2, values.E3]).toEqual([2, 1, 0]);
  });

  it("stretches a side of length 1 and gives #N/A where an array runs short", () => {
    const sums = evaluateSheet({ ...column, B1: "10", B2: "20", C1: "=A1:A3+B1:B2" });
    expect([sums.C1, sums.C2, sums.C3]).toEqual([11, 22, { error: "#N/A" }]);
    // A column against a row is every pairing.
    const grid = evaluateSheet({ ...column, G1: "10", H1: "100", E1: "=A1:A3*G1:H1" });
    expect([grid.E1, grid.F1, grid.E3, grid.F3]).toEqual([10, 100, 30, 300]);
  });

  it("aggregates an array without spilling it", () => {
    const { values, spills } = evaluateBook({ ...column, C1: "=SUM(A1:A3*2)" });
    expect(values.C1).toBe(12);
    expect(spills).toEqual({});
  });

  it("lets a cell read a spilled value before its formula has been worked out", () => {
    const values = evaluateSheet({ ...column, B1: "=D2", D1: "=A1:A3" });
    expect(values.B1).toBe(2);
  });

  it("shows #SPILL! with the intended range when a cell is in the way", () => {
    const { values, blocked, spills } = evaluateBook({ ...column, C1: "=A1:A3", C3: "x" });
    expect(values.C1).toEqual({ error: "#SPILL!" });
    expect(values.C2).toBeUndefined();
    expect(rangeName(blocked.C1)).toBe("C1:C3");
    expect(spills).toEqual({});
  });

  it("lets the earlier of two colliding spills win", () => {
    const { values } = evaluateBook({ ...column, C1: "=A1:A3", D1: "x", B2: "=A1:B1" });
    expect(values.B2).toEqual({ error: "#SPILL!" });
    expect([values.C1, values.C2, values.C3]).toEqual([1, 2, 3]);
  });

  it("refuses to spill off the sheet", () => {
    const { values } = evaluateBook({ ...column, C2: "=A1:A3" }, { bounds: { rows: 3, cols: 5 } });
    expect(values.C2).toEqual({ error: "#SPILL!" });
  });

  it("survives a spill that reads itself", () => {
    const values = evaluateSheet({ A1: "=A1:A2" });
    expect(values.A1).toBe(0);
  });

  it("returns an array whole from evaluateExpression", () => {
    expect(evaluateExpression(column, "=A1:A3*10")).toEqual([[10], [20], [30]]);
  });
});

describe("Excel details the array engine brings in", () => {
  it("shows 0 for a formula that points at an empty cell", () => {
    expect(evaluateSheet({ A1: "=Z99" }).A1).toBe(0);
  });

  it("treats a single-cell reference as a range in SUM, skipping text", () => {
    expect(evaluateSheet({ A1: "abc", B1: "=SUM(A1)", B2: "=SUM(A1,2)" })).toMatchObject({ B1: 0, B2: 2 });
  });

  it("counts errors and empty text in COUNTA, but not empty cells", () => {
    const values = evaluateSheet({ A1: "1", A2: "=1/0", A3: '=""', B1: "=COUNTA(A1:A4)", B2: "=COUNTA(A4)" });
    expect(values.B1).toBe(3);
    expect(values.B2).toBe(0);
  });

  it("joins numbers as text with 15 significant digits", () => {
    expect(evaluateSheet({ A1: '=0.1+0.2&""' }).A1).toBe("0.3");
  });

  it("accepts empty arguments, which IF treats as 0", () => {
    expect(() => checkFormula("=ROUND(2.5,)")).not.toThrow();
    const values = evaluateSheet({ A1: "=ROUND(2.5,)", A2: '=IF(TRUE,,"x")', A3: "=IF(FALSE,1)" });
    expect(values).toMatchObject({ A1: 3, A2: 0, A3: false });
  });

  it("works out IF element by element for an array test", () => {
    const values = evaluateSheet({ ...column, C1: '=IF(A1:A3>1,"big","small")' });
    expect([values.C1, values.C2, values.C3]).toEqual(["small", "big", "big"]);
  });
});

describe("the spilled range operator (A1#)", () => {
  it("refers to everything a formula spills, as in Microsoft's SUM(A2#) example", () => {
    const values = evaluateSheet({ A2: "=SEQUENCE(10)", C2: "=SUM(A2#)" });
    expect(values.C2).toBe(55);
    expect(evaluateSheet({ A2: "=SEQUENCE(20)", C2: "=SUM(A2#)" }).C2).toBe(210);
  });

  it("follows the spill as it grows and spills itself", () => {
    const { values } = evaluateBook({ A1: "=SEQUENCE(3)", C1: "=A1#*10" });
    expect([values.C1, values.C2, values.C3]).toEqual([10, 20, 30]);
  });

  it("is #REF! for a cell that doesn't spill", () => {
    expect(evaluateSheet({ A1: "5", B1: "=SUM(A1#)" }).B1).toEqual({ error: "#REF!" });
  });

  it("moves with a fill like any reference", () => {
    expect(shiftFormula("=SUM(A2#)", 1, 0)).toBe("=SUM(B2#)");
  });
});
