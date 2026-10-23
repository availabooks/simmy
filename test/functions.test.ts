/**
 * The functions MO-210 covers. Where Microsoft's support pages give a worked example, the test
 * uses it (see plans/MO-210_verification.md).
 */
import { describe, expect, it } from "vitest";
import { evaluateBook, evaluateExpression, evaluateSheet, type Inputs, rangeName } from "../src/engine/sheet.ts";

/** The value of one formula typed into Z1 next to `data`. */
const calc = (formula: string, data: Inputs = {}) => evaluateSheet({ ...data, Z1: formula }).Z1;
/** The array a formula returns. */
const array = (formula: string, data: Inputs = {}) => evaluateExpression(data, formula);

describe("text functions", () => {
  const fluid = { A2: "Fluid Flow" };

  it("LEFT and RIGHT take characters from either end", () => {
    expect(calc('=LEFT("Sale Price",4)')).toBe("Sale");
    expect(calc('=LEFT("Sweden")')).toBe("S");
    expect(calc('=RIGHT("Sale Price",5)')).toBe("Price");
    expect(calc('=LEFT("Excel",10)')).toBe("Excel");
    expect(calc('=RIGHT("Excel",0)')).toBe("");
    expect(calc('=LEFT("Excel",-1)')).toEqual({ error: "#VALUE!" });
    expect(calc("=LEFT(123.5,3)")).toBe("123");
  });

  it("MID takes characters from the middle", () => {
    expect(calc("=MID(A2,1,5)", fluid)).toBe("Fluid");
    expect(calc("=MID(A2,7,20)", fluid)).toBe("Flow");
    expect(calc("=MID(A2,20,5)", fluid)).toBe("");
    expect(calc("=MID(A2,0,5)", fluid)).toEqual({ error: "#VALUE!" });
    expect(calc("=MID(A2,2,-1)", fluid)).toEqual({ error: "#VALUE!" });
  });

  it("UPPER, LOWER and LEN", () => {
    expect(calc('=UPPER("total")')).toBe("TOTAL");
    expect(calc('=LOWER("Apt. 2B")')).toBe("apt. 2b");
    expect(calc('=LEN("Phoenix, AZ")')).toBe(11);
    expect(calc('=LEN("        One")')).toBe(11);
    expect(calc("=LEN(B1)")).toBe(0);
  });

  it("works element by element over a range", () => {
    const { values } = evaluateBook({ A1: "ann", A2: "bo", C1: "=UPPER(A1:A2)", D1: "=LEN(A1:A2)" });
    expect([values.C1, values.C2, values.D1, values.D2]).toEqual(["ANN", "BO", 3, 2]);
  });

  it("CONCAT joins ranges row by row, with no delimiter", () => {
    const grid = { B2: "a1", C2: "b1", B3: "a2", C3: "b2" };
    expect(calc("=CONCAT(B2:C3)", grid)).toBe("a1b1a2b2");
    expect(calc('=CONCAT(B2," ",C2)', grid)).toBe("a1 b1");
    expect(calc("=CONCAT(TRUE,1/0)")).toEqual({ error: "#DIV/0!" });
  });

  it("TEXTJOIN joins with a delimiter, skipping empty cells only when asked", () => {
    const data = { A1: "a", A2: "", A3: "c", B1: "-", B2: "+" };
    expect(calc('=TEXTJOIN(", ",TRUE,A1:A3)', data)).toBe("a, c");
    expect(calc('=TEXTJOIN(", ",FALSE,A1:A3)', data)).toBe("a, , c");
    // A range of delimiters is used in turn.
    expect(calc('=TEXTJOIN(B1:B2,TRUE,A1,"b",A3,"d")', data)).toBe("a-b+c-d");
    expect(calc('=TEXTJOIN(",",TRUE)')).toEqual({ error: "#VALUE!" });
  });
});

describe("counting", () => {
  it("COUNTBLANK counts empty cells and empty text, not zeros", () => {
    const data = { A1: "6", A2: '=""', A3: "0", A4: "27" };
    expect(calc("=COUNTBLANK(A1:A5)", data)).toBe(2);
    expect(calc("=COUNTBLANK(5)")).toEqual({ error: "#VALUE!" });
  });

  it("COUNTA counts everything but empty cells (Microsoft's example gives 5)", () => {
    const data = { A2: "39790", A3: "19", A4: "22.24", A5: "TRUE", A6: "=1/0" };
    expect(calc("=COUNTA(A2:A6)", data)).toBe(5);
    expect(calc("=COUNT(A2:A6)", data)).toBe(3);
  });
});

describe("SEQUENCE", () => {
  it("fills rows, then columns, from a start by a step", () => {
    expect(array("=SEQUENCE(2,3)")).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ]);
    expect(array("=SEQUENCE(5,1,1001,1000)")).toEqual([[1001], [2001], [3001], [4001], [5001]]);
    expect(array("=SEQUENCE(3,,10)")).toEqual([[10], [11], [12]]);
  });

  it("refuses fewer than one row or column", () => {
    expect(calc("=SEQUENCE(0)")).toEqual({ error: "#CALC!" });
    expect(calc("=SEQUENCE(2,-1)")).toEqual({ error: "#CALC!" });
  });

  it("spills, and A1-style references read it (SUM of SEQUENCE(10) is 55)", () => {
    const { values, spills } = evaluateBook({ A2: "=SEQUENCE(10)", C2: "=SUM(A2:A11)" });
    expect(rangeName(spills.A2)).toBe("A2:A11");
    expect(values.C2).toBe(55);
  });
});

describe("SORT", () => {
  const data = { A1: "pear", B1: "3", A2: "Apple", B2: "1", A3: "fig", B3: "2" };

  it("sorts rows by the first column, ignoring case", () => {
    expect(array("=SORT(A1:B3)", data)).toEqual([
      ["Apple", 1],
      ["fig", 2],
      ["pear", 3],
    ]);
  });

  it("sorts by another column, descending", () => {
    expect(array("=SORT(A1:B3,2,-1)", data)).toEqual([
      ["pear", 3],
      ["fig", 2],
      ["Apple", 1],
    ]);
  });

  it("sorts columns with by_col", () => {
    expect(array("=SORT(A1:C1,1,1,TRUE)", { A1: "3", B1: "1", C1: "2" })).toEqual([[1, 2, 3]]);
  });

  it("orders numbers, then text, then logicals, with blanks last", () => {
    const mixed = { A1: "b", A2: "TRUE", A3: "2", A5: "a" };
    expect(array("=SORT(A1:A5)", mixed)).toEqual([[2], ["a"], ["b"], [true], [0]]);
    expect(array("=SORT(A1:A5,1,-1)", mixed)).toEqual([[true], ["b"], ["a"], [2], [0]]);
  });

  it("rejects a sort column or order that doesn't exist", () => {
    expect(calc("=SORT(A1:B3,3)", data)).toEqual({ error: "#VALUE!" });
    expect(calc("=SORT(A1:B3,1,0)", data)).toEqual({ error: "#VALUE!" });
  });
});

describe("UNIQUE", () => {
  const data = { A1: "red", A2: "Red", A3: "blue", A4: "green", A5: "blue" };

  it("keeps the first of each value, ignoring case", () => {
    expect(array("=UNIQUE(A1:A5)", data)).toEqual([["red"], ["blue"], ["green"]]);
  });

  it("keeps only values that appear once with exactly_once", () => {
    expect(array("=UNIQUE(A1:A5,FALSE,TRUE)", data)).toEqual([["green"]]);
    expect(calc("=UNIQUE(A1:A2,FALSE,TRUE)", data)).toEqual({ error: "#CALC!" });
  });

  it("compares whole rows, and tells 1 from \"1\"", () => {
    const rows = { A1: "a", B1: "1", A2: "a", B2: '="1"', A3: "a", B3: "1" };
    expect(array("=UNIQUE(A1:B3)", rows)).toEqual([
      ["a", 1],
      ["a", "1"],
    ]);
  });

  it("works with SORT, as in Microsoft's tip", () => {
    const names = { A1: "Ann", B1: "Lee", A2: "Bo", B2: "Kim", A3: "Ann", B3: "Lee" };
    expect(array('=SORT(UNIQUE(B1:B3&" "&A1:A3))', names)).toEqual([["Kim Bo"], ["Lee Ann"]]);
  });
});

describe("RANDBETWEEN", () => {
  it("returns a whole number in range, the same for the same seed", () => {
    const data = { A1: "=RANDBETWEEN(1,100)", A2: "=RANDBETWEEN(1,100)" };
    const first = evaluateSheet(data, { seed: 7 });
    expect(evaluateSheet(data, { seed: 7 })).toEqual(first);
    for (const value of [first.A1, first.A2]) {
      expect(Number.isInteger(value)).toBe(true);
      expect(value).toBeGreaterThanOrEqual(1);
      expect(value).toBeLessThanOrEqual(100);
    }
  });

  it("changes when the seed does, as on a recalculation", () => {
    const data = { A1: "=RANDBETWEEN(1,1000000)" };
    const seen = new Set([1, 2, 3, 4, 5].map((seed) => evaluateSheet(data, { seed }).A1));
    expect(seen.size).toBeGreaterThan(1);
  });

  it("rounds its bounds inward and rejects bottom above top", () => {
    const values = Array.from({ length: 20 }, (_, seed) => evaluateSheet({ A1: "=RANDBETWEEN(1.2,3.8)" }, { seed }).A1);
    expect(new Set(values).size).toBeGreaterThan(0);
    for (const value of values) {
      expect([2, 3]).toContain(value);
    }
    expect(calc("=RANDBETWEEN(5,1)")).toEqual({ error: "#NUM!" });
  });
});
