import { describe, expect, it } from "vitest";
import {
  absoluteRef,
  formulaNames,
  guessLabelSides,
  labelToName,
  lookupName,
  nameForRange,
  nameProblem,
  namesFromSelection,
  parseRefersTo,
  refersToText,
  suggestName,
} from "../src/engine/names.ts";
import {
  evaluateBook,
  evaluateSheet,
  formulaRanges,
  formulaReferences,
  normalizeFormula,
  parseRange,
  type Range,
  rangeName,
  shiftFormula,
} from "../src/engine/sheet.ts";

const range = (text: string) => parseRange(text) as Range;

describe("naming rules (Microsoft's names-in-formulas page)", () => {
  it("accepts names Excel accepts", () => {
    for (const name of ["Rate", "_x", "\\x", "Q1_", "ABCD1", "Tax.Rate", "a".repeat(255)]) {
      expect(nameProblem(name), name).toBeNull();
    }
  });

  it("refuses names Excel refuses", () => {
    const refused = ["", "1x", "A1", "xfd99", "R", "c", "R1C1", "RC", "Tax Rate", "TRUE", "a-b", "a".repeat(256)];
    for (const name of refused) {
      expect(nameProblem(name), name).not.toBeNull();
    }
  });
});

describe("names in formulas", () => {
  const names = { Rate: "$G$2", Prices: "$A$1:$A$3" };
  const data = { A1: "10", A2: "20", A3: "30", G2: "0.5" };

  it("evaluate as the range they refer to, ignoring case", () => {
    const values = evaluateSheet({ ...data, B1: "=A1*Rate", B2: "=a2*rate", B3: "=SUM(Prices)" }, { names });
    expect([values.B1, values.B2, values.B3]).toEqual([5, 10, 60]);
  });

  it("spill when they name more than one cell", () => {
    const { values, spills } = evaluateBook({ ...data, C1: "=Prices*2" }, { names });
    expect([values.C1, values.C2, values.C3]).toEqual([20, 40, 60]);
    expect(rangeName(spills.C1)).toBe("C1:C3");
  });

  it("give #NAME? when the name isn't defined, until it is", () => {
    expect(evaluateSheet({ ...data, B1: "=A1*Bonus" }).B1).toEqual({ error: "#NAME?" });
    expect(evaluateSheet({ ...data, B1: "=A1*Bonus" }, { names: { Bonus: "$G$2" } }).B1).toBe(5);
  });

  it("are written as defined when the formula is entered; unknown names stay as typed", () => {
    expect(normalizeFormula("=sum(prices)*rate+other", names)).toBe("=SUM(Prices)*Rate+other");
  });

  it("are left alone when a formula is filled", () => {
    expect(shiftFormula("=Rate*C6", 0, 1)).toBe("=Rate*C7");
  });

  it("count as their references for grading and outlines", () => {
    expect(formulaReferences("=A1*Rate", names).map((r) => r.text)).toEqual(["A1", "$G$2"]);
    expect(formulaRanges("=SUM(Prices)", names).map(rangeName)).toEqual(["A1:A3"]);
    expect(formulaNames("=SUM(Prices)*Rate+TRUE")).toEqual(["Prices", "Rate"]);
  });
});

describe("name helpers", () => {
  it("look names up without regard to case", () => {
    expect(lookupName({ Rate: "$G$2" }, "RATE")).toEqual({ name: "Rate", ref: "$G$2" });
    expect(lookupName({ Rate: "$G$2" }, "Bonus")).toBeNull();
  });

  it("write and read references the way Define Name shows them", () => {
    expect(absoluteRef(range("C2"))).toBe("$C$2");
    expect(absoluteRef(range("A2:A9"))).toBe("$A$2:$A$9");
    expect(refersToText(range("B1"), "Sheet1")).toBe("=Sheet1!$B$1");
    expect(refersToText(range("B1"), "My Sheet")).toBe("='My Sheet'!$B$1");
    expect(rangeName(parseRefersTo("='My Sheet'!$C$2:$D$4", "My Sheet") as Range)).toBe("C2:D4");
    expect(rangeName(parseRefersTo("C2", "Sheet1") as Range)).toBe("C2");
    expect(parseRefersTo("=Other!$C$2", "Sheet1")).toBeNull();
  });

  it("find the name the Name Box shows for a selection", () => {
    expect(nameForRange({ Rate: "$G$2", Sales: "$D$2:$D$9" }, range("D2:D9"))).toBe("Sales");
    expect(nameForRange({ Rate: "$G$2" }, range("G3"))).toBeNull();
  });

  it("turn labels into names the way Create from Selection does", () => {
    expect(labelToName("Unit Price")).toBe("Unit_Price");
    expect(labelToName("2024")).toBe("_2024");
    expect(labelToName("Q1")).toBe("Q1_");
    expect(labelToName("  ")).toBeNull();
  });
});

describe("Create from Selection", () => {
  const table = {
    A1: "Region",
    B1: "Q1",
    C1: "Q2",
    A2: "East",
    B2: "10",
    C2: "20",
    A3: "West",
    B3: "30",
    C3: "40",
  };

  it("checks Top row and Left column when those hold labels", () => {
    expect(guessLabelSides(table, range("A1:C3"))).toMatchObject({ top: true, left: true });
    expect(guessLabelSides(table, range("B1:B3"))).toMatchObject({ top: true, left: false });
    const row = guessLabelSides({ A1: "Rate", B1: "0.05" }, range("A1:B1"));
    expect(row).toMatchObject({ top: false, left: true });
  });

  it("names each column and row beside its label, leaving out the corner", () => {
    const sides = { top: true, left: true, bottom: false, right: false };
    const names = namesFromSelection(table, range("A1:C3"), sides);
    expect(names).toEqual({ Q1_: "$B$2:$B$3", Q2_: "$C$2:$C$3", East: "$B$2:$C$2", West: "$B$3:$C$3" });
  });

  it("skips empty labels", () => {
    const names = namesFromSelection({ A1: "Sales" }, range("A1:B4"), {
      top: true,
      left: false,
      bottom: false,
      right: false,
    });
    expect(names).toEqual({ Sales: "$A$2:$A$4" });
  });

  it("suggests a name from the label to the left or above", () => {
    expect(suggestName({ F2: "Tax Rate" }, range("G2"))).toBe("Tax_Rate");
    expect(suggestName({ D1: "Sales" }, range("D2:D9"))).toBe("Sales");
    expect(suggestName({}, range("D2"))).toBe("");
  });
});
