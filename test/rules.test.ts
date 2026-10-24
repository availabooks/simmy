/**
 * The rule kinds that mirror more of the add-in's grading steps: value over a spilled range,
 * equals (value_correct), contains (formula_contains), name (check_name_exists and
 * check_named_range), between (value_between) and selected (activecell).
 */
import { describe, expect, it } from "vitest";
import type { WorkbookState } from "../src/engine/workbook.ts";
import { gradeAssignment, solvedState } from "../src/module/grade.ts";
import { ModuleError, parseModule } from "../src/module/parse.ts";

/** A one-task module with these rules, over a small sheet. */
const moduleWith = (rules: object[], task: object = {}) =>
  parseModule({
    workbook: { name: "Book1" },
    assignment: { title: "Rules", text: "Do the task." },
    sheet: { cells: { A1: "pear", A2: "Apple", A3: "fig", G2: 0.5 } },
    tasks: [{ text: "Work.", range: "C1", solution: "=SORT(A1:A3)", rules, ...task }],
  });

const rule = (extra: object) => ({ points: 1, message: "missed", ...extra });

/** Points earned for a workbook. */
const earned = (rules: object[], state: Partial<WorkbookState>, task: object = {}) =>
  gradeAssignment(moduleWith(rules, task), { inputs: {}, names: {}, seed: 1, ...state }).earned;

const problems = (rules: object[]) => {
  try {
    moduleWith(rules);
  } catch (error) {
    return (error as ModuleError).problems;
  }
  return [];
};

const sheet = { A1: "pear", A2: "Apple", A3: "fig", G2: "0.5" };

describe("value over a range (a spilled result)", () => {
  const sorted = [rule({ kind: "value", range: "C1:C3", formula: "=SORT($A$1:$A$3)" })];

  it("passes when the spill fills the range with the right values", () => {
    expect(earned(sorted, { inputs: { ...sheet, C1: "=SORT(A1:A3)" } })).toBe(1);
  });

  it("fails for the wrong order, or a range of the wrong size", () => {
    expect(earned(sorted, { inputs: { ...sheet, C1: "=A1:A3" } })).toBe(0);
    const short = [rule({ kind: "value", range: "C1:C2", formula: "=SORT($A$1:$A$3)" })];
    expect(earned(short, { inputs: { ...sheet, C1: "=SORT(A1:A3)" } })).toBe(0);
  });
});

describe("grading an array that runs off the sheet", () => {
  it("sees #SPILL!, as the student does, not the spilled values", () => {
    const mod = parseModule({
      workbook: { name: "Book1" },
      assignment: { title: "Edge", text: "Count." },
      sheet: { rows: 10, columns: 3 },
      tasks: [
        {
          text: "Count to five in A9.",
          range: "A9",
          solution: "=SEQUENCE(5)",
          rules: [rule({ kind: "between", range: "A9:A10", lower: 1, upper: 5 })],
        },
      ],
    });
    expect(gradeAssignment(mod, { A9: "=SEQUENCE(5)" }).earned).toBe(0);
    expect(gradeAssignment(mod, { A9: "=SEQUENCE(2)" }).earned).toBe(1);
  });
});

describe("equals", () => {
  it("compares exactly, as value_correct does", () => {
    const upper = [rule({ kind: "equals", cell: "B1", value: "PEAR" })];
    expect(earned(upper, { inputs: { ...sheet, B1: "=UPPER(A1)" } })).toBe(1);
    expect(earned(upper, { inputs: { ...sheet, B1: "=LOWER(A1)" } })).toBe(0);
  });

  it("accepts any of a list, a tolerance, spaces ignored with trim, and error codes", () => {
    expect(earned([rule({ kind: "equals", cell: "B1", value: [1, 2] })], { inputs: { B1: "2" } })).toBe(1);
    const near = [rule({ kind: "equals", cell: "B1", value: 0.33, tolerance: 0.01 })];
    expect(earned(near, { inputs: { B1: "=1/3" } })).toBe(1);
    const spaced = [rule({ kind: "equals", cell: "B1", value: "a b", trim: true })];
    expect(earned(spaced, { inputs: { B1: "ab " } })).toBe(1);
    const blocked = [rule({ kind: "equals", cell: "C1", value: "#SPILL!" })];
    expect(earned(blocked, { inputs: { ...sheet, C1: "=A1:A3", C2: "x" } })).toBe(1);
  });
});

describe("contains", () => {
  const used = [rule({ kind: "contains", range: "B1:B2", text: "textjoin(" })];

  it("checks every cell's formula for the text, ignoring case", () => {
    expect(earned(used, { inputs: { B1: '=TEXTJOIN(",",TRUE,A1:A3)', B2: '=TEXTJOIN("-",TRUE,A1)' } })).toBe(1);
    expect(earned(used, { inputs: { B1: '=TEXTJOIN(",",TRUE,A1:A3)', B2: "=A1" } })).toBe(0);
  });

  it("never passes a typed value", () => {
    expect(earned([rule({ kind: "contains", cell: "B1", text: "TEXTJOIN(" })], { inputs: { B1: "TEXTJOIN(" } })).toBe(0);
  });
});

describe("name", () => {
  it("checks a name exists, in any case", () => {
    const exists = [rule({ kind: "name", name: "Rate" })];
    expect(earned(exists, { names: { RATE: "$G$2" } })).toBe(1);
    expect(earned(exists, { names: {} })).toBe(0);
  });

  it("with a range, checks the exact name refers to it, ignoring $", () => {
    const ranged = [rule({ kind: "name", name: "Rate", range: "G2" })];
    expect(earned(ranged, { names: { Rate: "$G$2" } })).toBe(1);
    expect(earned(ranged, { names: { RATE: "$G$2" } })).toBe(0);
    expect(earned(ranged, { names: { Rate: "$G$3" } })).toBe(0);
  });
});

describe("between", () => {
  it("passes numbers within the bounds, endpoints included", () => {
    const draw = [rule({ kind: "between", range: "B1:B2", lower: 1, upper: 10 })];
    expect(earned(draw, { inputs: { B1: "1", B2: "10" } })).toBe(1);
    expect(earned(draw, { inputs: { B1: "1", B2: "11" } })).toBe(0);
    expect(earned(draw, { inputs: { B1: "1", B2: "x" } })).toBe(0);
  });

  it("takes formula bounds and can exclude the endpoints", () => {
    const strict = [rule({ kind: "between", cell: "B1", lower: "=G2", upper: 10, endpoints: "exclude" })];
    expect(earned(strict, { inputs: { G2: "1", B1: "1" } })).toBe(0);
    expect(earned(strict, { inputs: { G2: "1", B1: "2" } })).toBe(1);
  });
});

describe("selected", () => {
  const there = [rule({ kind: "selected", range: "G2" })];

  it("checks where the selection ended", () => {
    expect(earned(there, { selection: "G2" })).toBe(1);
    expect(earned(there, { selection: "A1" })).toBe(0);
    expect(earned(there, {})).toBe(0);
  });

  it("ends there in the solved workbook, so validation passes", () => {
    const mod = moduleWith(there, { range: "G2", solution: "=1" });
    expect(solvedState(mod).selection).toBe("G2");
    expect(gradeAssignment(mod, solvedState(mod)).earned).toBe(1);
  });
});

describe("rule problems", () => {
  it("reports each with its path", () => {
    expect(
      problems([
        rule({ kind: "value", cell: "B1", range: "B1:B2", formula: "=1" }),
        rule({ kind: "equals", cell: "B1", value: { no: 1 }, trim: "yes" }),
        rule({ kind: "name", name: "Q1" }),
        rule({ kind: "between", cell: "B1", lower: 1, upper: "2", endpoints: "sometimes" }),
        rule({ kind: "selected", range: "A1" }),
        rule({ kind: "selected", range: "A2" }),
        rule({ kind: "maybe" }),
      ]),
    ).toEqual([
      'tasks[0].rules[0]: give "cell" or "range", not both',
      "tasks[0].rules[1].value: must be text, a number, or true or false",
      "tasks[0].rules[1].trim: must be true or false",
      'tasks[0].rules[2].name: "Q1" looks like a cell reference, so it can\'t be a name',
      'tasks[0].rules[3].endpoints: must be "include" or "exclude"',
      'tasks[0].rules[3].upper: "2" must start with =',
      'tasks[0].rules[6].kind: must be one of "value", "reference", "copied", "equals", "contains", "name", "between", "selected"',
      'tasks[0].rules[5]: only one "selected" rule makes sense, since the selection ends in one place',
    ]);
  });
});
