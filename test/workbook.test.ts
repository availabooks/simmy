import { describe, expect, it } from "vitest";
import { applyChange, toState, type WorkbookState } from "../src/engine/workbook.ts";
import { gradeTask } from "../src/module/grade.ts";
import type { SimTask } from "../src/module/types.ts";

const empty: WorkbookState = { inputs: {}, names: {}, seed: 1 };

describe("applyChange", () => {
  it("writes and clears cells, and recalculates", () => {
    const state = applyChange({ ...empty, inputs: { A1: "1", A2: "2" } }, { cells: { A1: "", B1: "=A2" } });
    expect(state.inputs).toEqual({ A2: "2", B1: "=A2" });
    expect(state.seed).not.toBe(1);
  });

  it("defines, redefines and deletes names without regard to case", () => {
    let state = applyChange(empty, { names: { Rate: "$G$2" } });
    state = applyChange(state, { names: { RATE: "$G$3" } });
    expect(state.names).toEqual({ RATE: "$G$3" });
    state = applyChange(state, { names: { rate: null } });
    expect(state.names).toEqual({});
  });

  it("doesn't recalculate for a change to names alone", () => {
    expect(applyChange(empty, { names: { Rate: "$G$2" } }).seed).toBe(1);
  });

  it("leaves the state it was given alone", () => {
    const before = { ...empty, inputs: { A1: "1" } };
    applyChange(before, { cells: { A1: "2" }, names: { X: "$A$1" } });
    expect(before).toEqual({ ...empty, inputs: { A1: "1" } });
  });
});

describe("grading a workbook", () => {
  const task: SimTask = {
    id: "A",
    name: "A",
    text: "",
    details: [],
    hint: "B2",
    points: 2,
    range: "B2",
    solution: "=A2*Rate",
    names: {},
    rules: [
      { kind: "value", cell: "B2", formula: "=A2*G2", points: 1, message: "value" },
      { kind: "reference", cell: "B2", ref: "G2", forms: ["$G$2"], points: 1, message: "reference" },
    ],
  };
  const inputs = { A2: "10", G2: "0.5", B2: "=A2*Rate" };

  it("follows names, so a name counts as the cell it refers to", () => {
    const grade = gradeTask(task, { inputs, names: { Rate: "$G$2" }, seed: 1 });
    expect(grade.earned).toBe(2);
  });

  it("treats plain cells as a workbook without names", () => {
    expect(toState(inputs)).toEqual({ inputs, names: {}, seed: 1 });
    expect(gradeTask(task, inputs).earned).toBe(0);
  });
});

describe("Do it and Undo it with names", () => {
  it("defines the task's names and deletes them again", async () => {
    const { parseModule } = await import("../src/module/parse.ts");
    const { doIt, initialState, undoIt } = await import("../src/module/grade.ts");
    const mod = parseModule({
      workbook: { name: "Book1", names: { Given: "A1" } },
      assignment: { title: "Names", text: "Name things." },
      sheet: { cells: { A1: 2, A2: 3 } },
      tasks: [{ text: "Name A2 Rate.", names: { Rate: "A2" }, rules: [] }],
    });
    const start = initialState(mod);
    expect(start.names).toEqual({ Given: "$A$1" });
    const done = applyChange(start, doIt(mod.tasks[0]));
    expect(done.names).toEqual({ Given: "$A$1", Rate: "$A$2" });
    expect(applyChange(done, undoIt(mod.tasks[0])).names).toEqual({ Given: "$A$1" });
  });
});
