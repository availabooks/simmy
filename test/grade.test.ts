import { describe, expect, it } from "vitest";
import { evaluateSheet, type Inputs } from "../src/engine/sheet.ts";
import {
  doItWrites,
  gradeAssignment,
  initialInputs as initialInputsOf,
  undoItWrites,
} from "../src/module/grade.ts";
import { loadFixture } from "./fixture.ts";

const mod = loadFixture();
const TASKS = mod.tasks;
const TOTAL_POINTS = mod.totalPoints;
const initialInputs = () => initialInputsOf(mod);

const byId = (id: string) => {
  const task = TASKS.find((t) => t.id === id);
  if (!task) {
    throw new Error(id);
  }
  return task;
};

const apply = (inputs: Inputs, writes: Inputs) => {
  const next = { ...inputs };
  for (const [key, raw] of Object.entries(writes)) {
    if (raw === "") {
      delete next[key];
    } else {
      next[key] = raw;
    }
  }
  return next;
};

const grade = (inputs: Inputs, id: string) => gradeAssignment(mod, inputs).tasks.find((t) => t.taskId === id);

describe("References and Calculations Practice", () => {
  it("builds the pool sale sheet the add-in sets up", () => {
    const inputs = initialInputs();
    expect(inputs.B2).toBe("Discount");
    expect(inputs.C2).toBe("0.3");
    expect(inputs.B6).toBe("Pool");
    expect(inputs.C12).toBe("13.25");
    expect(inputs.E14).toBe("=SUM(E6:E12)");
    expect(TOTAL_POINTS).toBe(100);
  });

  it("weights every rule the way the rubric does", () => {
    for (const task of TASKS) {
      expect(task.rules.reduce((sum, rule) => sum + rule.points, 0)).toBe(task.points);
    }
  });

  it("does a task the way the add-in does: shifted formulas down the range", () => {
    expect(doItWrites(byId("A"))).toMatchObject({ D6: "=C6*$C$2", D12: "=C12*$C$2" });
    expect(doItWrites(byId("B")).E9).toBe("=C9-D9");
    expect(doItWrites(byId("C"))).toEqual({ E15: "=E14*C3" });
    expect(Object.values(undoItWrites(byId("A")))).toEqual(["", "", "", "", "", "", ""]);
  });

  it("gives full credit when every task is done for you", () => {
    let inputs = initialInputs();
    for (const task of TASKS) {
      inputs = apply(inputs, doItWrites(task));
    }
    const report = gradeAssignment(mod, inputs);
    expect(report.tasks.flatMap((t) => t.deductions)).toEqual([]);
    expect(report.earned).toBe(100);
    const values = evaluateSheet(inputs);
    expect(values.D6).toBeCloseTo(105.297, 6);
    expect(values.E16).toBeCloseTo(405.9493, 3);
  });

  it("takes points for a relative reference that should be absolute", () => {
    // =C6*C2 is right in D6 but breaks when copied down, so D7:D12 come out wrong.
    const inputs: Inputs = { ...initialInputs() };
    for (let row = 6; row <= 12; row++) {
      inputs[`D${row}`] = `=C${row}*C${row - 4}`;
    }
    const result = grade(inputs, "A");
    expect(result?.deductions.map((d) => d.message)).toEqual([
      "The formula in cell D6 does not contain an absolute reference to cell C2",
    ]);
    expect(result?.earned).toBe(20);
  });

  it("takes points when a formula is typed over instead of copied", () => {
    const inputs: Inputs = { ...initialInputs(), E6: "=C6-D6", E7: "50" };
    expect(grade(inputs, "B")?.deductions).toEqual([
      { points: 8, message: "The formulas in range E6:E12 are not reused" },
    ]);
  });

  it("grades a blank task as zero with every rule listed", () => {
    const result = grade(initialInputs(), "D");
    expect(result?.earned).toBe(0);
    expect(result?.deductions.map((d) => d.points)).toEqual([15, 5, 5]);
  });

  it("checks values against the student's own earlier answers", () => {
    // A wrong subtotal still lets a correct tax formula earn its points.
    const inputs: Inputs = { ...initialInputs(), E15: "=E14*C3" };
    expect(grade(inputs, "C")?.earned).toBe(25);
  });
});

describe("Do it and Undo it, task by task", () => {
  const clean = initialInputs();
  for (const task of TASKS) {
    it(`task ${task.id} (${task.name}) fills its range, grades clean, and undoes to the original sheet`, () => {
      let done = clean;
      for (const earlier of TASKS.slice(0, TASKS.indexOf(task))) {
        done = apply(done, doItWrites(earlier));
      }
      const before = done;
      done = apply(done, doItWrites(task));
      const values = evaluateSheet(done);
      for (const key of Object.keys(doItWrites(task))) {
        expect(typeof values[key]).toBe("number");
      }
      expect(grade(done, task.id)?.earned).toBe(task.points);
      expect(apply(done, undoItWrites(task))).toEqual(before);
    });
  }
});
