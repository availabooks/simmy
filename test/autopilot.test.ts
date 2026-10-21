import { describe, expect, it } from "vitest";
import { fillWrites, type Inputs, normalizeFormula, parseRange, type Range } from "../src/engine/sheet.ts";
import { doItWrites, gradeAssignment, initialInputs } from "../src/module/grade.ts";
import { defaultAutopilot } from "../src/module/parse.ts";
import type { PilotStep, SimModule } from "../src/module/types.ts";
import { loadFixture } from "./fixture.ts";

const mod = loadFixture();
const TASKS = mod.tasks;

/** Play a plan back on the sheet without the UI. */
function play(m: SimModule, plan: PilotStep[]): Inputs {
  let inputs: Inputs = initialInputs(m);
  for (const step of plan) {
    switch (step.kind) {
      case "hold":
      case "open":
      case "submit":
        break;
      case "doit": {
        const task = m.tasks.find((t) => t.id === step.taskId);
        inputs = { ...inputs, ...doItWrites(task as (typeof m.tasks)[number]) };
        break;
      }
      case "fill":
        inputs = {
          ...inputs,
          ...fillWrites(inputs, parseRange(step.cell) as Range, parseRange(step.fill) as Range),
        };
        break;
      case "type":
        inputs = { ...inputs, [step.cell]: normalizeFormula(step.formula) };
        if (step.fill) {
          inputs = {
            ...inputs,
            ...fillWrites(inputs, parseRange(step.cell) as Range, parseRange(step.fill) as Range),
          };
        }
        break;
      default: {
        const unreachable: never = step;
        return unreachable;
      }
    }
  }
  return inputs;
}

const AUTOPILOT_PLAN = mod.autopilot;

describe("autopilot plan", () => {
  it("covers every task", () => {
    const acted = AUTOPILOT_PLAN.flatMap((step) =>
      step.kind === "type" || step.kind === "fill" || step.kind === "doit" ? [step.taskId] : [],
    );
    expect(new Set(acted)).toEqual(new Set(TASKS.map((task) => task.id)));
  });

  it("earns full credit when played back on the sheet", () => {
    const report = gradeAssignment(mod, play(mod, AUTOPILOT_PLAN));
    expect(report.tasks.flatMap((t) => t.deductions)).toEqual([]);
    expect(report.earned).toBe(mod.totalPoints);
  });

  it("earns full credit with the plan made from the tasks alone", () => {
    const report = gradeAssignment(mod, play(mod, defaultAutopilot(TASKS)));
    expect(report.earned).toBe(mod.totalPoints);
  });

  it("fills only to the cells each task's range asks for", () => {
    for (const task of TASKS) {
      const covered = AUTOPILOT_PLAN.flatMap((step) => {
        if (step.kind === "fill" && step.taskId === task.id) {
          return [step.fill];
        }
        if (step.kind === "type" && step.taskId === task.id) {
          return [step.fill ?? step.cell];
        }
        return [];
      });
      expect(covered).toContain(task.range);
    }
  });
});
