/**
 * What Do it writes, and a local version of the add-in's rubric. Grading is
 * generic: every check comes from the module's rules.
 */
import {
  addrName,
  evaluateFormulaAt,
  evaluateSheet,
  formulaReferences,
  inRange,
  type Inputs,
  isFormula,
  normalizeFormula,
  parseAddr,
  parseRange,
  type Range,
  rangeAddrs,
  shiftFormula,
  type Value,
} from "../engine/sheet.ts";
import type { RubricRule, SimModule, SimTask } from "./types.ts";

/** The workbook as the module builds it, before the student types anything. */
export function initialInputs(mod: SimModule): Inputs {
  return { ...mod.sheet.cells };
}

/** Cell writes for Do it: the solution in every cell of the task range, shifted like a fill. */
export function doItWrites(task: SimTask): Inputs {
  const range = parseRange(task.range) as Range;
  const writes: Inputs = {};
  for (const addr of rangeAddrs(range)) {
    writes[addrName(addr)] = shiftFormula(task.solution, addr.col - range.c1, addr.row - range.r1);
  }
  return writes;
}

/** Cell writes for Undo it: clear the task range. */
export function undoItWrites(task: SimTask): Inputs {
  const range = parseRange(task.range) as Range;
  const writes: Inputs = {};
  for (const addr of rangeAddrs(range)) {
    writes[addrName(addr)] = "";
  }
  return writes;
}

export type Deduction = { points: number; message: string };

export type TaskGrade = {
  taskId: string;
  earned: number;
  points: number;
  /** The rules the work missed, in the rubric's own wording. */
  deductions: Deduction[];
};

export type GradeReport = { tasks: TaskGrade[]; earned: number; points: number };

/** A cell well away from the sheet, where the grader works out expected values. */
const SCRATCH = { col: 25, row: 98 };

function sameValue(actual: Value, expected: Value): boolean {
  if (typeof expected === "number" && typeof actual === "number") {
    return Math.abs(actual - expected) <= 1e-9 * Math.max(1, Math.abs(expected));
  }
  if (typeof expected === "string" || typeof expected === "boolean") {
    return actual === expected;
  }
  return false;
}

function passes(rule: RubricRule, inputs: Inputs, values: Record<string, Value>): boolean {
  switch (rule.kind) {
    case "value":
      return sameValue(values[rule.cell] ?? null, evaluateFormulaAt(inputs, rule.formula, SCRATCH));
    case "reference": {
      const raw = inputs[rule.cell];
      if (!isFormula(raw)) {
        return false;
      }
      const refs = formulaReferences(raw);
      if (rule.forms) {
        const forms = rule.forms;
        return refs.some((ref) => forms.includes(ref.text));
      }
      const target = parseAddr(rule.ref);
      return target !== null && refs.some((ref) => inRange(ref.range, target));
    }
    case "copied": {
      const range = parseRange(rule.range) as Range;
      const first = inputs[addrName({ col: range.c1, row: range.r1 })];
      if (!isFormula(first)) {
        return false;
      }
      const base = normalizeFormula(first);
      return rangeAddrs(range).every((addr) => {
        const raw = inputs[addrName(addr)];
        return (
          isFormula(raw) &&
          normalizeFormula(raw) === shiftFormula(base, addr.col - range.c1, addr.row - range.r1)
        );
      });
    }
    default: {
      const _exhaustive: never = rule;
      return _exhaustive;
    }
  }
}

export function gradeTask(task: SimTask, inputs: Inputs, values = evaluateSheet(inputs)): TaskGrade {
  const deductions = task.rules
    .filter((rule) => !passes(rule, inputs, values))
    .map((rule) => ({ points: rule.points, message: rule.message }));
  const lost = deductions.reduce((sum, d) => sum + d.points, 0);
  return { taskId: task.id, points: task.points, earned: Math.max(0, task.points - lost), deductions };
}

export function gradeAssignment(mod: SimModule, inputs: Inputs): GradeReport {
  const values = evaluateSheet(inputs);
  const tasks = mod.tasks.map((task) => gradeTask(task, inputs, values));
  return {
    tasks,
    earned: tasks.reduce((sum, grade) => sum + grade.earned, 0),
    points: mod.totalPoints,
  };
}

/** Every task done with Do it: the sheet a perfect student hands in. */
export function solvedInputs(mod: SimModule): Inputs {
  return mod.tasks.reduce<Inputs>((inputs, task) => ({ ...inputs, ...doItWrites(task) }), initialInputs(mod));
}
