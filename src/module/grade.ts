/**
 * What Do it writes, and a local version of the add-in's rubric. Grading is
 * generic: every check comes from the module's rules.
 */
import {
  addrName,
  evaluateFormulaAt,
  evaluateSheet,
  type EvalOptions,
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
import { applyChange, type Change, toState, type WorkbookState } from "../engine/workbook.ts";
import type { RubricRule, SimModule, SimTask } from "./types.ts";

/** The cells as the module builds them, before the student types anything. */
export function initialInputs(mod: SimModule): Inputs {
  return { ...mod.sheet.cells };
}

/** The workbook as the module builds it, before the student does anything. */
export function initialState(mod: SimModule): WorkbookState {
  return { inputs: initialInputs(mod), names: { ...mod.workbook.names }, seed: 1 };
}

/** Do it as a change to the workbook: the solution written across the range, and the task's names. */
export function doIt(task: SimTask): Change {
  return { cells: doItWrites(task), names: task.names };
}

/** Undo it as a change to the workbook: the range cleared, and the task's names deleted. */
export function undoIt(task: SimTask): Change {
  return {
    cells: undoItWrites(task),
    names: Object.fromEntries(Object.keys(task.names).map((name) => [name, null])),
  };
}

/** Cell writes for Do it: the solution in every cell of the task range, shifted like a fill. */
export function doItWrites(task: SimTask): Inputs {
  const range = task.range ? parseRange(task.range) : null;
  const writes: Inputs = {};
  if (!range || !task.solution) {
    return writes;
  }
  for (const addr of rangeAddrs(range)) {
    writes[addrName(addr)] = shiftFormula(task.solution, addr.col - range.c1, addr.row - range.r1);
  }
  return writes;
}

/** Cell writes for Undo it: clear the task range. */
export function undoItWrites(task: SimTask): Inputs {
  const range = task.range ? parseRange(task.range) : null;
  const writes: Inputs = {};
  if (!range) {
    return writes;
  }
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

function passes(rule: RubricRule, book: WorkbookState, values: Record<string, Value>): boolean {
  const { inputs, names } = book;
  switch (rule.kind) {
    case "value": {
      const expected = evaluateFormulaAt(inputs, rule.formula, SCRATCH, options(book));
      return sameValue(values[rule.cell] ?? null, expected);
    }
    case "reference": {
      const raw = inputs[rule.cell];
      if (!isFormula(raw)) {
        return false;
      }
      const refs = formulaReferences(raw, names);
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
      const base = normalizeFormula(first, names);
      return rangeAddrs(range).every((addr) => {
        const raw = inputs[addrName(addr)];
        return (
          isFormula(raw) &&
          normalizeFormula(raw, names) === shiftFormula(base, addr.col - range.c1, addr.row - range.r1)
        );
      });
    }
    default: {
      const _exhaustive: never = rule;
      return _exhaustive;
    }
  }
}

/** How the grader evaluates a workbook: with its names, and the numbers its seed picks. */
function options(book: WorkbookState): EvalOptions {
  return { names: book.names, seed: book.seed };
}

/** Grade one task. `book` is the workbook, or just its cells. */
export function gradeTask(task: SimTask, book: WorkbookState | Inputs, values?: Record<string, Value>): TaskGrade {
  const state = toState(book);
  const cells = values ?? evaluateSheet(state.inputs, options(state));
  const deductions = task.rules
    .filter((rule) => !passes(rule, state, cells))
    .map((rule) => ({ points: rule.points, message: rule.message }));
  const lost = deductions.reduce((sum, d) => sum + d.points, 0);
  return { taskId: task.id, points: task.points, earned: Math.max(0, task.points - lost), deductions };
}

/** Grade every task. `book` is the workbook, or just its cells. */
export function gradeAssignment(mod: SimModule, book: WorkbookState | Inputs): GradeReport {
  const state = toState(book);
  const values = evaluateSheet(state.inputs, options(state));
  const tasks = mod.tasks.map((task) => gradeTask(task, state, values));
  return {
    tasks,
    earned: tasks.reduce((sum, grade) => sum + grade.earned, 0),
    points: mod.totalPoints,
  };
}

/** Every task done with Do it: the cells a perfect student hands in. */
export function solvedInputs(mod: SimModule): Inputs {
  return solvedState(mod).inputs;
}

/** Every task done with Do it: the workbook a perfect student hands in. */
export function solvedState(mod: SimModule): WorkbookState {
  return mod.tasks.reduce((state, task) => applyChange(state, doIt(task)), initialState(mod));
}
