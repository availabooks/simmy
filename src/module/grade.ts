/**
 * What Do it writes, and a local version of the add-in's rubric. Grading is
 * generic: every check comes from the module's rules.
 */
import { lookupName, nameRange } from "../engine/names.ts";
import {
  addrName,
  createEvaluator,
  type EvalOptions,
  formulaReferences,
  inRange,
  type Inputs,
  isError,
  isFormula,
  type Matrix,
  normalizeFormula,
  parseAddr,
  parseRange,
  type Range,
  rangeAddrs,
  rangeName,
  shiftFormula,
  type Value,
} from "../engine/sheet.ts";
import { applyChange, type Change, toState, type WorkbookState } from "../engine/workbook.ts";
import type { RubricRule, Scalar, SimModule, SimTask } from "./types.ts";

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

/**
 * value_correct: exactly the value, as the add-in compares it. A number may be off by up to
 * `tolerance`; with `trim`, text ignores every space. An error matches its code ("#N/A").
 */
function equalsValue(actual: Value, expected: Scalar, tolerance?: number, trim?: boolean): boolean {
  if (isError(actual)) {
    return actual.error === expected;
  }
  if (typeof actual === "number" && typeof expected === "number" && tolerance !== undefined) {
    return Math.abs(actual - expected) <= tolerance;
  }
  if (trim && typeof actual === "string" && typeof expected === "string") {
    return actual.replace(/\s+/g, "") === expected.replace(/\s+/g, "");
  }
  return actual === expected;
}

function sameValue(actual: Value, expected: Value): boolean {
  if (typeof expected === "number" && typeof actual === "number") {
    return Math.abs(actual - expected) <= 1e-9 * Math.max(1, Math.abs(expected));
  }
  if (typeof expected === "string" || typeof expected === "boolean") {
    return actual === expected;
  }
  return false;
}

/**
 * A workbook ready to grade: its cells' values, and the expected values of rule formulas,
 * worked out in the same pass over the sheet.
 */
type Grader = {
  book: WorkbookState;
  values: Record<string, Value>;
  evaluate: (formula: string) => Value | Matrix;
};

function grader(book: WorkbookState, opts: EvalOptions): Grader {
  const evaluator = createEvaluator(book.inputs, opts);
  return { book, values: evaluator.book().values, evaluate: evaluator.expression };
}

/** A formula's value as a cell would show it: an array's top-left value. */
const firstValue = (result: Value | Matrix): Value => (Array.isArray(result) ? result[0][0] : result);

function passes(rule: RubricRule, { book, values, evaluate }: Grader): boolean {
  const { inputs, names } = book;
  switch (rule.kind) {
    case "value": {
      if (!rule.range) {
        return sameValue(values[rule.cell] ?? null, firstValue(evaluate(rule.formula)));
      }
      // A spilled result: the formula's array must fill the range exactly.
      const range = parseRange(rule.range) as Range;
      const result = evaluate(rule.formula);
      const expected = Array.isArray(result) ? result : [[result]];
      if (expected.length !== range.r2 - range.r1 + 1 || expected[0].length !== range.c2 - range.c1 + 1) {
        return false;
      }
      return rangeAddrs(range).every((addr) =>
        sameValue(values[addrName(addr)] ?? null, expected[addr.row - range.r1][addr.col - range.c1]),
      );
    }
    case "equals": {
      const actual = values[rule.cell] ?? null;
      const accepted = Array.isArray(rule.value) ? rule.value : [rule.value];
      return accepted.some((value) => equalsValue(actual, value, rule.tolerance, rule.trim));
    }
    case "contains": {
      // As the add-in does, only a formula counts: a typed value never contains the text.
      const text = rule.text.toUpperCase();
      return rangeAddrs(parseRange(rule.range) as Range).every((addr) => {
        const raw = inputs[addrName(addr)];
        return isFormula(raw) && raw.toUpperCase().includes(text);
      });
    }
    case "name": {
      if (!rule.range) {
        return lookupName(names, rule.name) !== null;
      }
      // With a range the add-in matches the name's case exactly, and ignores $ signs.
      const ref = names[rule.name];
      const named = ref === undefined ? null : nameRange(ref);
      return named !== null && rangeName(named) === rangeName(parseRange(rule.range) as Range);
    }
    case "between": {
      const bound = (value: number | string): Value =>
        typeof value === "number" ? value : firstValue(evaluate(value));
      const lower = bound(rule.lower);
      const upper = bound(rule.upper);
      if (typeof lower !== "number" || typeof upper !== "number") {
        return false;
      }
      return rangeAddrs(parseRange(rule.range) as Range).every((addr) => {
        const value = values[addrName(addr)];
        if (typeof value !== "number") {
          return false;
        }
        return rule.endpoints === "exclude" ? value > lower && value < upper : value >= lower && value <= upper;
      });
    }
    case "selected": {
      const at = book.selection ? parseRange(book.selection) : null;
      return at !== null && rangeName(at) === rangeName(parseRange(rule.range) as Range);
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

/**
 * How a module's workbook is evaluated, by the grid and the grader alike: with its names, the
 * numbers its seed picks, and the sheet's size, so an array that runs off the sheet is #SPILL!
 * in both.
 */
export function evalOptions(mod: SimModule, book: WorkbookState): EvalOptions {
  return {
    bounds: { rows: mod.sheet.rows, cols: mod.sheet.widths.length },
    names: book.names,
    seed: book.seed,
  };
}

function gradeWith(task: SimTask, graded: Grader): TaskGrade {
  const deductions = task.rules
    .filter((rule) => !passes(rule, graded))
    .map((rule) => ({ points: rule.points, message: rule.message }));
  const lost = deductions.reduce((sum, d) => sum + d.points, 0);
  return { taskId: task.id, points: task.points, earned: Math.max(0, task.points - lost), deductions };
}

/**
 * Grade one task. `book` is the workbook, or just its cells. Without the module the sheet's
 * size isn't known, so an array is graded as if the sheet had room for it; gradeAssignment
 * knows the size. `values` are ignored: they're worked out alongside the rules' formulas.
 */
export function gradeTask(task: SimTask, book: WorkbookState | Inputs, _values?: Record<string, Value>): TaskGrade {
  const state = toState(book);
  return gradeWith(task, grader(state, { names: state.names, seed: state.seed }));
}

/** Grade every task. `book` is the workbook, or just its cells. */
export function gradeAssignment(mod: SimModule, book: WorkbookState | Inputs): GradeReport {
  const state = toState(book);
  const graded = grader(state, evalOptions(mod, state));
  const tasks = mod.tasks.map((task) => gradeWith(task, graded));
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

/**
 * Every task done with Do it: the workbook a perfect student hands in. Where a rule asks for
 * the selection to end somewhere, it ends there.
 */
export function solvedState(mod: SimModule): WorkbookState {
  const solved = mod.tasks.reduce((state, task) => applyChange(state, doIt(task)), initialState(mod));
  const selected = mod.tasks.flatMap((task) => task.rules).find((rule) => rule.kind === "selected");
  return selected ? { ...solved, selection: selected.range } : solved;
}
