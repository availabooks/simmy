/**
 * The workbook a student works on: the cells, the defined names, and what picks RANDBETWEEN's
 * numbers. The grid, Do it and grading all read and change this one state.
 */
import type { Inputs, Names } from "./sheet.ts";

export type WorkbookState = {
  inputs: Inputs;
  names: Names;
  /**
   * Picks RANDBETWEEN's numbers. It changes with every change to the cells (and with
   * Calculate Now), the way Excel recalculates, so the same state always shows the same numbers.
   */
  seed: number;
  /** The selected range, "B2" or "B2:C5", when grading needs to know where the student went. */
  selection?: string;
};

/** A change to the workbook. A cell set to "" is cleared; a name set to null is deleted. */
export type Change = { cells?: Inputs; names?: Record<string, string | null> };

export function applyChange(state: WorkbookState, change: Change): WorkbookState {
  let { inputs, names, seed } = state;
  if (change.cells && Object.keys(change.cells).length > 0) {
    inputs = { ...inputs };
    for (const [key, raw] of Object.entries(change.cells)) {
      if (raw === "") {
        delete inputs[key];
      } else {
        inputs[key] = raw;
      }
    }
    seed = nextSeed(seed);
  }
  if (change.names) {
    names = { ...names };
    for (const [name, ref] of Object.entries(change.names)) {
      // Names match without regard to case, so redefining one replaces it.
      for (const key of Object.keys(names)) {
        if (key.toUpperCase() === name.toUpperCase()) {
          delete names[key];
        }
      }
      if (ref !== null) {
        names[name] = ref;
      }
    }
  }
  return { ...state, inputs, names, seed };
}

/** The seed after a recalculation. */
export function nextSeed(seed: number): number {
  return (Math.imul(seed, 1664525) + 1013904223) >>> 0;
}

/** A workbook from what the grading functions are given: a full state, or just the cells. */
export function toState(book: WorkbookState | Inputs): WorkbookState {
  if (isState(book)) {
    return book;
  }
  return { inputs: book, names: {}, seed: 1 };
}

function isState(book: WorkbookState | Inputs): book is WorkbookState {
  return typeof book.inputs === "object" && typeof (book as WorkbookState).names === "object";
}
