/**
 * The workbook a student works on: the cells, the defined names, and what picks RANDBETWEEN's
 * numbers. The grid, Do it and grading all read and change this one state.
 */
import type { Inputs, Names } from "./sheet.ts";

export type WorkbookState = {
  inputs: Inputs;
  names: Names;
  /** Comments on defined names (up to 255 characters), shown in the Name Manager. */
  nameComments?: Record<string, string>;
  /**
   * Picks RANDBETWEEN's numbers. It changes with every change to the cells (and with
   * Calculate Now), the way Excel recalculates, so the same state always shows the same numbers.
   */
  seed: number;
  /** The selected range, "B2" or "B2:C5", when grading needs to know where the student went. */
  selection?: string;
};

/**
 * A change to the workbook. A cell set to "" is cleared; a name set to null is deleted, along
 * with its comment. A comment set to "" is removed.
 */
export type Change = {
  cells?: Inputs;
  names?: Record<string, string | null>;
  nameComments?: Record<string, string>;
};

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
  let nameComments = state.nameComments;
  // Names match without regard to case, so defining one replaces any spelled differently.
  const without = (map: Record<string, string>, name: string) => {
    for (const key of Object.keys(map)) {
      if (key.toUpperCase() === name.toUpperCase()) {
        delete map[key];
      }
    }
  };
  if (change.names) {
    names = { ...names };
    nameComments = { ...nameComments };
    for (const [name, ref] of Object.entries(change.names)) {
      const upper = name.toUpperCase();
      const comment = Object.entries(nameComments).find(([key]) => key.toUpperCase() === upper);
      without(names, name);
      without(nameComments, name);
      if (ref !== null) {
        names[name] = ref;
        if (comment) {
          nameComments[name] = comment[1];
        }
      }
    }
  }
  if (change.nameComments) {
    nameComments = { ...nameComments };
    for (const [name, comment] of Object.entries(change.nameComments)) {
      without(nameComments, name);
      if (comment !== "") {
        nameComments[name] = comment;
      }
    }
  }
  return { ...state, inputs, names, nameComments, seed };
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
