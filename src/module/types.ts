/**
 * A simulator module: one workbook, one sheet, and the Excel Educator task list
 * that grades it. Hosts write it as YAML (see hosted-modules/); `parseModule`
 * checks it and resolves it into the `SimModule` the UI reads.
 */
import type { Inputs, Names, NumberFormat, Range } from "../engine/sheet.ts";

export type CellStyle = {
  fill?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  align?: "center";
  format?: NumberFormat;
  /** Thin gray borders on the right and bottom edges. */
  lined?: boolean;
};

export type Outline = { range: Range; weight: 1 | 2; color: string };

/** A literal value a rule compares with. */
export type Scalar = string | number | boolean;

/** One rubric rule, mirroring one of the add-in's grading steps (named after each kind). */
export type RubricRule =
  /**
   * values_match: the cell equals what the solution formula gives on the student's own sheet.
   * With `range`, every cell of it matches the formula's array, for a result that spills;
   * `cell` is then the range's top-left.
   */
  | { kind: "value"; points: number; message: string; cell: string; range?: string; formula: string }
  /** find_references: the formula refers to `ref`; with `forms`, only those $ spellings count. */
  | { kind: "reference"; points: number; message: string; cell: string; ref: string; forms?: string[] }
  /** cells_copied: every cell holds the same formula as the first, copied down or across. */
  | { kind: "copied"; points: number; message: string; range: string }
  /**
   * value_correct and value_correct_tolerance: the cell holds this value exactly, or one of a
   * list; within `tolerance` for a number; ignoring all spaces with `trim`.
   */
  | {
      kind: "equals";
      points: number;
      message: string;
      cell: string;
      value: Scalar | Scalar[];
      tolerance?: number;
      trim?: boolean;
    }
  /** formula_contains: every cell of the range holds a formula that includes `text`, ignoring case. */
  | { kind: "contains"; points: number; message: string; range: string; text: string }
  /** check_name_exists, or check_named_range with `range`: the name is defined, for that range. */
  | { kind: "name"; points: number; message: string; name: string; range?: string }
  /**
   * value_between: every cell of the range is a number from `lower` to `upper` (each a number
   * or a formula), including both unless `endpoints` is "exclude".
   */
  | {
      kind: "between";
      points: number;
      message: string;
      range: string;
      lower: number | string;
      upper: number | string;
      endpoints: "include" | "exclude";
    }
  /** activecell: the student ended on this range, by Go To, the Name Box or by clicking. */
  | { kind: "selected"; points: number; message: string; range: string };

export type SimTask = {
  id: string;
  name: string;
  text: string;
  details: string[];
  /** The cell the add-in selects when the task opens. */
  hint: string;
  points: number;
  /** Where Do it writes; given with `solution`, or left out for a task that only defines names. */
  range?: string;
  /** The formula Do it writes into the top-left cell; the rest are shifted copies. */
  solution?: string;
  /** Names Do it defines (and Undo it deletes), as absolute ranges. */
  names: Names;
  rules: RubricRule[];
};

/** One narration clip and the line spoken in it. The clip is loaded when this entry starts. */
export type PilotLine = { audio: string; transcript: string };

export type PilotStep =
  | ({ taskId: string; kind: "type"; cell: string; formula: string; fill?: string; finish?: false } & Partial<PilotLine>)
  | ({ taskId: string; kind: "fill"; cell: string; fill: string; finish?: false } & Partial<PilotLine>)
  | ({ taskId: string; kind: "open" } & Partial<PilotLine>)
  | ({ taskId: string; kind: "doit"; showUndo?: boolean; finish?: false } & Partial<PilotLine>)
  | ({ kind: "hold"; cell?: string } & PilotLine)
  | ({ kind: "submit" } & PilotLine);

export type UnsupportedTool = "video" | "book";

export type NoteMessage = { title: string; text: string };

export type SimCopy = {
  /** The add-in's name in the task pane title bar. */
  paneTitle: string;
  unsupported: Record<UnsupportedTool, NoteMessage>;
  submitSteps: string[];
};

export type SimModule = {
  /** `names` are defined from the start, as absolute ranges. */
  workbook: { name: string; sheet: string; names: Names };
  assignment: { title: string; text: string; section: string };
  sheet: {
    rows: number;
    /** Width in pixels of every column, A first. */
    widths: number[];
    /** The workbook as it opens, before the student types anything. */
    cells: Inputs;
    /** Resolved formatting, cell by cell. */
    styles: Record<string, CellStyle>;
    outlines: Outline[];
    start: string;
  };
  tasks: SimTask[];
  totalPoints: number;
  note: { title: string; steps: string[] };
  autopilot: PilotStep[];
  /** The still a phone sees before opening the simulator. */
  poster: Range;
  copy: SimCopy;
};
