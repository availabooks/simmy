/**
 * A simulator module: one workbook, one sheet, and the Excel Educator task list
 * that grades it. Hosts write it as YAML (see modules/*.yaml); `parseModule`
 * checks it and resolves it into the `SimModule` the UI reads.
 */
import type { Inputs, NumberFormat, Range } from "../engine/sheet.ts";

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

/** One rubric rule, as the add-in's grader defines it. */
export type RubricRule =
  /** values_match: the cell equals what the solution formula gives on the student's own sheet. */
  | { kind: "value"; points: number; message: string; cell: string; formula: string }
  /** find_references: the formula refers to `ref`; with `forms`, only those $ spellings count. */
  | { kind: "reference"; points: number; message: string; cell: string; ref: string; forms?: string[] }
  /** cells_copied: every cell holds the same formula as the first, copied down or across. */
  | { kind: "copied"; points: number; message: string; range: string };

export type SimTask = {
  id: string;
  name: string;
  text: string;
  details: string[];
  /** The cell the add-in selects when the task opens. */
  hint: string;
  points: number;
  /** Where Do it writes. */
  range: string;
  /** The formula Do it writes into the top-left cell; the rest are shifted copies. */
  solution: string;
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
  workbook: { name: string; sheet: string };
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
