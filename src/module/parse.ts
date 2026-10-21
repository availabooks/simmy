/**
 * Turns a module as hosts write it (YAML or plain objects) into a checked
 * `SimModule`. Every problem is collected and reported together, with the path
 * to the offending field, so an author can fix a file in one pass.
 */
import {
  addrName,
  checkFormula,
  colIndex,
  colName,
  type Inputs,
  normalizeRange,
  type NumberFormat,
  parseAddr,
  parseRange,
  type Range,
  rangeAddrs,
} from "../engine/sheet.ts";
import type { CellStyle, Outline, PilotStep, RubricRule, SimCopy, SimModule, SimTask } from "./types.ts";

export class ModuleError extends Error {
  constructor(readonly problems: string[]) {
    super(`Invalid simulator module:\n${problems.map((p) => `  - ${p}`).join("\n")}`);
    this.name = "ModuleError";
  }
}

const FORMATS: NumberFormat[] = ["General", "0", "0.00", "$#,##0.00", "0.00%", "accounting"];

export const DEFAULT_COPY: SimCopy = {
  paneTitle: "Excel Educator",
  unsupported: {
    video: {
      title: "Short video solution",
      text: "In the course, this plays a short video that walks through the task. It isn't part of this demo.",
    },
    book: {
      title: "Textbook",
      text: "In the course, this opens the textbook section for the task. It isn't part of this demo.",
    },
  },
  submitSteps: ["Validate Key", "Prepare Rubric", "Check for Completeness", "Evaluate Tasks", "Finalize Score"],
};

type Raw = Record<string, unknown>;

const isObject = (value: unknown): value is Raw =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Parse YAML (or JSON, which is YAML) text into a checked module. */
export async function loadModule(text: string): Promise<SimModule> {
  const { parse } = await import("yaml");
  let data: unknown;
  try {
    data = parse(text);
  } catch (error) {
    throw new ModuleError([`YAML: ${(error as Error).message}`]);
  }
  return parseModule(data);
}

/** Check and resolve an already-parsed module. Throws `ModuleError` listing every problem. */
export function parseModule(data: unknown): SimModule {
  const problems: string[] = [];
  const fail = (path: string, message: string) => problems.push(`${path}: ${message}`);

  const root = isObject(data) ? data : (fail("module", "must be a mapping"), {});
  const obj = (value: unknown, path: string): Raw => {
    if (value === undefined) {
      return {};
    }
    if (!isObject(value)) {
      fail(path, "must be a mapping");
      return {};
    }
    return value;
  };
  const str = (value: unknown, path: string, fallback?: string): string => {
    if (value === undefined && fallback !== undefined) {
      return fallback;
    }
    if (typeof value === "number") {
      return String(value);
    }
    if (typeof value !== "string" || value.trim() === "") {
      fail(path, "must be text");
      return fallback ?? "";
    }
    return value;
  };
  const num = (value: unknown, path: string, fallback?: number): number => {
    if (value === undefined && fallback !== undefined) {
      return fallback;
    }
    if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
      fail(path, "must be a number, zero or more");
      return fallback ?? 0;
    }
    return value;
  };
  const list = (value: unknown, path: string): unknown[] => {
    if (value === undefined) {
      return [];
    }
    if (!Array.isArray(value)) {
      fail(path, "must be a list");
      return [];
    }
    return value;
  };
  const strings = (value: unknown, path: string): string[] =>
    list(value, path).map((item, i) => str(item, `${path}[${i}]`));
  const formula = (value: unknown, path: string): string => {
    const text = str(value, path);
    if (!text.startsWith("=")) {
      fail(path, `"${text}" must start with =`);
      return text;
    }
    try {
      checkFormula(text);
    } catch (error) {
      fail(path, `"${text}" is not a formula Excel would accept (${(error as Error).message})`);
    }
    return text;
  };

  if (root.version !== undefined && root.version !== 1) {
    fail("version", "only version 1 is supported");
  }

  const workbookRaw = obj(root.workbook, "workbook");
  const workbook = {
    name: str(workbookRaw.name, "workbook.name"),
    sheet: str(workbookRaw.sheet, "workbook.sheet", "Sheet1"),
  };

  const assignmentRaw = obj(root.assignment, "assignment");
  const assignment = {
    title: str(assignmentRaw.title, "assignment.title"),
    text: str(assignmentRaw.text, "assignment.text"),
    section: str(assignmentRaw.section, "assignment.section", "Assessment Section 1"),
  };

  /* ---------------------------------------------------------------- sheet */

  const sheetRaw = obj(root.sheet, "sheet");
  const rows = num(sheetRaw.rows, "sheet.rows", 40);
  const columnCount = num(sheetRaw.columns, "sheet.columns", 20);
  const columnWidth = num(sheetRaw.columnWidth, "sheet.columnWidth", 64);
  const widths = Array.from({ length: columnCount }, () => columnWidth);
  for (const [name, width] of Object.entries(obj(sheetRaw.widths, "sheet.widths"))) {
    const col = /^[A-Za-z]{1,3}$/.test(name) ? colIndex(name) : -1;
    if (col < 0 || col >= columnCount) {
      fail(`sheet.widths.${name}`, `not a column in A:${colName(columnCount - 1)}`);
      continue;
    }
    widths[col] = num(width, `sheet.widths.${name}`);
  }

  const inSheet = (range: Range) => range.c2 < columnCount && range.r2 < rows;
  const cellRef = (value: unknown, path: string): string => {
    const text = str(value, path).toUpperCase();
    const addr = parseAddr(text);
    if (!addr) {
      fail(path, `"${text}" is not a cell address like B6`);
    } else if (!inSheet(normalizeRange(addr, addr))) {
      fail(path, `${text} is outside the sheet (${columnCount} columns, ${rows} rows)`);
    }
    return text;
  };
  const rangeRef = (value: unknown, path: string): string => {
    const text = str(value, path).toUpperCase();
    const range = parseRange(text);
    if (!range) {
      fail(path, `"${text}" is not a range like D6:D12`);
    } else if (!inSheet(range)) {
      fail(path, `${text} is outside the sheet (${columnCount} columns, ${rows} rows)`);
    }
    return text;
  };
  /** "D6:E12, E14:E16" or a list of ranges. */
  const rangeList = (value: unknown, path: string): Range[] => {
    const parts = Array.isArray(value) ? value : typeof value === "string" ? value.split(/[,\s]+/) : [value];
    return parts
      .filter((part) => part !== "")
      .map((part, i) => parseRange(rangeRef(part, Array.isArray(value) ? `${path}[${i}]` : path)))
      .filter((range): range is Range => range !== null);
  };

  const cells: Inputs = {};
  for (const [key, raw] of Object.entries(obj(sheetRaw.cells, "sheet.cells"))) {
    const path = `sheet.cells.${key}`;
    const ref = cellRef(key, path);
    if (raw === null || raw === undefined || raw === "") {
      continue;
    }
    if (typeof raw !== "string" && typeof raw !== "number" && typeof raw !== "boolean") {
      fail(path, "must be text, a number or a formula");
      continue;
    }
    const text = String(raw);
    cells[ref] = text.startsWith("=") ? formula(text, path) : text;
  }

  const styles: Record<string, CellStyle> = {};
  list(sheetRaw.styles, "sheet.styles").forEach((item, i) => {
    const path = `sheet.styles[${i}]`;
    const raw = obj(item, path);
    const style: CellStyle = {};
    for (const [key, value] of Object.entries(raw)) {
      switch (key) {
        case "range":
          break;
        case "fill":
        case "color":
          style[key] = str(value, `${path}.${key}`);
          break;
        case "bold":
        case "italic":
        case "lined":
          if (typeof value !== "boolean") {
            fail(`${path}.${key}`, "must be true or false");
          }
          style[key] = value === true;
          break;
        case "align":
          if (value !== "center") {
            fail(`${path}.align`, 'only "center" is supported');
          }
          style.align = "center";
          break;
        case "format":
          if (!FORMATS.includes(value as NumberFormat)) {
            fail(`${path}.format`, `must be one of ${FORMATS.map((f) => `"${f}"`).join(", ")}`);
          }
          style.format = value as NumberFormat;
          break;
        default:
          fail(`${path}.${key}`, "is not a style this simulator knows");
      }
    }
    for (const range of rangeList(raw.range, `${path}.range`)) {
      for (const addr of rangeAddrs(range)) {
        const key = addrName(addr);
        styles[key] = { ...styles[key], ...style };
      }
    }
  });

  const outlines: Outline[] = list(sheetRaw.outlines, "sheet.outlines").flatMap((item, i) => {
    const path = `sheet.outlines[${i}]`;
    const raw = obj(item, path);
    const weight = num(raw.weight, `${path}.weight`, 1);
    if (weight !== 1 && weight !== 2) {
      fail(`${path}.weight`, "must be 1 or 2");
    }
    const color = str(raw.color, `${path}.color`, "#a6a6a6");
    return rangeList(raw.range, `${path}.range`).map((range) => ({
      range,
      weight: weight === 2 ? 2 : 1,
      color,
    }));
  });

  /* ---------------------------------------------------------------- tasks */

  const tasks: SimTask[] = list(root.tasks, "tasks").map((item, i) => {
    const path = `tasks[${i}]`;
    const raw = obj(item, path);
    const id = str(raw.id, `${path}.id`, String.fromCharCode(65 + i));
    const range = rangeRef(raw.range, `${path}.range`);
    const rules: RubricRule[] = list(raw.rules, `${path}.rules`).map((ruleItem, j) => {
      const rp = `${path}.rules[${j}]`;
      const r = obj(ruleItem, rp);
      const points = num(r.points, `${rp}.points`);
      const message = str(r.message, `${rp}.message`);
      switch (r.kind) {
        case "value":
          return { kind: "value", points, message, cell: cellRef(r.cell, `${rp}.cell`), formula: formula(r.formula, `${rp}.formula`) };
        case "reference": {
          const forms = r.forms === undefined ? undefined : strings(r.forms, `${rp}.forms`).map((f) => f.toUpperCase());
          return { kind: "reference", points, message, cell: cellRef(r.cell, `${rp}.cell`), ref: cellRef(r.ref, `${rp}.ref`), forms };
        }
        case "copied":
          return { kind: "copied", points, message, range: rangeRef(r.range, `${rp}.range`) };
        default:
          fail(`${rp}.kind`, 'must be "value", "reference" or "copied"');
          return { kind: "copied", points, message, range: "A1" };
      }
    });
    const rulePoints = rules.reduce((sum, rule) => sum + rule.points, 0);
    const parsedRange = parseRange(range);
    const topLeft = parsedRange ? addrName({ col: parsedRange.c1, row: parsedRange.r1 }) : "A1";
    return {
      id,
      name: str(raw.name, `${path}.name`, id),
      text: str(raw.text, `${path}.text`),
      details: strings(raw.details, `${path}.details`),
      hint: raw.hint === undefined ? topLeft : cellRef(raw.hint, `${path}.hint`),
      points: num(raw.points, `${path}.points`, rulePoints),
      range,
      solution: formula(raw.solution, `${path}.solution`),
      rules,
    };
  });
  if (tasks.length === 0) {
    fail("tasks", "needs at least one task");
  }
  const ids = new Set<string>();
  tasks.forEach((task, i) => {
    if (ids.has(task.id)) {
      fail(`tasks[${i}].id`, `"${task.id}" is used by another task`);
    }
    ids.add(task.id);
  });

  /* ------------------------------------------------------- note, autopilot */

  const noteRaw = obj(root.note, "note");
  const note = {
    title: str(noteRaw.title, "note.title", "Try it"),
    steps: strings(noteRaw.steps, "note.steps"),
  };

  const autopilot: PilotStep[] =
    root.autopilot === undefined
      ? defaultAutopilot(tasks)
      : list(root.autopilot, "autopilot").map((item, i) => {
          const path = `autopilot[${i}]`;
          const raw = obj(item, path);
          const taskId = str(raw.task, `${path}.task`);
          const task = tasks.find((t) => t.id === taskId);
          if (!task) {
            fail(`${path}.task`, `no task has id "${taskId}"`);
          }
          if (raw.doit !== undefined) {
            return { taskId, kind: "doit", showUndo: raw.undo === true };
          }
          const cell = raw.at === undefined ? (task?.hint ?? "A1") : cellRef(raw.at, `${path}.at`);
          const fill = raw.fill === undefined ? undefined : rangeRef(raw.fill, `${path}.fill`);
          return { taskId, kind: "type", cell, formula: formula(raw.type ?? task?.solution, `${path}.type`), fill };
        });

  /* ----------------------------------------------------------- the rest */

  const posterRaw = obj(root.poster, "poster");
  const poster =
    posterRaw.range === undefined ? defaultPoster(cells) : (parseRange(rangeRef(posterRaw.range, "poster.range")) ?? defaultPoster(cells));

  const copyRaw = obj(root.copy, "copy");
  const unsupportedRaw = obj(copyRaw.unsupported, "copy.unsupported");
  const message = (key: "video" | "book") => {
    const raw = obj(unsupportedRaw[key], `copy.unsupported.${key}`);
    return {
      title: str(raw.title, `copy.unsupported.${key}.title`, DEFAULT_COPY.unsupported[key].title),
      text: str(raw.text, `copy.unsupported.${key}.text`, DEFAULT_COPY.unsupported[key].text),
    };
  };
  const copy: SimCopy = {
    paneTitle: str(copyRaw.paneTitle, "copy.paneTitle", DEFAULT_COPY.paneTitle),
    unsupported: { video: message("video"), book: message("book") },
    submitSteps:
      copyRaw.submitSteps === undefined ? DEFAULT_COPY.submitSteps : strings(copyRaw.submitSteps, "copy.submitSteps"),
  };
  if (copy.submitSteps.length !== DEFAULT_COPY.submitSteps.length) {
    fail("copy.submitSteps", `needs exactly ${DEFAULT_COPY.submitSteps.length} steps`);
  }

  const start =
    sheetRaw.start === undefined ? (tasks[0]?.hint ?? "A1") : cellRef(sheetRaw.start, "sheet.start");

  if (problems.length > 0) {
    throw new ModuleError(problems);
  }
  return {
    workbook,
    assignment,
    sheet: { rows, widths, cells, styles, outlines, start },
    tasks,
    totalPoints: tasks.reduce((sum, task) => sum + task.points, 0),
    note,
    autopilot,
    poster,
    copy,
  };
}

/** Without a script, the autopilot types each solution where the task starts and fills it across the range. */
export function defaultAutopilot(tasks: SimTask[]): PilotStep[] {
  return tasks.map((task) => {
    const range = parseRange(task.range);
    const single = !range || (range.c1 === range.c2 && range.r1 === range.r2);
    const cell = range ? addrName({ col: range.c1, row: range.r1 }) : task.hint;
    return { taskId: task.id, kind: "type", cell, formula: task.solution, fill: single ? undefined : task.range };
  });
}

/** The top-left of the sheet's content, at most five rows by four columns. */
function defaultPoster(cells: Inputs): Range {
  const addrs = Object.keys(cells)
    .map(parseAddr)
    .filter((addr): addr is NonNullable<typeof addr> => addr !== null);
  if (addrs.length === 0) {
    return { c1: 0, r1: 0, c2: 3, r2: 4 };
  }
  const c1 = Math.min(...addrs.map((a) => a.col));
  const r1 = Math.min(...addrs.map((a) => a.row));
  const c2 = Math.min(c1 + 3, Math.max(...addrs.map((a) => a.col)));
  const r2 = Math.min(r1 + 4, Math.max(...addrs.map((a) => a.row)));
  return { c1, r1, c2, r2 };
}
