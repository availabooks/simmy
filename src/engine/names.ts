/**
 * Defined names, as Excel's Name Box, Define Name, Create from Selection and Name Manager use
 * them. A name refers to a range on the sheet and is workbook-wide; names are compared
 * without regard to case but keep the case they were defined with.
 */
import {
  addrName,
  type Inputs,
  isFormula,
  type Names,
  parseConstant,
  parseRange,
  type Range,
  rangeName,
  tokenize,
} from "./sheet.ts";

/**
 * Why `text` can't be a name, or null if it can. Excel's rules: up to 255 characters; a letter,
 * underscore or backslash first, then letters, numbers, periods and underscores; no spaces;
 * nothing that reads as a cell reference (A1, R1C1), and not C, R, TRUE or FALSE.
 */
export function nameProblem(text: string): string | null {
  if (text === "") {
    return "a name can't be empty";
  }
  if (text.length > 255) {
    return "a name can be at most 255 characters";
  }
  if (/\s/.test(text)) {
    return `"${text}" has a space, which a name can't`;
  }
  if (!/^[A-Za-z_\\]/.test(text)) {
    return `"${text}" must start with a letter, an underscore or a backslash`;
  }
  if (!/^.[A-Za-z0-9_.]*$/.test(text)) {
    return `"${text}" can only use letters, numbers, periods and underscores`;
  }
  if (/^[A-Za-z]{1,3}\d{1,7}$/.test(text) || /^(R\d*C\d*|R\d*|C\d*)$/i.test(text)) {
    return `"${text}" looks like a cell reference, so it can't be a name`;
  }
  if (/^(TRUE|FALSE)$/i.test(text)) {
    return `"${text}" is a logical value, so it can't be a name`;
  }
  return null;
}

/** A name as defined, matched without regard to case, with what it refers to. */
export function lookupName(names: Names, text: string): { name: string; ref: string } | null {
  const upper = text.toUpperCase();
  const name = Object.keys(names).find((key) => key.toUpperCase() === upper);
  return name === undefined ? null : { name, ref: names[name] };
}

/** "$C$2" or "$A$2:$A$9": how a name stores what it refers to. */
export function absoluteRef(range: Range): string {
  return rangeName(range)
    .split(":")
    .map((part) => part.replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2"))
    .join(":");
}

/** The range a name refers to. */
export function nameRange(ref: string): Range | null {
  return parseRange(ref.replace(/\$/g, ""));
}

/** The name the Name Box shows for a selection: the first, alphabetically, that is exactly it. */
export function nameForRange(names: Names, range: Range): string | null {
  const ref = absoluteRef(range);
  return (
    Object.keys(names)
      .sort((a, b) => a.localeCompare(b))
      .find((name) => names[name] === ref) ?? null
  );
}

/**
 * What a Define Name dialog's "Refers to" box means: "=Sheet1!$C$2", "='My Sheet'!$C$2:$D$4"
 * or plain "C2". Null if it isn't a range on `sheet`.
 */
export function parseRefersTo(text: string, sheet: string): Range | null {
  let body = text.trim().replace(/^=/, "");
  const qualified = /^(?:'((?:[^']|'')+)'|([^!'\s]+))!(.+)$/.exec(body);
  if (qualified) {
    const name = qualified[1]?.replace(/''/g, "'") ?? qualified[2];
    if (name.toUpperCase() !== sheet.toUpperCase()) {
      return null;
    }
    body = qualified[3];
  }
  return parseRange(body.replace(/\$/g, ""));
}

/** "=Sheet1!$C$2": a range written the way Define Name's "Refers to" box shows it. */
export function refersToText(range: Range, sheet: string): string {
  const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  return `=${quoted}!${absoluteRef(range)}`;
}

/**
 * The name Create from Selection makes from a label, as Excel does: characters a name can't
 * have become underscores, a label that would read as a cell reference gets one on the end, and
 * one that starts with a number gets one in front. Null for an empty label.
 */
export function labelToName(label: string): string | null {
  let text = label.trim().replace(/[^A-Za-z0-9_.\\]/g, "_");
  if (text === "") {
    return null;
  }
  if (!/^[A-Za-z_\\]/.test(text)) {
    text = `_${text}`;
  }
  if (nameProblem(text) !== null) {
    text = `${text}_`;
  }
  text = text.slice(0, 255);
  return nameProblem(text) === null ? text : null;
}

/** Which sides of a selection hold its labels, for Create from Selection's check boxes. */
export type Sides = { top: boolean; left: boolean; bottom: boolean; right: boolean };

/** Whether a cell holds a label: text rather than a number or a formula. */
function isLabel(inputs: Inputs, col: number, row: number): boolean {
  const raw = inputs[addrName({ col, row })];
  return raw !== undefined && raw !== "" && !isFormula(raw) && typeof parseConstant(raw) === "string";
}

/**
 * The check boxes Create from Selection starts with: Top row when the top row holds labels,
 * Left column when the left column does. In a table with both, the corner cell they share
 * doesn't count; in a single column or row, it's the label.
 */
export function guessLabelSides(inputs: Inputs, range: Range): Sides {
  const wide = range.c2 > range.c1;
  const tall = range.r2 > range.r1;
  const span = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => from + i);
  const topCells = wide ? span(range.c1 + 1, range.c2) : [range.c1];
  const leftCells = tall ? span(range.r1 + 1, range.r2) : [range.r1];
  return {
    top: tall && topCells.every((col) => isLabel(inputs, col, range.r1)),
    left: wide && leftCells.every((row) => isLabel(inputs, range.c1, row)),
    bottom: false,
    right: false,
  };
}

/**
 * The names Create from Selection makes: one per labelled column (from the top or bottom row)
 * and one per labelled row (from the left or right column), each covering the cells beside its
 * label and not the other labels. Empty labels make no name.
 */
export function namesFromSelection(inputs: Inputs, range: Range, sides: Sides): Names {
  const out: Names = {};
  const body = {
    c1: range.c1 + (sides.left ? 1 : 0),
    c2: range.c2 - (sides.right ? 1 : 0),
    r1: range.r1 + (sides.top ? 1 : 0),
    r2: range.r2 - (sides.bottom ? 1 : 0),
  };
  if (body.c1 > body.c2 || body.r1 > body.r2) {
    return out;
  }
  const add = (labelCol: number, labelRow: number, target: Range) => {
    const name = labelToName(inputs[addrName({ col: labelCol, row: labelRow })] ?? "");
    if (name) {
      out[name] = absoluteRef(target);
    }
  };
  for (let col = body.c1; col <= body.c2; col++) {
    const column = { c1: col, c2: col, r1: body.r1, r2: body.r2 };
    if (sides.top) {
      add(col, range.r1, column);
    }
    if (sides.bottom) {
      add(col, range.r2, column);
    }
  }
  for (let row = body.r1; row <= body.r2; row++) {
    const line = { c1: body.c1, c2: body.c2, r1: row, r2: row };
    if (sides.left) {
      add(range.c1, row, line);
    }
    if (sides.right) {
      add(range.c2, row, line);
    }
  }
  return out;
}

/**
 * The name Define Name suggests: a label in the cell to the left of the selection, or else
 * above it. Empty if there's none.
 */
export function suggestName(inputs: Inputs, range: Range): string {
  const candidates = [
    { col: range.c1 - 1, row: range.r1 },
    { col: range.c1, row: range.r1 - 1 },
  ];
  for (const { col, row } of candidates) {
    if (col >= 0 && row >= 0 && isLabel(inputs, col, row)) {
      return labelToName(inputs[addrName({ col, row })]) ?? "";
    }
  }
  return "";
}

/** The names a formula uses (not its functions), as typed. */
export function formulaNames(formula: string): string[] {
  if (!isFormula(formula)) {
    return [];
  }
  try {
    const tokens = tokenize(formula.slice(1));
    return tokens
      .filter((token, i) => {
        const next = tokens[i + 1];
        const call = next?.kind === "op" && next.text === "(";
        return token.kind === "name" && !call && !/^(TRUE|FALSE)$/i.test(token.text);
      })
      .map((token) => token.text);
  } catch {
    return [];
  }
}

/**
 * A formula with every use of the name `from` (in any case) changed to `to`, as Excel rewrites
 * formulas when a name is renamed. A function of the same name, as in RATE(...), is left alone.
 */
export function renameInFormula(formula: string, from: string, to: string): string {
  if (!isFormula(formula)) {
    return formula;
  }
  const body = formula.slice(1);
  let tokens;
  try {
    tokens = tokenize(body);
  } catch {
    return formula;
  }
  const upper = from.toUpperCase();
  let out = "";
  let last = 0;
  tokens.forEach((token, i) => {
    const next = tokens[i + 1];
    const call = next?.kind === "op" && next.text === "(";
    if (token.kind === "name" && !call && token.text.toUpperCase() === upper) {
      out += body.slice(last, token.start) + to;
      last = token.start + token.text.length;
    }
  });
  return `=${out}${body.slice(last)}`;
}

/** The cell writes that rename a name in every formula that uses it. */
export function renameWrites(inputs: Inputs, from: string, to: string): Inputs {
  const writes: Inputs = {};
  for (const [key, raw] of Object.entries(inputs)) {
    const renamed = renameInFormula(raw, from, to);
    if (renamed !== raw) {
      writes[key] = renamed;
    }
  }
  return writes;
}
