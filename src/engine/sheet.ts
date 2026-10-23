/**
 * A small one-sheet spreadsheet engine for the simulator, with Excel 365's dynamic arrays.
 *
 * It parses and evaluates the formulas a student types, shifts relative
 * references the way Excel does when a formula is filled or pasted, and formats
 * values the way a cell shows them.
 */

export type Addr = { col: number; row: number };
export type Range = { c1: number; r1: number; c2: number; r2: number };

/** Raw cell input keyed by address, exactly as typed ("3.38", "=AVERAGE(D3:D37)"). */
export type Inputs = Record<string, string>;

/** Defined names: each name as defined (keeping its case) → its absolute range, "$C$2" or "$A$2:$A$9". */
export type Names = Record<string, string>;

export type ErrorCode = "#CALC!" | "#DIV/0!" | "#N/A" | "#NAME?" | "#NUM!" | "#REF!" | "#SPILL!" | "#VALUE!";
export type CellError = { error: ErrorCode };
export type Value = number | string | boolean | CellError | null;

/** "accounting" is Excel's `_($* #,##0.00_)` family: the $ sits at the cell's left edge. */
export type NumberFormat = "General" | "0" | "0.00" | "$#,##0.00" | "0.00%" | "accounting";

export class FormulaSyntaxError extends Error {}

export function colName(col: number): string {
  let n = col;
  let name = "";
  do {
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return name;
}

export function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

export function addrName({ col, row }: Addr): string {
  return `${colName(col)}${row + 1}`;
}

export function parseAddr(text: string): Addr | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(text.trim());
  if (!match) {
    return null;
  }
  const row = Number(match[2]) - 1;
  if (row < 0) {
    return null;
  }
  return { col: colIndex(match[1]), row };
}

export function parseRange(text: string): Range | null {
  const [a, b = a] = text.split(":");
  const start = parseAddr(a);
  const end = parseAddr(b);
  if (!start || !end) {
    return null;
  }
  return normalizeRange(start, end);
}

export function normalizeRange(a: Addr, b: Addr): Range {
  return {
    c1: Math.min(a.col, b.col),
    r1: Math.min(a.row, b.row),
    c2: Math.max(a.col, b.col),
    r2: Math.max(a.row, b.row),
  };
}

export function rangeName(range: Range): string {
  const start = addrName({ col: range.c1, row: range.r1 });
  if (range.c1 === range.c2 && range.r1 === range.r2) {
    return start;
  }
  return `${start}:${addrName({ col: range.c2, row: range.r2 })}`;
}

export function rangeAddrs(range: Range): Addr[] {
  const out: Addr[] = [];
  for (let row = range.r1; row <= range.r2; row++) {
    for (let col = range.c1; col <= range.c2; col++) {
      out.push({ col, row });
    }
  }
  return out;
}

export function inRange(range: Range, { col, row }: Addr): boolean {
  return col >= range.c1 && col <= range.c2 && row >= range.r1 && row <= range.r2;
}

export function isFormula(raw: string | undefined): raw is string {
  return typeof raw === "string" && raw.startsWith("=") && raw.length > 1;
}

export function isError(value: unknown): value is CellError {
  return typeof value === "object" && value !== null && "error" in value;
}

const err = (error: ErrorCode): CellError => ({ error });

/* ------------------------------------------------------------------ tokens */

type Token =
  | { kind: "num"; text: string; value: number; start: number }
  | { kind: "str"; text: string; value: string; start: number }
  | { kind: "ref"; text: string; start: number }
  | { kind: "name"; text: string; start: number }
  | { kind: "op"; text: string; start: number };

const REF_RE = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(?![A-Za-z0-9_.(])/;
const NUM_RE = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/;
const NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.]*/;
const OPS = ["<>", "<=", ">=", "+", "-", "*", "/", "^", "&", "=", "<", ">", "(", ")", ",", ":", "%"];

export function tokenize(body: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  while (i < body.length) {
    const rest = body.slice(i);
    if (/^\s/.test(rest)) {
      i++;
      continue;
    }
    if (rest[0] === '"') {
      let j = 1;
      let value = "";
      while (j < rest.length) {
        if (rest[j] === '"') {
          if (rest[j + 1] === '"') {
            value += '"';
            j += 2;
            continue;
          }
          break;
        }
        value += rest[j];
        j++;
      }
      if (j >= rest.length) {
        throw new FormulaSyntaxError("Unclosed string");
      }
      tokens.push({ kind: "str", text: rest.slice(0, j + 1), value, start: i });
      i += j + 1;
      continue;
    }
    const ref = REF_RE.exec(rest);
    if (ref) {
      tokens.push({ kind: "ref", text: ref[0], start: i });
      i += ref[0].length;
      continue;
    }
    const num = NUM_RE.exec(rest);
    if (num) {
      tokens.push({ kind: "num", text: num[0], value: Number(num[0]), start: i });
      i += num[0].length;
      continue;
    }
    const name = NAME_RE.exec(rest);
    if (name) {
      tokens.push({ kind: "name", text: name[0], start: i });
      i += name[0].length;
      continue;
    }
    const op = OPS.find((candidate) => rest.startsWith(candidate));
    if (op) {
      tokens.push({ kind: "op", text: op, start: i });
      i += op.length;
      continue;
    }
    throw new FormulaSyntaxError(`Unexpected "${rest[0]}"`);
  }
  return tokens;
}

/* --------------------------------------------------------------------- ast */

type Node =
  | { type: "num"; value: number }
  | { type: "str"; value: string }
  | { type: "bool"; value: boolean }
  | { type: "ref"; addr: Addr }
  | { type: "range"; range: Range }
  | { type: "call"; name: string; args: Node[] }
  /** An argument left empty, as in SEQUENCE(3,,10). */
  | { type: "missing" }
  /** A defined name, upper-cased; what it refers to is looked up when the formula runs. */
  | { type: "name"; name: string }
  | { type: "unary"; op: string; arg: Node }
  | { type: "percent"; arg: Node }
  | { type: "binary"; op: string; left: Node; right: Node };

const BINARY_LEVELS: string[][] = [["=", "<>", "<", ">", "<=", ">="], ["&"], ["+", "-"], ["*", "/"], ["^"]];

export function parseFormula(formula: string): Node {
  const body = formula.startsWith("=") ? formula.slice(1) : formula;
  const tokens = tokenize(body);
  let pos = 0;

  const peek = () => tokens[pos];
  const isOp = (text: string) => {
    const token = tokens[pos];
    return token !== undefined && token.kind === "op" && token.text === text;
  };
  const expectOp = (text: string) => {
    if (!isOp(text)) {
      throw new FormulaSyntaxError(`Expected "${text}"`);
    }
    pos++;
  };

  const parseLevel = (level: number): Node => {
    if (level >= BINARY_LEVELS.length) {
      return parseUnary();
    }
    let left = parseLevel(level + 1);
    for (;;) {
      const token = peek();
      if (!token || token.kind !== "op" || !BINARY_LEVELS[level].includes(token.text)) {
        return left;
      }
      pos++;
      const right = parseLevel(level + 1);
      left = { type: "binary", op: token.text, left, right };
    }
  };

  const parseUnary = (): Node => {
    if (isOp("-") || isOp("+")) {
      const op = tokens[pos].text;
      pos++;
      return { type: "unary", op, arg: parseUnary() };
    }
    let node = parsePrimary();
    while (isOp("%")) {
      pos++;
      node = { type: "percent", arg: node };
    }
    return node;
  };

  const parsePrimary = (): Node => {
    const token = peek();
    if (!token) {
      throw new FormulaSyntaxError("Formula ends too early");
    }
    pos++;
    switch (token.kind) {
      case "num":
        return { type: "num", value: token.value };
      case "str":
        return { type: "str", value: token.value };
      case "ref": {
        const start = parseAddr(token.text);
        if (!start) {
          throw new FormulaSyntaxError("Bad reference");
        }
        if (isOp(":")) {
          pos++;
          const next = peek();
          const end = next && next.kind === "ref" ? parseAddr(next.text) : null;
          if (!end) {
            throw new FormulaSyntaxError("Bad range");
          }
          pos++;
          return { type: "range", range: normalizeRange(start, end) };
        }
        return { type: "ref", addr: start };
      }
      case "name": {
        const upper = token.text.toUpperCase();
        if (isOp("(")) {
          pos++;
          const args: Node[] = [];
          const parseArg = (): Node => (isOp(",") || isOp(")") ? { type: "missing" } : parseLevel(0));
          if (!isOp(")")) {
            args.push(parseArg());
            while (isOp(",")) {
              pos++;
              args.push(parseArg());
            }
          }
          expectOp(")");
          return { type: "call", name: upper, args };
        }
        if (upper === "TRUE" || upper === "FALSE") {
          return { type: "bool", value: upper === "TRUE" };
        }
        return { type: "name", name: upper };
      }
      case "op":
        if (token.text === "(") {
          const inner = parseLevel(0);
          expectOp(")");
          return inner;
        }
        throw new FormulaSyntaxError(`Unexpected "${token.text}"`);
      default: {
        const _exhaustive: never = token;
        return _exhaustive;
      }
    }
  };

  if (tokens.length === 0) {
    throw new FormulaSyntaxError("Empty formula");
  }
  const tree = parseLevel(0);
  if (pos < tokens.length) {
    throw new FormulaSyntaxError(`Unexpected "${tokens[pos].text}"`);
  }
  return tree;
}

/** Throws FormulaSyntaxError when Excel would refuse to accept the formula. */
export function checkFormula(formula: string): void {
  parseFormula(formula);
}

/* --------------------------------------------------------------- functions */

/** Where Excel files a function: the Formulas tab's Function Library menus. */
export type FunctionCategory =
  | "Logical"
  | "Text"
  | "Math & Trig"
  | "Lookup & Reference"
  | "Statistical"
  | "Compatibility";

/** A function as Excel describes it: its tooltip, its arguments, and where it's filed. */
export type FunctionInfo = { name: string; summary: string; args: string; category: FunctionCategory };

export const FUNCTIONS: FunctionInfo[] = [
  {
    name: "ABS",
    summary: "Returns the absolute value of a number.",
    args: "number",
    category: "Math & Trig",
  },
  {
    name: "AVERAGE",
    summary: "Returns the average (arithmetic mean) of its arguments.",
    args: "number1, [number2], …",
    category: "Statistical",
  },
  {
    name: "CONCAT",
    summary: "Concatenates a list or range of text strings.",
    args: "text1, [text2], …",
    category: "Text",
  },
  {
    name: "CORREL",
    summary: "Returns the correlation coefficient between two data sets.",
    args: "array1, array2",
    category: "Statistical",
  },
  {
    name: "COUNT",
    summary: "Counts the number of cells in a range that contain numbers.",
    args: "value1, [value2], …",
    category: "Statistical",
  },
  {
    name: "COUNTA",
    summary: "Counts the number of cells in a range that are not empty.",
    args: "value1, [value2], …",
    category: "Statistical",
  },
  {
    name: "COUNTBLANK",
    summary: "Counts the number of empty cells in a specified range of cells.",
    args: "range",
    category: "Statistical",
  },
  {
    name: "COUNTIF",
    summary: "Counts the number of cells within a range that meet the given condition.",
    args: "range, criteria",
    category: "Statistical",
  },
  {
    name: "IF",
    summary: "Checks whether a condition is met, and returns one value if TRUE, and another if FALSE.",
    args: "logical_test, [value_if_true], [value_if_false]",
    category: "Logical",
  },
  {
    name: "LEFT",
    summary: "Returns the specified number of characters from the start of a text string.",
    args: "text, [num_chars]",
    category: "Text",
  },
  {
    name: "LEN",
    summary: "Returns the number of characters in a text string.",
    args: "text",
    category: "Text",
  },
  {
    name: "LOWER",
    summary: "Converts all letters in a text string to lowercase.",
    args: "text",
    category: "Text",
  },
  {
    name: "MAX",
    summary: "Returns the largest value in a set of values.",
    args: "number1, [number2], …",
    category: "Statistical",
  },
  {
    name: "MEDIAN",
    summary: "Returns the median, or the number in the middle of the set of given numbers.",
    args: "number1, [number2], …",
    category: "Statistical",
  },
  {
    name: "MID",
    summary: "Returns the characters from the middle of a text string, given a starting position and length.",
    args: "text, start_num, num_chars",
    category: "Text",
  },
  {
    name: "MIN",
    summary: "Returns the smallest number in a set of values.",
    args: "number1, [number2], …",
    category: "Statistical",
  },
  {
    name: "MODE",
    summary: "Returns the most frequently occurring value in a range of data.",
    args: "number1, [number2], …",
    category: "Compatibility",
  },
  {
    name: "MODE.SNGL",
    summary: "Returns the most frequently occurring value in a range of data.",
    args: "number1, [number2], …",
    category: "Statistical",
  },
  {
    name: "RANDBETWEEN",
    summary: "Returns a random number between the numbers you specify.",
    args: "bottom, top",
    category: "Math & Trig",
  },
  {
    name: "RIGHT",
    summary: "Returns the specified number of characters from the end of a text string.",
    args: "text, [num_chars]",
    category: "Text",
  },
  {
    name: "ROUND",
    summary: "Rounds a number to a specified number of digits.",
    args: "number, num_digits",
    category: "Math & Trig",
  },
  {
    name: "SEQUENCE",
    summary: "Returns a sequence of numbers.",
    args: "rows, [columns], [start], [step]",
    category: "Math & Trig",
  },
  {
    name: "SORT",
    summary: "Sorts a range or array.",
    args: "array, [sort_index], [sort_order], [by_col]",
    category: "Lookup & Reference",
  },
  { name: "SQRT", summary: "Returns the square root of a number.", args: "number", category: "Math & Trig" },
  {
    name: "STDEV",
    summary: "Estimates standard deviation based on a sample.",
    args: "number1, [number2], …",
    category: "Compatibility",
  },
  {
    name: "STDEV.P",
    summary: "Calculates standard deviation based on the entire population.",
    args: "number1, [number2], …",
    category: "Statistical",
  },
  {
    name: "STDEV.S",
    summary: "Estimates standard deviation based on a sample.",
    args: "number1, [number2], …",
    category: "Statistical",
  },
  {
    name: "SUM",
    summary: "Adds all the numbers in a range of cells.",
    args: "number1, [number2], …",
    category: "Math & Trig",
  },
  {
    name: "TEXTJOIN",
    summary: "Concatenates a list or range of text strings using a delimiter.",
    args: "delimiter, ignore_empty, text1, [text2], …",
    category: "Text",
  },
  {
    name: "UNIQUE",
    summary: "Returns the unique values from a range or array.",
    args: "array, [by_col], [exactly_once]",
    category: "Lookup & Reference",
  },
  {
    name: "UPPER",
    summary: "Converts a text string to all uppercase letters.",
    args: "text",
    category: "Text",
  },
  {
    name: "VAR",
    summary: "Estimates variance based on a sample.",
    args: "number1, [number2], …",
    category: "Compatibility",
  },
  {
    name: "VAR.P",
    summary: "Calculates variance based on the entire population.",
    args: "number1, [number2], …",
    category: "Statistical",
  },
  {
    name: "VAR.S",
    summary: "Estimates variance based on a sample.",
    args: "number1, [number2], …",
    category: "Statistical",
  },
];

const FUNCTION_SET = new Set(FUNCTIONS.map((fn) => fn.name));

export function suggestFunctions(prefix: string): FunctionInfo[] {
  const upper = prefix.toUpperCase();
  if (!upper) {
    return [];
  }
  return FUNCTIONS.filter((fn) => fn.name.startsWith(upper));
}

/* -------------------------------------------------------------- evaluation */

/**
 * An array a formula works with, as rows of values (at least 1×1). A range is one, and so is
 * what SEQUENCE or SORT returns. Cells never hold one: an array bigger than a cell spills.
 */
export type Matrix = Value[][];
type Result = Value | Matrix;

const isMatrix = (result: Result): result is Matrix => Array.isArray(result);

/**
 * A function argument: one value, or every value of a range or array in row order. `ref` says
 * it came from the sheet (A1, A1:B5) rather than being computed, which some functions care
 * about: COUNTIF and COUNTBLANK only take references.
 */
type Arg =
  | { kind: "scalar"; value: Value }
  | { kind: "range"; values: Value[]; rows: number; cols: number; ref: boolean };

function toNumber(value: Value): number | CellError {
  if (value === null) {
    return 0;
  }
  if (typeof value === "number") {
    return value;
  }
  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    const n = Number(trimmed);
    return trimmed !== "" && Number.isFinite(n) ? n : err("#VALUE!");
  }
  return value;
}

/** A number as Excel turns it into text: 15 significant digits, so 0.1+0.2 is "0.3". */
function numberText(n: number): string {
  return String(Number(n.toPrecision(15)));
}

function toText(value: Value): string {
  if (value === null) {
    return "";
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (typeof value === "number") {
    return numberText(value);
  }
  if (typeof value === "string") {
    return value;
  }
  return value.error;
}

/** Numbers in an aggregate: ranges skip text and blanks, direct arguments coerce. */
function collectNumbers(args: Arg[]): number[] | CellError {
  const out: number[] = [];
  for (const arg of args) {
    if (arg.kind === "range") {
      for (const value of arg.values) {
        if (isError(value)) {
          return value;
        }
        if (typeof value === "number") {
          out.push(value);
        }
      }
      continue;
    }
    if (arg.value === null) {
      continue;
    }
    const n = toNumber(arg.value);
    if (isError(n)) {
      return n;
    }
    out.push(n);
  }
  return out;
}

/** Run `fn` on the values as numbers, or return the first error among them. */
function withNumbers(values: Value[], fn: (ns: number[]) => Value): Value {
  const ns: number[] = [];
  for (const value of values) {
    const n = toNumber(value);
    if (isError(n)) {
      return n;
    }
    ns.push(n);
  }
  return fn(ns);
}

function variance(nums: number[], sample: boolean): number | CellError {
  const n = nums.length;
  if (n < (sample ? 2 : 1)) {
    return err("#DIV/0!");
  }
  const mean = nums.reduce((a, b) => a + b, 0) / n;
  const ss = nums.reduce((acc, x) => acc + (x - mean) ** 2, 0);
  return ss / (sample ? n - 1 : n);
}

function matchesCriteria(value: Value, criteria: Value): boolean {
  if (typeof criteria === "number") {
    return typeof value === "number" ? value === criteria : toNumber(value) === criteria && value !== null;
  }
  const text = toText(criteria);
  const match = /^(<>|<=|>=|=|<|>)?(.*)$/.exec(text);
  const op = match?.[1] ?? "=";
  const operand = match?.[2] ?? "";
  const asNumber = operand.trim() !== "" ? Number(operand) : Number.NaN;
  if (Number.isFinite(asNumber) && typeof value === "number") {
    switch (op) {
      case "=":
        return value === asNumber;
      case "<>":
        return value !== asNumber;
      case "<":
        return value < asNumber;
      case ">":
        return value > asNumber;
      case "<=":
        return value <= asNumber;
      case ">=":
        return value >= asNumber;
    }
  }
  const left = toText(value).toLowerCase();
  const right = operand.toLowerCase();
  if (op === "<>") {
    return left !== right;
  }
  return op === "=" && left === right && value !== null;
}

const AGGREGATES: Record<string, (nums: number[]) => Value> = {
  SUM: (nums) => nums.reduce((a, b) => a + b, 0),
  AVERAGE: (nums) => (nums.length === 0 ? err("#DIV/0!") : nums.reduce((a, b) => a + b, 0) / nums.length),
  MIN: (nums) => (nums.length === 0 ? 0 : Math.min(...nums)),
  MAX: (nums) => (nums.length === 0 ? 0 : Math.max(...nums)),
  MEDIAN: (nums) => {
    if (nums.length === 0) {
      return err("#NUM!");
    }
    const sorted = [...nums].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 === 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  },
  MODE: (nums) => {
    const counts = new Map<number, number>();
    let best: number | null = null;
    let bestCount = 1;
    for (const n of nums) {
      const count = (counts.get(n) ?? 0) + 1;
      counts.set(n, count);
      if (count > bestCount) {
        best = n;
        bestCount = count;
      }
    }
    if (best === null) {
      return err("#N/A");
    }
    // Excel breaks ties by first appearance, not by who reached the count first.
    for (const n of nums) {
      if (counts.get(n) === bestCount) {
        return n;
      }
    }
    return best;
  },
  "STDEV.S": (nums) => {
    const v = variance(nums, true);
    return isError(v) ? v : Math.sqrt(v);
  },
  "STDEV.P": (nums) => {
    const v = variance(nums, false);
    return isError(v) ? v : Math.sqrt(v);
  },
  "VAR.S": (nums) => variance(nums, true),
  "VAR.P": (nums) => variance(nums, false),
};
AGGREGATES["MODE.SNGL"] = AGGREGATES.MODE;
AGGREGATES.STDEV = AGGREGATES["STDEV.S"];
AGGREGATES.VAR = AGGREGATES["VAR.S"];

/**
 * Functions of single values. Given an array anywhere, they run once per element and return
 * an array (=ABS(A1:A3) spills three values), the way Excel 365 lifts them.
 */
type ScalarFunction = { min: number; max: number; fn: (...values: Value[]) => Value };

/** Text from a value, or the error it holds. */
function textOf(value: Value): string | CellError {
  return isError(value) ? value : toText(value);
}

/** LEFT and RIGHT: how many characters to take (1 when left out), or #VALUE! below 0. */
function charCount(value: Value | undefined): number | CellError {
  if (value === undefined) {
    return 1;
  }
  const n = toNumber(value);
  return isError(n) ? n : n < 0 ? err("#VALUE!") : Math.trunc(n);
}

const SCALAR_FUNCTIONS: Record<string, ScalarFunction> = {
  ABS: { min: 1, max: 1, fn: (x) => withNumbers([x], ([n]) => Math.abs(n)) },
  LEFT: {
    min: 1,
    max: 2,
    fn: (value, count) => {
      const text = textOf(value);
      const n = charCount(count);
      return isError(text) ? text : isError(n) ? n : text.slice(0, n);
    },
  },
  RIGHT: {
    min: 1,
    max: 2,
    fn: (value, count) => {
      const text = textOf(value);
      const n = charCount(count);
      return isError(text) ? text : isError(n) ? n : n === 0 ? "" : text.slice(-n);
    },
  },
  MID: {
    min: 3,
    max: 3,
    fn: (value, start, count) => {
      const text = textOf(value);
      if (isError(text)) {
        return text;
      }
      return withNumbers([start, count], ([from, n]) => {
        if (from < 1 || n < 0) {
          return err("#VALUE!");
        }
        const begin = Math.trunc(from) - 1;
        return text.slice(begin, begin + Math.trunc(n));
      });
    },
  },
  UPPER: {
    min: 1,
    max: 1,
    fn: (value) => {
      const text = textOf(value);
      return isError(text) ? text : text.toUpperCase();
    },
  },
  LOWER: {
    min: 1,
    max: 1,
    fn: (value) => {
      const text = textOf(value);
      return isError(text) ? text : text.toLowerCase();
    },
  },
  LEN: {
    min: 1,
    max: 1,
    fn: (value) => {
      const text = textOf(value);
      return isError(text) ? text : text.length;
    },
  },
  SQRT: { min: 1, max: 1, fn: (x) => withNumbers([x], ([n]) => (n < 0 ? err("#NUM!") : Math.sqrt(n))) },
  ROUND: {
    min: 2,
    max: 2,
    fn: (x, digits) =>
      withNumbers([x, digits], ([n, d]) => {
        const factor = 10 ** Math.trunc(d);
        return (Math.sign(n) * Math.round(Math.abs(n) * factor + 1e-9)) / factor;
      }),
  },
};

type Ctx = {
  cell: (addr: Addr) => Value;
  /** The range a defined name refers to (given upper-cased), or null if there's no such name. */
  name: (upper: string) => Range | null;
  /** A random number in [0, 1), the same for the same seed, cell and call. */
  rand: () => number;
};

function rangeMatrix(range: Range, ctx: Ctx): Matrix {
  const rows: Matrix = [];
  for (let row = range.r1; row <= range.r2; row++) {
    const line: Value[] = [];
    for (let col = range.c1; col <= range.c2; col++) {
      line.push(ctx.cell({ col, row }));
    }
    rows.push(line);
  }
  return rows;
}

function evalArg(node: Node, ctx: Ctx): Arg {
  // A reference is a range to a function, even a single cell: SUM(A1) skips text in A1.
  const ref =
    node.type === "ref" || node.type === "range" || (node.type === "name" && ctx.name(node.name) !== null);
  const result = node.type === "ref" ? [[ctx.cell(node.addr)]] : evalNode(node, ctx);
  if (!isMatrix(result)) {
    return { kind: "scalar", value: result };
  }
  return { kind: "range", values: result.flat(), rows: result.length, cols: result[0].length, ref };
}

/** An argument as rows of values: a single value is a 1×1 array. */
function argMatrix(arg: Arg): Matrix {
  if (arg.kind === "scalar") {
    return [[arg.value]];
  }
  const rows: Matrix = [];
  for (let r = 0; r < arg.rows; r++) {
    rows.push(arg.values.slice(r * arg.cols, (r + 1) * arg.cols));
  }
  return rows;
}

function transpose(m: Matrix): Matrix {
  return m[0].map((_, c) => m.map((row) => row[c]));
}

/**
 * SORT's order, ascending: numbers, then text (ignoring case), then TRUE/FALSE, then errors.
 * Empty cells always go last, whichever way the sort runs.
 */
function sortRank(value: Value): number {
  if (typeof value === "number") {
    return 0;
  }
  if (typeof value === "string") {
    return 1;
  }
  if (typeof value === "boolean") {
    return 2;
  }
  return value === null ? 4 : 3;
}

function sortCompare(a: Value, b: Value): number {
  const rank = sortRank(a) - sortRank(b);
  if (rank !== 0) {
    return rank;
  }
  if (typeof a === "number" && typeof b === "number") {
    return a - b;
  }
  if (typeof a === "string" && typeof b === "string") {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof a === "boolean" && typeof b === "boolean") {
    return Number(a) - Number(b);
  }
  return 0;
}

/** What makes two values the same to UNIQUE: text ignores case, and 1 and "1" differ. */
function uniqueKey(value: Value): string {
  if (value === null) {
    return "blank";
  }
  if (isError(value)) {
    return `error:${value.error}`;
  }
  return `${typeof value}:${typeof value === "string" ? value.toLowerCase() : String(value)}`;
}

/** A whole-number argument, or `fallback` when the argument was left out or left empty. */
function optionalNumber(node: Node | undefined, ctx: Ctx, fallback: number): number | CellError {
  if (!node || node.type === "missing") {
    return fallback;
  }
  const value = evalNode(node, ctx);
  if (isMatrix(value)) {
    return err("#VALUE!");
  }
  if (isError(value)) {
    return value;
  }
  return toNumber(value);
}

function scalarOf(arg: Arg): Value {
  if (arg.kind === "scalar") {
    return arg.value;
  }
  return arg.values.length === 1 ? arg.values[0] : err("#VALUE!");
}

/**
 * Apply a function of single values across arrays, element by element. A side of length 1
 * stretches to fit the others; where an array is too short for the result, Excel gives #N/A.
 */
function lift(args: Result[], fn: (...values: Value[]) => Value): Result {
  if (!args.some(isMatrix)) {
    return fn(...(args as Value[]));
  }
  const rows = Math.max(...args.map((arg) => (isMatrix(arg) ? arg.length : 1)));
  const cols = Math.max(...args.map((arg) => (isMatrix(arg) ? arg[0].length : 1)));
  const out: Matrix = [];
  for (let i = 0; i < rows; i++) {
    const line: Value[] = [];
    for (let j = 0; j < cols; j++) {
      const values: Value[] = [];
      let short = false;
      for (const arg of args) {
        if (!isMatrix(arg)) {
          values.push(arg);
          continue;
        }
        const r = arg.length === 1 ? 0 : i;
        const c = arg[0].length === 1 ? 0 : j;
        if (r >= arg.length || c >= arg[0].length) {
          short = true;
          break;
        }
        values.push(arg[r][c]);
      }
      line.push(short ? err("#N/A") : fn(...values));
    }
    out.push(line);
  }
  return out;
}

/** IF's test as true or false, or the error that stops it. */
function truth(value: Value): boolean | CellError {
  const n = toNumber(value);
  return isError(n) ? n : n !== 0;
}

function callIf(nodes: Node[], ctx: Ctx): Result {
  if (nodes.length < 2 || nodes.length > 3) {
    return err("#VALUE!");
  }
  // A branch left empty, as in IF(A1,,"x"), returns 0; a missing third argument returns FALSE.
  const branch = (i: number): Result => {
    const node = nodes[i];
    if (!node) {
      return false;
    }
    return node.type === "missing" ? 0 : evalNode(node, ctx);
  };
  const test = evalNode(nodes[0], ctx);
  if (isMatrix(test)) {
    return lift([test, branch(1), branch(2)], (t, yes, no) => {
      const pass = truth(t);
      return isError(pass) ? pass : pass ? yes : no;
    });
  }
  // Otherwise only the branch taken is worked out, as in Excel.
  const pass = truth(test);
  return isError(pass) ? pass : branch(pass ? 1 : 2);
}

function callFunction(name: string, nodes: Node[], ctx: Ctx): Result {
  if (!FUNCTION_SET.has(name)) {
    return err("#NAME?");
  }
  if (name === "IF") {
    return callIf(nodes, ctx);
  }
  if (name === "RANDBETWEEN") {
    if (nodes.length !== 2) {
      return err("#VALUE!");
    }
    return lift(
      nodes.map((node) => evalNode(node, ctx)),
      (bottom, top) =>
        withNumbers([bottom, top], ([lo, hi]) => {
          const low = Math.ceil(lo);
          const high = Math.floor(hi);
          return low > high ? err("#NUM!") : low + Math.floor(ctx.rand() * (high - low + 1));
        }),
    );
  }
  if (name === "SEQUENCE") {
    if (nodes.length < 1 || nodes.length > 4) {
      return err("#VALUE!");
    }
    const parts = [
      optionalNumber(nodes[0], ctx, 1),
      optionalNumber(nodes[1], ctx, 1),
      optionalNumber(nodes[2], ctx, 1),
      optionalNumber(nodes[3], ctx, 1),
    ];
    const problem = parts.find(isError);
    if (problem) {
      return problem;
    }
    const [rows, cols, start, step] = parts as number[];
    const height = Math.trunc(rows);
    const width = Math.trunc(cols);
    if (height < 1 || width < 1) {
      return err("#CALC!");
    }
    return Array.from({ length: height }, (_, r) =>
      Array.from({ length: width }, (_, c) => start + step * (r * width + c)),
    );
  }
  const scalar = SCALAR_FUNCTIONS[name];
  if (scalar) {
    if (nodes.length < scalar.min || nodes.length > scalar.max) {
      return err("#VALUE!");
    }
    return lift(
      nodes.map((node) => evalNode(node, ctx)),
      scalar.fn,
    );
  }

  const args = nodes.map((node) => evalArg(node, ctx));
  const aggregate = AGGREGATES[name];
  if (aggregate) {
    if (args.length === 0) {
      return err("#VALUE!");
    }
    const nums = collectNumbers(args);
    return isError(nums) ? nums : aggregate(nums);
  }

  switch (name) {
    case "COUNTA":
      // Everything but an empty cell counts, errors and "" from a formula included.
      return args.reduce(
        (count, arg) => count + (arg.kind === "range" ? arg.values.filter((v) => v !== null).length : 1),
        0,
      );
    case "COUNT":
      if (args.length === 0) {
        return err("#VALUE!");
      }
      // Unlike SUM, COUNT skips errors. A value typed into it counts if it reads as a number.
      return args.reduce(
        (count, arg) =>
          count +
          (arg.kind === "range"
            ? arg.values.filter((v) => typeof v === "number").length
            : arg.value !== null && !isError(arg.value) && !isError(toNumber(arg.value))
              ? 1
              : 0),
        0,
      );
    case "COUNTBLANK":
      if (args.length !== 1 || args[0].kind !== "range" || !args[0].ref) {
        return err("#VALUE!");
      }
      // A formula's "" counts as blank; a 0 doesn't.
      return args[0].values.filter((v) => v === null || v === "").length;
    case "CONCAT": {
      if (args.length === 0) {
        return err("#VALUE!");
      }
      let text = "";
      for (const value of args.flatMap((arg) => (arg.kind === "range" ? arg.values : [arg.value]))) {
        if (isError(value)) {
          return value;
        }
        text += toText(value);
      }
      return text.length > 32767 ? err("#VALUE!") : text;
    }
    case "TEXTJOIN": {
      if (args.length < 3) {
        return err("#VALUE!");
      }
      // A range of delimiters is used in turn.
      const delimiters = args[0].kind === "range" ? args[0].values : [args[0].value];
      const skipEmpty = truth(scalarOf(args[1]));
      if (isError(skipEmpty)) {
        return skipEmpty;
      }
      const parts: string[] = [];
      for (const value of args.slice(2).flatMap((arg) => (arg.kind === "range" ? arg.values : [arg.value]))) {
        if (isError(value)) {
          return value;
        }
        if (skipEmpty && (value === null || value === "")) {
          continue;
        }
        parts.push(toText(value));
      }
      let text = parts[0] ?? "";
      for (let i = 1; i < parts.length; i++) {
        const delimiter = delimiters[(i - 1) % delimiters.length];
        if (isError(delimiter)) {
          return delimiter;
        }
        text += toText(delimiter) + parts[i];
      }
      return text.length > 32767 ? err("#VALUE!") : text;
    }
    case "SORT": {
      if (args.length < 1 || args.length > 4) {
        return err("#VALUE!");
      }
      const byCol = args[3] ? truth(scalarOf(args[3])) : false;
      if (isError(byCol)) {
        return byCol;
      }
      const index = args[1] ? toNumber(scalarOf(args[1]) ?? 1) : 1;
      const order = args[2] ? toNumber(scalarOf(args[2]) ?? 1) : 1;
      if (isError(index)) {
        return index;
      }
      if (isError(order)) {
        return order;
      }
      // Sorting columns is sorting the rows of the turned-over array.
      const lines = byCol ? transpose(argMatrix(args[0])) : argMatrix(args[0]);
      const key = Math.trunc(index) - 1;
      if (key < 0 || key >= lines[0].length || (order !== 1 && order !== -1)) {
        return err("#VALUE!");
      }
      const sorted = [...lines].sort((a, b) => {
        const x = a[key];
        const y = b[key];
        if ((x === null) !== (y === null)) {
          return x === null ? 1 : -1;
        }
        return order * sortCompare(x, y);
      });
      return byCol ? transpose(sorted) : sorted;
    }
    case "UNIQUE": {
      if (args.length < 1 || args.length > 3) {
        return err("#VALUE!");
      }
      const byCol = args[1] ? truth(scalarOf(args[1])) : false;
      const once = args[2] ? truth(scalarOf(args[2])) : false;
      if (isError(byCol)) {
        return byCol;
      }
      if (isError(once)) {
        return once;
      }
      const lines = byCol ? transpose(argMatrix(args[0])) : argMatrix(args[0]);
      const counts = new Map<string, number>();
      const keyOf = (line: Value[]) => line.map(uniqueKey).join("\u0000");
      for (const line of lines) {
        counts.set(keyOf(line), (counts.get(keyOf(line)) ?? 0) + 1);
      }
      const seen = new Set<string>();
      const kept = lines.filter((line) => {
        const k = keyOf(line);
        if (seen.has(k) || (once && counts.get(k) !== 1)) {
          return false;
        }
        seen.add(k);
        return true;
      });
      if (kept.length === 0) {
        return err("#CALC!");
      }
      return byCol ? transpose(kept) : kept;
    }
    case "COUNTIF": {
      if (args.length !== 2 || args[0].kind !== "range" || !args[0].ref) {
        return err("#VALUE!");
      }
      const criteria = scalarOf(args[1]);
      if (isError(criteria)) {
        return criteria;
      }
      return args[0].values.filter((value) => matchesCriteria(value, criteria)).length;
    }
    case "CORREL": {
      if (args.length !== 2 || args[0].kind !== "range" || args[1].kind !== "range") {
        return err("#VALUE!");
      }
      const xs = args[0].values;
      const ys = args[1].values;
      if (xs.length !== ys.length) {
        return err("#N/A");
      }
      const pairs: [number, number][] = [];
      xs.forEach((x, i) => {
        const y = ys[i];
        if (typeof x === "number" && typeof y === "number") {
          pairs.push([x, y]);
        }
      });
      const n = pairs.length;
      if (n < 2) {
        return err("#DIV/0!");
      }
      const mx = pairs.reduce((a, [x]) => a + x, 0) / n;
      const my = pairs.reduce((a, [, y]) => a + y, 0) / n;
      let sxy = 0;
      let sxx = 0;
      let syy = 0;
      for (const [x, y] of pairs) {
        sxy += (x - mx) * (y - my);
        sxx += (x - mx) ** 2;
        syy += (y - my) ** 2;
      }
      return sxx === 0 || syy === 0 ? err("#DIV/0!") : sxy / Math.sqrt(sxx * syy);
    }
    default:
      return err("#NAME?");
  }
}

function compare(left: Value, right: Value, op: string): boolean {
  const l = left ?? (typeof right === "string" ? "" : 0);
  const r = right ?? (typeof left === "string" ? "" : 0);
  let diff: number;
  if (typeof l === "number" && typeof r === "number") {
    diff = l - r;
  } else if (typeof l === typeof r) {
    diff = toText(l).toLowerCase().localeCompare(toText(r).toLowerCase());
  } else {
    // Excel orders numbers < text < booleans.
    const rank = (v: Value) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
    diff = rank(l) - rank(r);
  }
  switch (op) {
    case "=":
      return diff === 0;
    case "<>":
      return diff !== 0;
    case "<":
      return diff < 0;
    case ">":
      return diff > 0;
    case "<=":
      return diff <= 0;
    default:
      return diff >= 0;
  }
}

function binary(op: string, left: Value, right: Value): Value {
  if (isError(left)) {
    return left;
  }
  if (isError(right)) {
    return right;
  }
  if (op === "&") {
    return toText(left) + toText(right);
  }
  if (["=", "<>", "<", ">", "<=", ">="].includes(op)) {
    return compare(left, right, op);
  }
  return withNumbers([left, right], ([a, b]) => {
    switch (op) {
      case "+":
        return a + b;
      case "-":
        return a - b;
      case "*":
        return a * b;
      case "/":
        return b === 0 ? err("#DIV/0!") : a / b;
      default: {
        const p = a ** b;
        return Number.isFinite(p) ? p : err("#NUM!");
      }
    }
  });
}

function evalNode(node: Node, ctx: Ctx): Result {
  switch (node.type) {
    case "num":
    case "str":
    case "bool":
      return node.value;
    case "missing":
      return null;
    case "ref":
      return ctx.cell(node.addr);
    case "range":
      return rangeMatrix(node.range, ctx);
    case "name": {
      const range = ctx.name(node.name);
      if (!range) {
        return err("#NAME?");
      }
      return range.c1 === range.c2 && range.r1 === range.r2
        ? ctx.cell({ col: range.c1, row: range.r1 })
        : rangeMatrix(range, ctx);
    }
    case "call":
      return callFunction(node.name, node.args, ctx);
    case "unary":
      return lift([evalNode(node.arg, ctx)], (value) =>
        withNumbers([value], ([n]) => (node.op === "-" ? -n : n)),
      );
    case "percent":
      return lift([evalNode(node.arg, ctx)], (value) => withNumbers([value], ([n]) => n / 100));
    case "binary":
      return lift([evalNode(node.left, ctx), evalNode(node.right, ctx)], (left, right) =>
        binary(node.op, left, right),
      );
    default: {
      const _exhaustive: never = node;
      return _exhaustive;
    }
  }
}

/** A typed constant: numbers become numbers, TRUE/FALSE become booleans, the rest stays text. */
export function parseConstant(raw: string): Value {
  const trimmed = raw.trim();
  if (trimmed === "") {
    return null;
  }
  const plain = trimmed.replace(/^\$/, "").replace(/,/g, "");
  const percent = /^(-?\d*\.?\d+)%$/.exec(plain);
  if (percent) {
    return Number(percent[1]) / 100;
  }
  if (/^-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(plain)) {
    return Number(plain);
  }
  const upper = trimmed.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") {
    return upper === "TRUE";
  }
  return raw;
}

export type EvalOptions = {
  /** The sheet's size. An array that would spill past it is #SPILL!. */
  bounds?: { rows: number; cols: number };
  /** Defined names the formulas can use. */
  names?: Names;
  /**
   * Picks RANDBETWEEN's numbers. The same seed gives the same numbers, so what a cell shows
   * and what grading sees agree. Change it to recalculate, as Excel does on every edit.
   */
  seed?: number;
};

/** A small fast generator (mulberry32): good enough to pick numbers, and repeatable. */
function mulberry32(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function hashText(text: string): number {
  let h = 2166136261;
  for (let i = 0; i < text.length; i++) {
    h = Math.imul(h ^ text.charCodeAt(i), 16777619);
  }
  return h >>> 0;
}

/** Every cell's value, and where dynamic arrays spilled. */
export type SheetEval = {
  /** Every cell with a value, including the cells a spill fills. */
  values: Record<string, Value>;
  /** Anchor cell → the range its array fills. */
  spills: Record<string, Range>;
  /** Anchor cell → the range it would fill, for a spill with something in the way. */
  blocked: Record<string, Range>;
  /** Cell filled by a spill → its anchor. */
  owner: Record<string, string>;
};

/** What a cell shows for a formula's result: an empty reference shows 0, infinity is #NUM!. */
function cellValue(value: Value): Value {
  if (value === null) {
    return 0;
  }
  return typeof value === "number" && !Number.isFinite(value) ? err("#NUM!") : value;
}

/** Names looked up without regard to case, as Excel does. */
function namedRanges(names: Names = {}): Map<string, Range> {
  const out = new Map<string, Range>();
  for (const [name, ref] of Object.entries(names)) {
    const range = parseRange(ref.replace(/\$/g, ""));
    if (range) {
      out.set(name.toUpperCase(), range);
    }
  }
  return out;
}

const byPosition = (a: string, b: string): number => {
  const x = parseAddr(a) as Addr;
  const y = parseAddr(b) as Addr;
  return x.row - y.row || x.col - y.col;
};

/**
 * Evaluates a sheet on demand. A formula whose result is an array spills it into the cells
 * below and to the right, unless any of them holds something, another spill got there first,
 * or it runs off the sheet; then the formula shows #SPILL!. Formulas are worked out in reading
 * order, so when two spills collide the earlier one always wins.
 */
function sheetEvaluator(inputs: Inputs, opts: EvalOptions) {
  const out: SheetEval = { values: {}, spills: {}, blocked: {}, owner: {} };
  const { values } = out;
  const visiting = new Set<string>();
  const done = new Set<string>();
  const parsed = new Map<string, Node | null>();
  const formulaKeys = Object.keys(inputs)
    .filter((key) => isFormula(inputs[key]) && parseAddr(key))
    .sort(byPosition);
  let settled = false;

  // Each cell gets its own stream of random numbers, so one cell's RANDBETWEEN doesn't move
  // when another formula is added.
  let current = "";
  const streams = new Map<string, () => number>();
  const nameRanges = namedRanges(opts.names);
  const ctx: Ctx = {
    cell: (addr) => cell(addr),
    name: (upper) => nameRanges.get(upper) ?? null,
    rand: () => {
      let next = streams.get(current);
      if (!next) {
        next = mulberry32(hashText(`${opts.seed ?? 1}|${current}`));
        streams.set(current, next);
      }
      return next();
    },
  };

  /**
   * Work out every formula not yet done, so all the spills are in place. It can run again from
   * inside itself (a formula it works out reads an empty cell), which is safe: formulas being
   * worked out are skipped, and each is worked out once.
   */
  const settle = () => {
    if (settled) {
      return;
    }
    for (const key of formulaKeys) {
      if (!done.has(key) && !visiting.has(key)) {
        evaluateCell(key);
      }
    }
    settled = done.size === formulaKeys.length;
  };

  function cell(addr: Addr): Value {
    const key = addrName(addr);
    if (key in values) {
      return values[key];
    }
    const raw = inputs[key];
    if (raw === undefined || raw === "") {
      // An empty cell may be part of a spill whose formula hasn't been worked out yet.
      settle();
      return values[key] ?? null;
    }
    if (!isFormula(raw)) {
      const constant = parseConstant(raw);
      values[key] = constant;
      return constant;
    }
    if (visiting.has(key)) {
      return 0;
    }
    evaluateCell(key);
    return values[key];
  }

  function evaluateCell(key: string): void {
    visiting.add(key);
    const outer = current;
    current = key;
    let tree = parsed.get(key);
    if (tree === undefined) {
      try {
        tree = parseFormula(inputs[key]);
      } catch {
        tree = null;
      }
      parsed.set(key, tree);
    }
    const result = tree ? evalNode(tree, ctx) : err("#NAME?");
    current = outer;
    visiting.delete(key);
    done.add(key);
    place(key, result);
  }

  function place(key: string, result: Result): void {
    if (!isMatrix(result)) {
      values[key] = cellValue(result);
      return;
    }
    const rows = result.length;
    const cols = result[0].length;
    if (rows === 1 && cols === 1) {
      values[key] = cellValue(result[0][0]);
      return;
    }
    const at = parseAddr(key) as Addr;
    const target = { c1: at.col, r1: at.row, c2: at.col + cols - 1, r2: at.row + rows - 1 };
    const addrs = rangeAddrs(target);
    const { bounds } = opts;
    const offSheet = bounds !== undefined && (target.c2 >= bounds.cols || target.r2 >= bounds.rows);
    const inTheWay = addrs.some((addr) => {
      const other = addrName(addr);
      return other !== key && ((inputs[other] ?? "") !== "" || other in out.owner);
    });
    if (offSheet || inTheWay) {
      values[key] = err("#SPILL!");
      out.blocked[key] = target;
      return;
    }
    out.spills[key] = target;
    // rangeAddrs runs in reading order, the same order as the array's rows.
    addrs.forEach((addr, i) => {
      const other = addrName(addr);
      values[other] = cellValue(result[Math.floor(i / cols)][i % cols]);
      if (other !== key) {
        out.owner[other] = key;
      }
    });
  }

  return {
    ctx,
    settle,
    /** Every cell's value and the spills. */
    finish(): SheetEval {
      settle();
      for (const key of Object.keys(inputs)) {
        const addr = parseAddr(key);
        if (addr) {
          cell(addr);
        }
      }
      return out;
    },
  };
}

/** Evaluate every cell, placing each dynamic array where it spills. */
export function evaluateBook(inputs: Inputs, opts: EvalOptions = {}): SheetEval {
  return sheetEvaluator(inputs, opts).finish();
}

/** Evaluate every cell. Circular references resolve to 0, as Excel shows after its warning. */
export function evaluateSheet(inputs: Inputs, opts: EvalOptions = {}): Record<string, Value> {
  return evaluateBook(inputs, opts).values;
}

/**
 * Evaluate a formula against the sheet without putting it in a cell. An array result comes
 * back whole, rather than spilled.
 */
export function evaluateExpression(inputs: Inputs, formula: string, opts: EvalOptions = {}): Value | Matrix {
  const sheet = sheetEvaluator(inputs, opts);
  sheet.settle();
  let tree: Node;
  try {
    tree = parseFormula(formula);
  } catch {
    return err("#NAME?");
  }
  const result = evalNode(tree, sheet.ctx);
  return isMatrix(result) ? result.map((row) => row.map(cellValue)) : cellValue(result);
}

/** What a formula would show in a cell: an array's top-left value. */
export function evaluateFormulaAt(inputs: Inputs, formula: string, _at?: Addr, opts: EvalOptions = {}): Value {
  const result = evaluateExpression(inputs, formula, opts);
  return isMatrix(result) ? result[0][0] : result;
}

/* ---------------------------------------------------------------- shifting */

/** Move relative references by (dCol, dRow); $-anchored parts stay put. */
export function shiftFormula(formula: string, dCol: number, dRow: number): string {
  if (!isFormula(formula) || (dCol === 0 && dRow === 0)) {
    return formula;
  }
  const body = formula.slice(1);
  let tokens: Token[];
  try {
    tokens = tokenize(body);
  } catch {
    return formula;
  }
  let out = "";
  let last = 0;
  for (const token of tokens) {
    if (token.kind !== "ref") {
      continue;
    }
    const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(token.text);
    if (!match) {
      continue;
    }
    const [, colAbs, colText, rowAbs, rowText] = match;
    const col = colAbs ? colIndex(colText) : colIndex(colText) + dCol;
    const row = rowAbs ? Number(rowText) - 1 : Number(rowText) - 1 + dRow;
    const replacement = col < 0 || row < 0 ? "#REF!" : `${colAbs}${colName(col)}${rowAbs}${row + 1}`;
    out += body.slice(last, token.start) + replacement;
    last = token.start + token.text.length;
  }
  return `=${out}${body.slice(last)}`;
}

/** The ranges a formula points at, in order, for the colored reference outlines. Names count. */
export function formulaRanges(formula: string, names: Names = {}): Range[] {
  if (!formula.startsWith("=")) {
    return [];
  }
  let tokens: Token[];
  try {
    tokens = tokenize(formula.slice(1));
  } catch {
    return [];
  }
  const out: Range[] = [];
  const named = namedRanges(names);
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    const call = tokens[i + 1]?.kind === "op" && tokens[i + 1].text === "(";
    const range = token.kind === "name" && !call ? named.get(token.text.toUpperCase()) : undefined;
    if (range) {
      out.push(range);
      continue;
    }
    if (token.kind !== "ref") {
      continue;
    }
    const start = parseAddr(token.text);
    if (!start) {
      continue;
    }
    const colon = tokens[i + 1];
    const next = tokens[i + 2];
    if (colon?.kind === "op" && colon.text === ":" && next?.kind === "ref") {
      const end = parseAddr(next.text);
      if (end) {
        out.push(normalizeRange(start, end));
        i += 2;
        continue;
      }
    }
    out.push(normalizeRange(start, start));
  }
  return out;
}

/** Function names used in a formula, upper-cased. */
export function formulaFunctions(formula: string): string[] {
  if (!isFormula(formula)) {
    return [];
  }
  try {
    const tokens = tokenize(formula.slice(1));
    return tokens
      .filter((token, i) => {
        const next = tokens[i + 1];
        return token.kind === "name" && next?.kind === "op" && next.text === "(";
      })
      .map((token) => token.text.toUpperCase());
  } catch {
    return [];
  }
}

/* -------------------------------------------------------------- formatting */

const CHAR_PX = 7.2;

function fitsWidth(text: string, widthPx: number): boolean {
  return text.length * CHAR_PX <= widthPx - 6;
}

function groupThousands(digits: string): string {
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function formatGeneralNumber(value: number, widthPx: number): string {
  if (value === 0) {
    return "0";
  }
  const maxChars = Math.max(1, Math.floor((widthPx - 6) / CHAR_PX));
  if (Number.isInteger(value) && Math.abs(value) < 1e11) {
    const text = String(value);
    if (text.length <= maxChars) {
      return text;
    }
  }
  for (let precision = 10; precision >= 1; precision--) {
    const fixed = Number(value.toPrecision(precision));
    let text = String(fixed);
    if (/e/.test(text)) {
      text = fixed
        .toExponential(Math.max(0, precision - 1))
        .replace(/\.?0+e/, "E")
        .replace("e", "E");
    }
    if (text.length <= maxChars) {
      return text;
    }
  }
  return "#".repeat(maxChars);
}

/** Display text for a value, given the cell's number format and column width. */
export function formatValue(value: Value, format: NumberFormat, widthPx: number): string {
  if (value === null) {
    return "";
  }
  if (isError(value)) {
    return value.error;
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (typeof value === "string") {
    return value;
  }
  let text: string;
  switch (format) {
    case "0":
      text = Math.round(value).toString();
      break;
    case "0.00":
      text = value.toFixed(2);
      break;
    case "$#,##0.00": {
      const [whole, cents] = Math.abs(value).toFixed(2).split(".");
      text = `${value < 0 ? "-" : ""}$${groupThousands(whole)}.${cents}`;
      break;
    }
    case "0.00%":
      text = `${(value * 100).toFixed(2)}%`;
      break;
    case "accounting": {
      const [whole, cents] = Math.abs(value).toFixed(2).split(".");
      const body = `${groupThousands(whole)}.${cents}`;
      if (whole === "0" && cents === "00") {
        text = "$ -   ";
      } else {
        text = value < 0 ? `$ (${body})` : `$ ${body} `;
      }
      break;
    }
    default:
      return formatGeneralNumber(value, widthPx);
  }
  if (text === "-0" || text === "-0.00") {
    text = text.slice(1);
  }
  return fitsWidth(text, widthPx) ? text : "#".repeat(Math.max(1, Math.floor((widthPx - 6) / CHAR_PX)));
}

/* ----------------------------------------------------------------- editing */

/**
 * What Excel does on Enter: close dangling parentheses, upper-case references and functions,
 * and write each defined name the way it was defined. Unknown names stay as typed.
 */
export function normalizeFormula(formula: string, names: Names = {}): string {
  if (!isFormula(formula)) {
    return formula;
  }
  let body = formula.slice(1);
  let tokens: Token[];
  try {
    tokens = tokenize(body);
  } catch {
    return formula;
  }
  const open = tokens.filter((t) => t.kind === "op" && t.text === "(").length;
  const close = tokens.filter((t) => t.kind === "op" && t.text === ")").length;
  const defined = new Map(Object.keys(names).map((name) => [name.toUpperCase(), name]));
  let out = "";
  let last = 0;
  tokens.forEach((token, i) => {
    if (token.kind !== "ref" && token.kind !== "name") {
      return;
    }
    const upper = token.text.toUpperCase();
    const call = tokens[i + 1]?.kind === "op" && tokens[i + 1].text === "(";
    const text =
      token.kind === "ref" || call || upper === "TRUE" || upper === "FALSE"
        ? upper
        : (defined.get(upper) ?? token.text);
    out += body.slice(last, token.start) + text;
    last = token.start + token.text.length;
  });
  body = out + body.slice(last);
  return `=${body}${")".repeat(Math.max(0, open - close))}`;
}

/**
 * F4 while typing a formula: the reference at the caret, or every reference the selection
 * touches, moves to its next form, A1 → $A$1 → A$1 → $A1 → A1. Both ends of a range such as
 * A1:B2 change together, taking the form after the first end's. Returns the new text and the
 * span to select (so the next F4 changes the same reference), or null when there's nothing to
 * change.
 */
export function cycleReference(
  text: string,
  selStart: number,
  selEnd: number,
): { text: string; start: number; end: number } | null {
  if (!text.startsWith("=")) {
    return null;
  }
  let tokens: Token[];
  try {
    tokens = tokenize(text.slice(1));
  } catch {
    return null;
  }
  // A reference, or a range's two references, with positions in the whole text.
  const units: { parts: Token[]; start: number; end: number }[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.kind !== "ref") {
      continue;
    }
    const colon = tokens[i + 1];
    const next = tokens[i + 2];
    const parts = colon?.kind === "op" && colon.text === ":" && next?.kind === "ref" ? [token, next] : [token];
    const last = parts[parts.length - 1];
    units.push({ parts, start: token.start + 1, end: last.start + last.text.length + 1 });
    i += parts.length === 2 ? 2 : 0;
  }
  const chosen =
    selStart === selEnd
      ? units.filter((u) => u.start <= selStart && selStart <= u.end).slice(0, 1)
      : units.filter((u) => u.start < selEnd && u.end > selStart);
  if (chosen.length === 0) {
    return null;
  }
  const first = /^(\$?)[A-Za-z]+(\$?)/.exec(chosen[0].parts[0].text);
  const [col, row] = [first?.[1] === "$", first?.[2] === "$"];
  // The order Excel steps through: relative, absolute, row locked, column locked.
  const [nextCol, nextRow] =
    !col && !row ? [true, true] : col && row ? [false, true] : !col ? [true, false] : [false, false];
  const restyle = (ref: string) => {
    const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref) as RegExpExecArray;
    return `${nextCol ? "$" : ""}${match[1]}${nextRow ? "$" : ""}${match[2]}`;
  };
  let out = text;
  let shift = 0;
  // Rewrite from the right, so earlier positions stay valid.
  for (const unit of [...chosen].reverse()) {
    for (const part of [...unit.parts].reverse()) {
      const at = part.start + 1;
      const replacement = restyle(part.text);
      out = out.slice(0, at) + replacement + out.slice(at + part.text.length);
      shift += replacement.length - part.text.length;
    }
  }
  // The selection runs from the first reference changed to the end of the last.
  return { text: out, start: chosen[0].start, end: chosen[chosen.length - 1].end + shift };
}

/** Writes for dragging the fill handle from `source` out to `target` (which contains it). */
export function fillWrites(inputs: Inputs, source: Range, target: Range): Inputs {
  const writes: Inputs = {};
  const height = source.r2 - source.r1 + 1;
  const width = source.c2 - source.c1 + 1;
  for (const addr of rangeAddrs(target)) {
    if (inRange(source, addr)) {
      continue;
    }
    const from = {
      col: source.c1 + ((((addr.col - source.c1) % width) + width) % width),
      row: source.r1 + ((((addr.row - source.r1) % height) + height) % height),
    };
    const raw = inputs[addrName(from)] ?? "";
    writes[addrName(addr)] = shiftFormula(raw, addr.col - from.col, addr.row - from.row);
  }
  return writes;
}

export type Clip = { range: Range; cells: Inputs; cut: boolean };

export function copyRange(inputs: Inputs, range: Range, cut = false): Clip {
  const cells: Inputs = {};
  for (const addr of rangeAddrs(range)) {
    cells[addrName(addr)] = inputs[addrName(addr)] ?? "";
  }
  return { range, cells, cut };
}

/** Writes for pasting a clip at `at`; copies shift relative references, cuts move them as-is. */
export function pasteWrites(clip: Clip, at: Addr): Inputs {
  const writes: Inputs = {};
  const dCol = at.col - clip.range.c1;
  const dRow = at.row - clip.range.r1;
  if (clip.cut) {
    for (const key of Object.keys(clip.cells)) {
      writes[key] = "";
    }
  }
  for (const [key, raw] of Object.entries(clip.cells)) {
    const from = parseAddr(key) as Addr;
    const dest = { col: from.col + dCol, row: from.row + dRow };
    if (dest.col < 0 || dest.row < 0) {
      continue;
    }
    writes[addrName(dest)] = clip.cut ? raw : shiftFormula(raw, dCol, dRow);
  }
  return writes;
}

/**
 * Every cell reference in a formula as typed (with any $), and the range it belongs to. A
 * defined name counts as the reference it stands for, written as defined ("$C$2").
 */
export function formulaReferences(formula: string, names: Names = {}): { text: string; range: Range }[] {
  if (!isFormula(formula)) {
    return [];
  }
  let tokens: Token[];
  try {
    tokens = tokenize(formula.slice(1));
  } catch {
    return [];
  }
  const out: { text: string; range: Range }[] = [];
  const defined = new Map(Object.entries(names).map(([name, ref]) => [name.toUpperCase(), ref]));
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    const call = tokens[i + 1]?.kind === "op" && tokens[i + 1].text === "(";
    const ref = token.kind === "name" && !call ? defined.get(token.text.toUpperCase()) : undefined;
    const named = ref ? parseRange(ref.replace(/\$/g, "")) : null;
    if (ref && named) {
      for (const part of ref.split(":")) {
        out.push({ text: part.toUpperCase(), range: named });
      }
      continue;
    }
    const start = token.kind === "ref" ? parseAddr(token.text) : null;
    if (!start) {
      continue;
    }
    const colon = tokens[i + 1];
    const next = tokens[i + 2];
    const end =
      colon?.kind === "op" && colon.text === ":" && next?.kind === "ref" ? parseAddr(next.text) : null;
    if (end) {
      const range = normalizeRange(start, end);
      out.push({ text: token.text.toUpperCase(), range }, { text: next.text.toUpperCase(), range });
      i += 2;
    } else {
      out.push({ text: token.text.toUpperCase(), range: normalizeRange(start, start) });
    }
  }
  return out;
}
