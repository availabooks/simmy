/**
 * A tiny one-sheet spreadsheet engine for the simulator.
 *
 * It parses and evaluates the formulas a student types, shifts relative
 * references the way Excel does when a formula is filled or pasted, and formats
 * values the way a cell shows them.
 */

export type Addr = { col: number; row: number };
export type Range = { c1: number; r1: number; c2: number; r2: number };

/** Raw cell input keyed by address, exactly as typed ("3.38", "=AVERAGE(D3:D37)"). */
export type Inputs = Record<string, string>;

export type ErrorCode = "#DIV/0!" | "#N/A" | "#NAME?" | "#NUM!" | "#REF!" | "#VALUE!";
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
const NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]*/;
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
          if (!isOp(")")) {
            args.push(parseLevel(0));
            while (isOp(",")) {
              pos++;
              args.push(parseLevel(0));
            }
          }
          expectOp(")");
          return { type: "call", name: upper, args };
        }
        if (upper === "TRUE" || upper === "FALSE") {
          return { type: "bool", value: upper === "TRUE" };
        }
        return { type: "call", name: `#NAME:${upper}`, args: [] };
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

export type FunctionInfo = { name: string; summary: string };

export const FUNCTIONS: FunctionInfo[] = [
  { name: "ABS", summary: "Returns the absolute value of a number." },
  { name: "AVERAGE", summary: "Returns the average (arithmetic mean) of its arguments." },
  { name: "CORREL", summary: "Returns the correlation coefficient between two data sets." },
  { name: "COUNT", summary: "Counts the number of cells in a range that contain numbers." },
  { name: "COUNTA", summary: "Counts the number of cells in a range that are not empty." },
  { name: "COUNTIF", summary: "Counts the number of cells within a range that meet the given condition." },
  {
    name: "IF",
    summary: "Checks whether a condition is met, and returns one value if TRUE, and another if FALSE.",
  },
  { name: "MAX", summary: "Returns the largest value in a set of values." },
  { name: "MEDIAN", summary: "Returns the median, or the number in the middle of the set of given numbers." },
  { name: "MIN", summary: "Returns the smallest number in a set of values." },
  { name: "MODE", summary: "Returns the most frequently occurring value in a range of data." },
  { name: "MODE.SNGL", summary: "Returns the most frequently occurring value in a range of data." },
  { name: "ROUND", summary: "Rounds a number to a specified number of digits." },
  { name: "SQRT", summary: "Returns the square root of a number." },
  { name: "STDEV", summary: "Estimates standard deviation based on a sample." },
  { name: "STDEV.P", summary: "Calculates standard deviation based on the entire population." },
  { name: "STDEV.S", summary: "Estimates standard deviation based on a sample." },
  { name: "SUM", summary: "Adds all the numbers in a range of cells." },
  { name: "VAR", summary: "Estimates variance based on a sample." },
  { name: "VAR.P", summary: "Calculates variance based on the entire population." },
  { name: "VAR.S", summary: "Estimates variance based on a sample." },
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

type Arg = { kind: "scalar"; value: Value } | { kind: "range"; values: Value[] };

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

function toText(value: Value): string {
  if (value === null) {
    return "";
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (typeof value === "number") {
    return String(value);
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
  COUNT: (nums) => nums.length,
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

type Ctx = {
  cell: (addr: Addr) => Value;
};

function evalArg(node: Node, ctx: Ctx): Arg {
  if (node.type === "range") {
    return { kind: "range", values: rangeAddrs(node.range).map((addr) => ctx.cell(addr)) };
  }
  return { kind: "scalar", value: evalNode(node, ctx) };
}

function scalarOf(arg: Arg): Value {
  if (arg.kind === "scalar") {
    return arg.value;
  }
  return arg.values.length === 1 ? arg.values[0] : err("#VALUE!");
}

function callFunction(name: string, nodes: Node[], ctx: Ctx): Value {
  if (name.startsWith("#NAME:") || !FUNCTION_SET.has(name)) {
    return err("#NAME?");
  }
  if (name === "IF") {
    if (nodes.length < 2 || nodes.length > 3) {
      return err("#VALUE!");
    }
    const test = scalarOf(evalArg(nodes[0], ctx));
    if (isError(test)) {
      return test;
    }
    const n = toNumber(test);
    if (isError(n)) {
      return n;
    }
    if (n !== 0) {
      return scalarOf(evalArg(nodes[1], ctx));
    }
    return nodes[2] ? scalarOf(evalArg(nodes[2], ctx)) : false;
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
      return args.reduce(
        (count, arg) =>
          count + (arg.kind === "range" ? arg.values.filter((v) => v !== null && v !== "").length : 1),
        0,
      );
    case "COUNTIF": {
      if (args.length !== 2 || args[0].kind !== "range") {
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
    case "ROUND": {
      if (args.length !== 2) {
        return err("#VALUE!");
      }
      const x = toNumber(scalarOf(args[0]));
      const digits = toNumber(scalarOf(args[1]));
      if (isError(x)) {
        return x;
      }
      if (isError(digits)) {
        return digits;
      }
      const factor = 10 ** Math.trunc(digits);
      return (Math.sign(x) * Math.round(Math.abs(x) * factor + 1e-9)) / factor;
    }
    case "ABS":
    case "SQRT": {
      if (args.length !== 1) {
        return err("#VALUE!");
      }
      const x = toNumber(scalarOf(args[0]));
      if (isError(x)) {
        return x;
      }
      if (name === "ABS") {
        return Math.abs(x);
      }
      return x < 0 ? err("#NUM!") : Math.sqrt(x);
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

function evalNode(node: Node, ctx: Ctx): Value {
  switch (node.type) {
    case "num":
    case "str":
    case "bool":
      return node.value;
    case "ref":
      return ctx.cell(node.addr);
    case "range": {
      const values = rangeAddrs(node.range).map((addr) => ctx.cell(addr));
      return values.length === 1 ? values[0] : err("#VALUE!");
    }
    case "call":
      return callFunction(node.name, node.args, ctx);
    case "unary": {
      const n = toNumber(evalNode(node.arg, ctx));
      if (isError(n)) {
        return n;
      }
      return node.op === "-" ? -n : n;
    }
    case "percent": {
      const n = toNumber(evalNode(node.arg, ctx));
      return isError(n) ? n : n / 100;
    }
    case "binary": {
      const left = evalNode(node.left, ctx);
      if (isError(left)) {
        return left;
      }
      const right = evalNode(node.right, ctx);
      if (isError(right)) {
        return right;
      }
      if (node.op === "&") {
        return toText(left) + toText(right);
      }
      if (["=", "<>", "<", ">", "<=", ">="].includes(node.op)) {
        return compare(left, right, node.op);
      }
      const a = toNumber(left);
      if (isError(a)) {
        return a;
      }
      const b = toNumber(right);
      if (isError(b)) {
        return b;
      }
      switch (node.op) {
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
    }
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

/** Evaluate every cell. Circular references resolve to 0, as Excel shows after its warning. */
export function evaluateSheet(inputs: Inputs): Record<string, Value> {
  const values: Record<string, Value> = {};
  const visiting = new Set<string>();
  const parsed = new Map<string, Node | null>();

  const cell = (addr: Addr): Value => {
    const key = addrName(addr);
    if (key in values) {
      return values[key];
    }
    const raw = inputs[key];
    if (raw === undefined || raw === "") {
      return null;
    }
    if (!isFormula(raw)) {
      const constant = parseConstant(raw);
      values[key] = constant;
      return constant;
    }
    if (visiting.has(key)) {
      return 0;
    }
    visiting.add(key);
    let tree = parsed.get(key);
    if (tree === undefined) {
      try {
        tree = parseFormula(raw);
      } catch {
        tree = null;
      }
      parsed.set(key, tree);
    }
    const value = tree ? evalNode(tree, { cell }) : err("#NAME?");
    visiting.delete(key);
    values[key] = typeof value === "number" && !Number.isFinite(value) ? err("#NUM!") : value;
    return values[key];
  };

  for (const key of Object.keys(inputs)) {
    const addr = parseAddr(key);
    if (addr) {
      cell(addr);
    }
  }
  return values;
}

/** Evaluate a formula as if it were typed into `at`, without touching the sheet. */
export function evaluateFormulaAt(inputs: Inputs, formula: string, at: Addr): Value {
  const key = addrName(at);
  return evaluateSheet({ ...inputs, [key]: formula })[key] ?? null;
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

/** The ranges a formula points at, in order, for the colored reference outlines. */
export function formulaRanges(formula: string): Range[] {
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
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
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

/** What Excel does on Enter: close dangling parentheses and upper-case names and references. */
export function normalizeFormula(formula: string): string {
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
  let out = "";
  let last = 0;
  for (const token of tokens) {
    if (token.kind === "ref" || token.kind === "name") {
      out += body.slice(last, token.start) + token.text.toUpperCase();
      last = token.start + token.text.length;
    }
  }
  body = out + body.slice(last);
  return `=${body}${")".repeat(Math.max(0, open - close))}`;
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

/** Every cell reference in a formula as typed (with any $), and the range it belongs to. */
export function formulaReferences(formula: string): { text: string; range: Range }[] {
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
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
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
