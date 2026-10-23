/**
 * The keys a module may use, for flagging misspellings. `parseModule` ignores keys it doesn't
 * know (so a module written for a newer simmy still loads), which means a typo in an optional
 * key like `columnWidth` silently falls back to the default. `simmy validate` warns about them.
 * Keep this in step with parse.ts; test/keys.test.ts checks the hosted modules against it.
 */

/** A mapping's known keys. A nested shape checks that key's value; `true` means any value. */
export type Shape = { [key: string]: Shape | Shape[] | true } | "any-keys";

const titled: Shape = { title: true, text: true };

export const RULE_KEYS: Record<string, string[]> = {
  value: ["cell", "formula"],
  reference: ["cell", "ref", "forms"],
  copied: ["range"],
};

export const MODULE: Shape = {
  version: true,
  workbook: { name: true, sheet: true, names: "any-keys" },
  assignment: { title: true, text: true, section: true },
  sheet: {
    rows: true,
    columns: true,
    columnWidth: true,
    widths: "any-keys",
    cells: "any-keys",
    // parseModule already rejects unknown style keys.
    styles: true,
    outlines: [{ range: true, weight: true, color: true }],
    start: true,
  },
  // Rules are checked by kind, below.
  tasks: [
    {
      id: true,
      name: true,
      text: true,
      details: true,
      hint: true,
      points: true,
      range: true,
      solution: true,
      names: "any-keys",
      rules: true,
    },
  ],
  note: { title: true, steps: true },
  autopilot: [
    { task: true, at: true, type: true, fill: true, doit: true, undo: true, done: true, submit: true, audio: true, transcript: true },
  ],
  poster: { range: true },
  copy: { paneTitle: true, unsupported: { video: titled, book: titled }, submitSteps: true },
};

const isObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Edit distance, to suggest the key that was probably meant. */
function distance(a: string, b: string): number {
  let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const row = [i];
    for (let j = 1; j <= b.length; j++) {
      row[j] = Math.min(prev[j] + 1, row[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = row;
  }
  return prev[b.length];
}

function unknownKey(path: string, key: string, known: string[]): string {
  const close = known
    .map((name) => ({ name, d: name.toLowerCase() === key.toLowerCase() ? 0 : distance(key, name) }))
    .filter(({ d }) => d <= 2)
    .sort((x, y) => x.d - y.d)[0];
  const at = path ? `${path}.${key}` : key;
  return close ? `${at}: unknown key; did you mean "${close.name}"?` : `${at}: unknown key, ignored`;
}

function check(value: unknown, shape: Shape | Shape[] | true, path: string, out: string[]): void {
  if (shape === true || shape === "any-keys") return;
  if (Array.isArray(shape)) {
    if (Array.isArray(value)) value.forEach((item, i) => check(item, shape[0], `${path}[${i}]`, out));
    return;
  }
  // Wrong types are parseModule's to report.
  if (!isObject(value)) return;
  const known = Object.keys(shape);
  for (const [key, child] of Object.entries(value)) {
    if (key in shape) check(child, shape[key], path ? `${path}.${key}` : key, out);
    else out.push(unknownKey(path, key, known));
  }
}

/** Keys in a module that simmy doesn't read, each with its path and a suggestion when one is close. */
export function unknownKeys(data: unknown): string[] {
  const out: string[] = [];
  check(data, MODULE, "", out);
  if (!isObject(data) || !Array.isArray(data.tasks)) return out;
  data.tasks.forEach((task, i) => {
    if (!isObject(task) || !Array.isArray(task.rules)) return;
    task.rules.forEach((rule, j) => {
      // An unknown kind is parseModule's to report.
      if (!isObject(rule) || typeof rule.kind !== "string" || !(rule.kind in RULE_KEYS)) return;
      const known = ["kind", "points", "message", ...RULE_KEYS[rule.kind]];
      for (const key of Object.keys(rule)) {
        if (known.includes(key)) continue;
        const path = `tasks[${i}].rules[${j}]`;
        const kinds = Object.keys(RULE_KEYS).filter((kind) => RULE_KEYS[kind].includes(key));
        out.push(
          kinds.length > 0
            ? `${path}.${key}: only ${kinds.map((kind) => `"${kind}"`).join(" and ")} rules use this, not "${rule.kind}"`
            : unknownKey(path, key, known),
        );
      }
    });
  });
  return out;
}
