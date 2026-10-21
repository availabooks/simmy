import { parse } from "yaml";
import { describe, expect, it } from "vitest";
import { loadModule, ModuleError, parseModule } from "../src/module/parse.ts";
import { moduleText } from "./fixture.ts";

const minimal = {
  workbook: { name: "Book1" },
  assignment: { title: "Totals", text: "Add the numbers." },
  sheet: { cells: { A1: 2, A2: 3 } },
  tasks: [
    {
      text: "Total A1 and A2 in A3.",
      range: "A3",
      solution: "=A1+A2",
      rules: [{ kind: "value", points: 10, message: "A3 is not the total", cell: "A3", formula: "=A1+A2" }],
    },
  ],
};

const problems = (data: unknown): string[] => {
  try {
    parseModule(data);
  } catch (error) {
    if (error instanceof ModuleError) {
      return error.problems;
    }
    throw error;
  }
  return [];
};

describe("parseModule", () => {
  it("reads the shipped module", async () => {
    const mod = await loadModule(moduleText());
    expect(mod.workbook).toEqual({ name: "References and Calculations", sheet: "Referencing Cells" });
    expect(mod.sheet.cells.C2).toBe("0.3");
    expect(mod.sheet.widths.slice(0, 6)).toEqual([12, 74, 54, 80, 80, 64]);
    expect(mod.sheet.widths).toHaveLength(20);
    expect(mod.sheet.styles.D6).toEqual({ format: "accounting", fill: "#f0e68c", lined: true });
    expect(mod.sheet.styles.B5).toMatchObject({ fill: "#800000", color: "#ffffff", bold: true });
    expect(mod.sheet.outlines).toHaveLength(2);
    expect(mod.tasks.map((t) => t.id)).toEqual(["A", "B", "C", "D"]);
    expect(mod.totalPoints).toBe(100);
    expect(mod.autopilot[0]).toEqual({
      taskId: "A",
      kind: "type",
      cell: "D6",
      formula: "=C6*C$2",
      fill: "D6:D12",
    });
    expect(mod.autopilot[1]).toEqual({ taskId: "B", kind: "doit", showUndo: true });
  });

  it("fills in defaults for a small module", () => {
    const mod = parseModule(minimal);
    expect(mod.workbook.sheet).toBe("Sheet1");
    expect(mod.assignment.section).toBe("Assessment Section 1");
    expect(mod.tasks[0]).toMatchObject({ id: "A", name: "A", hint: "A3", points: 10, details: [] });
    expect(mod.sheet.start).toBe("A3");
    expect(mod.note.title).toBe("Try it");
    expect(mod.autopilot).toEqual([{ taskId: "A", kind: "type", cell: "A3", formula: "=A1+A2", fill: undefined }]);
    expect(mod.copy.paneTitle).toBe("Excel Educator");
    expect(mod.poster).toEqual({ c1: 0, r1: 0, c2: 0, r2: 1 });
  });

  it("accepts JSON text as well as YAML", async () => {
    const mod = await loadModule(JSON.stringify(minimal));
    expect(mod.tasks).toHaveLength(1);
  });

  it("lists every problem with the path to it", () => {
    const bad = parse(`
workbook: {}
assignment: { title: T, text: X }
sheet:
  columns: 5
  cells: { A1: 1, Z99: 2, B1: "=SUM(" }
  styles:
    - { range: A1, format: currency, wobble: true }
tasks:
  - { id: A, text: One, range: A2, solution: A1*2, rules: [{ kind: guess, points: 1, message: m }] }
  - { id: A, text: Two, range: A3, solution: =A1, rules: [] }
autopilot:
  - { task: Q, doit: true }
`);
    expect(problems(bad)).toEqual([
      "workbook.name: must be text",
      "sheet.cells.Z99: Z99 is outside the sheet (5 columns, 40 rows)",
      expect.stringMatching(/^sheet\.cells\.B1: "=SUM\(" is not a formula Excel would accept/),
      expect.stringMatching(/^sheet\.styles\[0\]\.format: must be one of/),
      "sheet.styles[0].wobble: is not a style this simulator knows",
      'tasks[0].rules[0].kind: must be "value", "reference" or "copied"',
      'tasks[0].solution: "A1*2" must start with =',
      'tasks[1].id: "A" is used by another task',
      'autopilot[0].task: no task has id "Q"',
    ]);
  });

  it("reports YAML syntax errors as module errors", async () => {
    await expect(loadModule("tasks: [unclosed")).rejects.toBeInstanceOf(ModuleError);
  });

  it("needs at least one task", () => {
    expect(problems({ ...minimal, tasks: [] })).toContain("tasks: needs at least one task");
  });
});
