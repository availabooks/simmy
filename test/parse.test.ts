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
    expect(mod.workbook).toEqual({
      name: "References and Calculations",
      sheet: "Referencing Cells",
      names: {},
    });
    expect(mod.sheet.cells.C2).toBe("0.3");
    expect(mod.sheet.widths.slice(0, 6)).toEqual([12, 74, 54, 80, 80, 64]);
    expect(mod.sheet.widths).toHaveLength(20);
    expect(mod.sheet.styles.D6).toEqual({ format: "accounting", fill: "#f0e68c", lined: true });
    expect(mod.sheet.styles.B5).toMatchObject({ fill: "#800000", color: "#ffffff", bold: true });
    expect(mod.sheet.outlines).toHaveLength(2);
    expect(mod.tasks.map((t) => t.id)).toEqual(["A", "B", "C", "D"]);
    expect(mod.totalPoints).toBe(100);
    expect(mod.autopilot[0]).toMatchObject({
      kind: "hold",
      audio: "1.mp3",
      transcript:
        "Peter's Pool Palace is having an end-of-season pool sale. All pools and pool supplies are 30% off, and you have decided to buy a pool.",
    });
    expect(mod.autopilot[4]).toEqual({
      taskId: "A",
      kind: "type",
      cell: "D6",
      formula: "=C6*C$2",
      fill: undefined,
      finish: false,
      audio: "5.mp3",
      transcript:
        "Multiply the regular price in cell C6 by the store discount in cell C$2. The dollar sign keeps C2 fixed when the formula is copied.",
    });
    expect(mod.autopilot.at(-1)).toEqual({
      kind: "submit",
      audio: "16.mp3",
      transcript: "Press Next, then Submit Now to see your grade.",
    });
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
      'tasks[0].rules[0].kind: must be one of "value", "reference", "copied", "equals", "contains", "name", "between", "selected"',
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

  it("reads a clip and transcript on each autopilot entry", () => {
    const mod = parseModule({
      ...minimal,
      autopilot: [
        { audio: "/media/1.mp3", transcript: "Let's calculate the discount for each row." },
        { task: "A", audio: "/media/2.mp3", transcript: "Open the task." },
        {
          task: "A",
          type: "=A1+A2",
          at: "A3",
          done: false,
          audio: "/media/3.mp3",
          transcript: "Type the total in A3.",
        },
        { submit: true, audio: "/media/4.mp3", transcript: "Press Next, then Submit Now." },
      ],
    });
    expect(mod.autopilot).toEqual([
      { kind: "hold", cell: undefined, audio: "/media/1.mp3", transcript: "Let's calculate the discount for each row." },
      { taskId: "A", kind: "open", audio: "/media/2.mp3", transcript: "Open the task." },
      {
        taskId: "A",
        kind: "type",
        cell: "A3",
        formula: "=A1+A2",
        fill: undefined,
        finish: false,
        audio: "/media/3.mp3",
        transcript: "Type the total in A3.",
      },
      { kind: "submit", audio: "/media/4.mp3", transcript: "Press Next, then Submit Now." },
    ]);
  });

  it("rejects a narrated autopilot that is missing clips", () => {
    expect(
      problems({
        ...minimal,
        autopilot: [
          { audio: "/media/1.mp3", transcript: "Hello." },
          { task: "A", type: "=A1+A2" },
        ],
      }),
    ).toContain("autopilot[1]: needs an audio file and a transcript");
    expect(problems({ ...minimal, audio: "/media/one.mp3" })).toContain(
      "audio: give each autopilot entry its own audio file",
    );
    expect(problems({ ...minimal, autopilot: [{ audio: "/media/1.mp3" }] })).toContain(
      "autopilot[0].transcript: must be text",
    );
    expect(problems({ ...minimal, autopilot: [{}] })).toContain(
      "autopilot[0]: needs a task, or an audio file and a transcript",
    );
  });
});

describe("names in a module", () => {
  const withNames = (extra: object, task: object = {}) => ({
    ...minimal,
    ...extra,
    tasks: [{ ...minimal.tasks[0], ...task }],
  });

  it("stores workbook and task names as absolute ranges", () => {
    const mod = parseModule(
      withNames({ workbook: { name: "Book1", names: { Rate: "a2" } } }, { names: { Both: "A1:A2" } }),
    );
    expect(mod.workbook.names).toEqual({ Rate: "$A$2" });
    expect(mod.tasks[0].names).toEqual({ Both: "$A$1:$A$2" });
  });

  it("lets a task only define names, opening on the first one", () => {
    const mod = parseModule({
      ...minimal,
      tasks: [{ text: "Name A2 Rate.", names: { Rate: "A2" }, rules: [] }],
    });
    expect(mod.tasks[0]).toMatchObject({ range: undefined, solution: undefined, hint: "A2" });
    expect(mod.autopilot).toEqual([{ taskId: "A", kind: "doit" }]);
  });

  it("reports bad names, duplicates and ranges off the sheet, with their paths", () => {
    const names = { Q1: "A1", rate: "A1", RATE: "A2", Far: "ZZ1" };
    expect(problems(withNames({ workbook: { name: "B", names } }))).toEqual([
      'workbook.names.Q1: "Q1" looks like a cell reference, so it can\'t be a name',
      'workbook.names: "rate" and "RATE" are the same name',
      "workbook.names.Far: ZZ1 is outside the sheet (20 columns, 40 rows)",
    ]);
  });

  it("needs a range with a solution, and something for Do it", () => {
    expect(problems(withNames({}, { solution: undefined }))).toContain(
      "tasks[0].solution: is needed with range",
    );
    expect(problems(withNames({}, { range: undefined, solution: undefined }))).toContain(
      "tasks[0]: needs a range and solution, or names, for Do it",
    );
  });

  it("reports a formula using a name no one defines, or a function simmy lacks", () => {
    expect(problems(withNames({}, { solution: "=A1*Rate" }))).toContain(
      'tasks[0].solution: "=A1*Rate" uses Rate, which no name in workbook.names or a task\'s names defines',
    );
    expect(problems(withNames({}, { solution: "=VLOOKUP(A1,A1:A2,1)" }))).toContain(
      'tasks[0].solution: "=VLOOKUP(A1,A1:A2,1)" uses VLOOKUP, which this simulator doesn\'t have',
    );
    const defined = withNames({ workbook: { name: "B", names: { Rate: "A2" } } }, { solution: "=A1*rate" });
    expect(problems(defined)).toEqual([]);
  });
});
