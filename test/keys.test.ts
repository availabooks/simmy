import { readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { parse } from "yaml";
import { unknownKeys } from "../src/module/keys.ts";
import { moduleText } from "./fixture.ts";

describe("unknownKeys", () => {
  it("finds nothing in the hosted modules", () => {
    for (const name of readdirSync("hosted-modules")) {
      expect(unknownKeys(parse(moduleText(name))), name).toEqual([]);
    }
  });

  it("suggests the key that was probably meant", () => {
    const data = parse(moduleText());
    data.sheet.colunmWidth = 80;
    data.assignment.Section = "One";
    data.autopilot[0].transcrpt = "Hello";
    expect(unknownKeys(data)).toEqual([
      'assignment.Section: unknown key; did you mean "section"?',
      'sheet.colunmWidth: unknown key; did you mean "columnWidth"?',
      'autopilot[0].transcrpt: unknown key; did you mean "transcript"?',
    ]);
  });

  it("flags a key with no close match", () => {
    expect(unknownKeys({ theme: "dark", copy: { unsupported: { video: { link: "x" } } } })).toEqual([
      "theme: unknown key, ignored",
      "copy.unsupported.video.link: unknown key, ignored",
    ]);
  });

  it("checks rule keys against the rule's kind", () => {
    const rules = [
      { kind: "copied", points: 1, message: "m", range: "D6:D12", cell: "D6" },
      { kind: "value", points: 1, message: "m", cell: "D6", formula: "=1", ref: "B2" },
      { kind: "reference", points: 1, message: "m", cell: "D6", ref: "B2", froms: "x" },
    ];
    expect(unknownKeys({ tasks: [{ rules }] })).toEqual([
      'tasks[0].rules[0].cell: only "value" and "reference" rules use this, not "copied"',
      'tasks[0].rules[1].ref: only "reference" rules use this, not "value"',
      'tasks[0].rules[2].froms: unknown key; did you mean "forms"?',
    ]);
  });

  it("leaves cell addresses, column widths and wrong types alone", () => {
    expect(unknownKeys({ sheet: { cells: { Q7: 1 }, widths: { B: 90 } }, tasks: "nope", note: [1] })).toEqual([]);
  });
});
