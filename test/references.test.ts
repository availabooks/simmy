import { describe, expect, it } from "vitest";
import { cycleReference } from "../src/engine/sheet.ts";

/** Press F4 with the caret at `caret` (or a selection) and return the new text. */
const f4 = (text: string, start = text.length, end = start) => cycleReference(text, start, end)?.text;

describe("cycleReference (F4)", () => {
  it("steps a reference through Excel's four forms and back", () => {
    const forms = ["=A1"];
    for (let i = 0; i < 4; i++) {
      forms.push(f4(forms[forms.length - 1]) as string);
    }
    expect(forms).toEqual(["=A1", "=$A$1", "=A$1", "=$A1", "=A1"]);
  });

  it("changes the reference the caret is in or right after", () => {
    expect(f4("=A1+B2", 2)).toBe("=$A$1+B2");
    expect(f4("=A1+B2", 3)).toBe("=$A$1+B2");
    expect(f4("=A1+B2", 1)).toBe("=$A$1+B2");
    expect(f4("=A1+B2")).toBe("=A1+$B$2");
  });

  it("changes both ends of a range together", () => {
    expect(f4("=SUM(A1:B2)", 8)).toBe("=SUM($A$1:$B$2)");
    expect(f4("=SUM($A$1:B2)", 6)).toBe("=SUM(A$1:B$2)");
  });

  it("changes every reference a selection touches", () => {
    expect(f4("=A1*B2+C3", 1, 6)).toBe("=$A$1*$B$2+C3");
  });

  it("leaves text in quotes, names and plain text alone", () => {
    expect(cycleReference('="A1"', 3, 3)).toBeNull();
    expect(cycleReference("=SUM(", 5, 5)).toBeNull();
    expect(cycleReference("A1", 2, 2)).toBeNull();
  });

  it("returns where the changed reference now ends, for the caret", () => {
    expect(cycleReference("=A1+B2", 3, 3)).toEqual({ text: "=$A$1+B2", start: 1, end: 5 });
    expect(cycleReference("=SUM(A1:B2)", 8, 8)).toEqual({ text: "=SUM($A$1:$B$2)", start: 5, end: 14 });
  });
});
