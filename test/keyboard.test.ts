/**
 * Typing in a cell and moving with the keyboard, as in Excel: Enter, Tab and the arrows confirm
 * the entry and move one cell, and a typed character goes in at the cursor.
 *
 * The cell editor sits inside the grid, so a key pressed while typing reaches the editor's
 * listener and then the grid's. For real input, a browser runs microtasks between one listener
 * and the next, so Preact has already re-rendered (the edit may be over) when the grid's
 * listener runs. A script's dispatchEvent doesn't do that, so these tests render synchronously
 * instead, which gives the grid's listener the same view it gets in a browser.
 */
import { options } from "preact";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { parse } from "yaml";
import { defineSimmy, type SimmyElement } from "../src/element.ts";
import { moduleText } from "./fixture.ts";

defineSimmy();

const settle = () => new Promise((resolve) => setTimeout(resolve, 50));

let debounce: typeof options.debounceRendering;
beforeEach(() => {
  debounce = options.debounceRendering;
  options.debounceRendering = (render) => render();
});
afterEach(() => {
  options.debounceRendering = debounce;
  document.body.innerHTML = "";
});

/** A simulator with no demo running, and helpers to press keys and read the active cell. */
async function sheet() {
  const el = document.createElement("simmy-app") as SimmyElement;
  el.setAttribute("autoplay", "false");
  el.setAttribute("fonts", "none");
  document.body.append(el);
  el.module = parse(moduleText());
  await settle();
  const shadow = el.shadowRoot as ShadowRoot;
  const grid = shadow.querySelector(".xd-grid-scroll") as HTMLElement;
  const nameBox = shadow.querySelector('input[aria-label="Name Box"]') as HTMLInputElement;
  const editor = () => shadow.querySelector(".xd-cell-editor") as HTMLInputElement | null;

  /**
   * Press a key where the focus would be: the cell editor while typing, else the grid. Like a
   * browser, a character that no listener cancels goes into the editor at its cursor.
   */
  const press = (key: string, shiftKey = false, ctrlKey = false) => {
    const event = new KeyboardEvent("keydown", {
      key,
      shiftKey,
      ctrlKey,
      bubbles: true,
      cancelable: true,
      composed: true,
    });
    const input = editor();
    (input ?? grid).dispatchEvent(event);
    if (input && key.length === 1 && !event.defaultPrevented) {
      const start = input.selectionStart ?? input.value.length;
      const end = input.selectionEnd ?? start;
      input.value = input.value.slice(0, start) + key + input.value.slice(end);
      input.setSelectionRange(start + 1, start + 1);
      input.dispatchEvent(new Event("input", { bubbles: true }));
    }
  };
  const type = (text: string) => {
    for (const char of text) press(char);
  };
  const goTo = (addr: string) => {
    nameBox.value = addr;
    nameBox.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
  };
  /** What a cell shows, read from the grid. */
  const cellText = (addr: string) => {
    const col = addr.charCodeAt(0) - 65;
    const row = Number(addr.slice(1)) - 1;
    const line = shadow.querySelectorAll(".xd-row-line")[row];
    return line?.querySelectorAll(".xd-cell")[col]?.textContent ?? "";
  };
  const alert = () => shadow.querySelector(".xd-alert")?.textContent ?? null;
  const suggestions = () => [...shadow.querySelectorAll(".xd-suggest li")].map((li) => li.textContent);
  const dialog = () => shadow.querySelector<HTMLFormElement>(".xd-dialog");
  /** Fill a dialog's field by its label text. */
  const fill = (label: string, value: string) => {
    const field = [...(dialog()?.querySelectorAll("label") ?? [])]
      .find((l) => l.textContent?.trim().startsWith(label))
      ?.querySelector("input") as HTMLInputElement;
    field.value = value;
    field.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const ok = () => dialog()?.querySelector<HTMLButtonElement>(".xd-dialog-btn.is-default")?.click();
  return {
    press,
    type,
    goTo,
    editor,
    cellText,
    alert,
    suggestions,
    dialog,
    fill,
    ok,
    grid,
    active: () => nameBox.value,
  };
}

describe("keyboard movement", () => {
  it("moves one cell for Enter, Tab and the arrows when not typing", async () => {
    const s = await sheet();
    s.goTo("H8");
    expect(s.active()).toBe("H8");
    s.press("Enter");
    expect(s.active()).toBe("H9");
    s.press("Tab");
    expect(s.active()).toBe("I9");
    s.press("ArrowDown");
    expect(s.active()).toBe("I10");
  });

  it.each([
    ["Enter", false, "H9"],
    ["Enter", true, "H7"],
    ["Tab", false, "I8"],
    ["Tab", true, "G8"],
    ["ArrowDown", false, "H9"],
    ["ArrowRight", false, "I8"],
  ])("confirms the entry and moves one cell for %s (shift %s) after typing", async (key, shift, expected) => {
    const s = await sheet();
    s.goTo("H8");
    s.type("57");
    expect(s.editor()?.value).toBe("57");
    s.press(key, shift);
    expect(s.active()).toBe(expected);
    s.goTo("H8");
    s.press("F2");
    expect(s.editor()?.value).toBe("57");
  });

  it("cancels the entry with Escape and stays on the cell", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.type("57");
    s.press("Escape");
    expect(s.editor()).toBeNull();
    expect(s.active()).toBe("H8");
    s.press("F2");
    expect(s.editor()?.value).toBe("");
  });
});

describe("typing in the cell editor", () => {
  it("puts a character at the cursor, not at the end", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.type("=SUM(A1:A3)");
    s.press("Enter");
    s.goTo("H8");
    s.press("F2");
    // The cursor just after "A1", as four presses of the left arrow would leave it.
    s.editor()?.setSelectionRange(7, 7);
    s.press("x");
    expect(s.editor()?.value).toBe("=SUM(A1x:A3)");
  });

  it("replaces selected text with what is typed", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.type("123");
    s.editor()?.setSelectionRange(0, 3);
    s.press("9");
    expect(s.editor()?.value).toBe("9");
  });
});

describe("Tab then Enter", () => {
  it("returns to the column the Tabs started from, on the next row", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.type("1");
    s.press("Tab");
    s.type("2");
    s.press("Tab");
    s.type("3");
    expect(s.active()).toBe("J8");
    s.press("Enter");
    expect(s.active()).toBe("H9");
  });

  it("does the same when moving with Tab without typing", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.press("Tab");
    s.press("Tab");
    s.press("Enter");
    expect(s.active()).toBe("H9");
  });

  it("goes straight down once another move has ended the run", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.type("1");
    s.press("Tab");
    s.press("ArrowDown");
    expect(s.active()).toBe("I9");
    s.press("Enter");
    expect(s.active()).toBe("I10");
  });

  it("starts a new run after Enter", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.press("Tab");
    s.press("Enter");
    expect(s.active()).toBe("H9");
    s.press("Enter");
    expect(s.active()).toBe("H10");
  });
});

describe("F4 while typing a formula", () => {
  it("steps the reference before the caret through $A$1, A$1 and $A1", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.type("=C2*D2");
    const forms = [];
    for (let i = 0; i < 4; i++) {
      s.press("F4");
      forms.push(s.editor()?.value);
    }
    expect(forms).toEqual(["=C2*$D$2", "=C2*D$2", "=C2*$D2", "=C2*D2"]);
  });

  it("keeps the $ when the formula is entered", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.type("=C2");
    s.press("F4");
    s.press("Enter");
    s.goTo("H8");
    s.press("F2");
    expect(s.editor()?.value).toBe("=$C$2");
  });
});

describe("Calculate Now (F9)", () => {
  it("picks new RANDBETWEEN numbers, and only when asked", async () => {
    const s = await sheet();
    s.goTo("H8");
    // Up to seven digits, which the column is wide enough to show.
    s.type("=RANDBETWEEN(1,1000000)");
    s.press("Enter");
    const cell = () => s.cellText("H8");
    const first = cell();
    expect(Number(first)).toBeGreaterThan(0);
    s.goTo("H8");
    expect(cell()).toBe(first);
    s.press("F9");
    expect(cell()).not.toBe(first);
  });
});

describe("the Name Box", () => {
  it("names the selection, then shows the name and goes to it", async () => {
    const s = await sheet();
    s.goTo("C2");
    s.goTo("Rate");
    expect(s.active()).toBe("Rate");
    s.goTo("H8");
    expect(s.active()).toBe("H8");
    s.goTo("rate");
    expect(s.active()).toBe("Rate");
  });

  it("lets formulas use the name, suggesting it as it's typed", async () => {
    const s = await sheet();
    s.goTo("C2");
    s.goTo("Rate");
    s.goTo("H8");
    s.type("=Ra");
    expect(s.suggestions()).toContain("▭Rate");
    s.type("te*100");
    s.press("Enter");
    // C2 holds the 30% discount.
    expect(s.cellText("H8")).toBe("30");
  });

  it("refuses an entry that is neither a reference nor a valid name", async () => {
    const s = await sheet();
    s.goTo("Tax Rate");
    expect(s.alert()).toContain("You must enter a valid reference you want to go to");
  });

  it("can undo naming", async () => {
    const s = await sheet();
    s.goTo("C2");
    s.goTo("Rate");
    s.press("z", false, true);
    s.goTo("C2");
    expect(s.active()).toBe("C2");
  });
});

describe("Go To", () => {
  it("opens with F5 or Ctrl+G and goes to a reference", async () => {
    const s = await sheet();
    s.press("F5");
    expect(s.dialog()?.getAttribute("aria-label")).toBe("Go To");
    s.fill("Reference", "J20");
    s.ok();
    expect(s.dialog()).toBeNull();
    expect(s.active()).toBe("J20");
    s.press("g", false, true);
    expect(s.dialog()).not.toBeNull();
  });

  it("lists the names and goes to one", async () => {
    const s = await sheet();
    s.goTo("C2");
    s.goTo("Rate");
    s.goTo("H8");
    s.press("F5");
    expect(s.dialog()?.querySelector(".xd-listbox")?.textContent).toBe("Rate");
    s.fill("Reference", "Rate");
    s.ok();
    expect(s.active()).toBe("Rate");
  });

  it("goes to a whole spill with A1#", async () => {
    const s = await sheet();
    s.goTo("H8");
    s.type("=SEQUENCE(3)");
    s.press("Enter");
    s.press("F5");
    s.fill("Reference", "H8#");
    s.ok();
    // The selection box is drawn over three rows of 20px (plus the gridline).
    const box = s.grid.querySelector<HTMLElement>(".xd-selbox");
    expect(box?.style.height).toBe("61px");
  });

  it("says when a reference isn't valid, and closes on Escape without the grid seeing it", async () => {
    const s = await sheet();
    s.press("F5");
    s.fill("Reference", "Nowhere");
    s.ok();
    expect(s.dialog()?.textContent).toContain("Reference isn't valid.");
    s.dialog()?.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, composed: true }));
    expect(s.dialog()).toBeNull();
  });
});

describe("the name dialogs", () => {
  /** Click a dialog button by its text, in the top dialog. */
  const click = (s: Awaited<ReturnType<typeof sheet>>, text: string) => {
    const dialogs = s.grid.getRootNode() as ShadowRoot;
    const all = [...dialogs.querySelectorAll(".xd-dialog")];
    const top = all[all.length - 1];
    [...top.querySelectorAll("button")].find((b) => b.textContent?.trim() === text)?.click();
  };
  const topDialog = (s: Awaited<ReturnType<typeof sheet>>) => {
    const all = [...(s.grid.getRootNode() as ShadowRoot).querySelectorAll(".xd-dialog")];
    return all[all.length - 1] ?? null;
  };

  it("defines a name with New Name from the Name Manager (Ctrl+F3)", async () => {
    const s = await sheet();
    s.goTo("C3");
    s.press("F3", false, true);
    expect(topDialog(s)?.getAttribute("aria-label")).toBe("Name Manager");
    click(s, "New…");
    expect(topDialog(s)?.getAttribute("aria-label")).toBe("New Name");
    // The label to the left, "Tax Rate", is suggested as Tax_Rate.
    const name = topDialog(s)?.querySelector("input") as HTMLInputElement;
    expect(name.value).toBe("Tax_Rate");
    click(s, "OK");
    expect(topDialog(s)?.getAttribute("aria-label")).toBe("Name Manager");
    // A sheet name with a space is quoted, as Excel writes it.
    expect(topDialog(s)?.textContent).toContain("='Referencing Cells'!$C$3");
    click(s, "Close");
    s.goTo("A1");
    s.goTo("tax_rate");
    expect(s.active()).toBe("Tax_Rate");
  });

  it("creates names from a table's labels (Ctrl+Shift+F3)", async () => {
    const s = await sheet();
    s.goTo("B5:C12");
    s.press("F3", true, true);
    const boxes = [...(topDialog(s)?.querySelectorAll<HTMLInputElement>("input[type=checkbox]") ?? [])];
    expect(boxes.map((box) => box.checked)).toEqual([true, true, false, false]);
    click(s, "OK");
    s.goTo("Price");
    expect(s.active()).toBe("Price");
    s.goTo("Pool");
    expect(s.active()).toBe("Pool");
  });

  it("renames with Edit and deletes after asking", async () => {
    const s = await sheet();
    s.goTo("C2");
    s.goTo("Rate");
    s.press("F3", false, true);
    click(s, "Edit…");
    const name = topDialog(s)?.querySelector("input") as HTMLInputElement;
    name.value = "Discount";
    name.dispatchEvent(new Event("input", { bubbles: true }));
    click(s, "OK");
    expect(topDialog(s)?.textContent).toContain("Discount");
    expect(topDialog(s)?.textContent).not.toContain("Rate");
    click(s, "Delete");
    expect(topDialog(s)?.textContent).toContain("Are you sure you want to delete the name Discount?");
    click(s, "OK");
    expect(topDialog(s)?.querySelectorAll(".xd-table-row:not(.is-head)").length).toBe(0);
  });

  it("pastes a name into the formula being typed (F3)", async () => {
    const s = await sheet();
    s.goTo("C2");
    s.goTo("Rate");
    s.goTo("H8");
    s.type("=100*");
    s.press("F3");
    expect(topDialog(s)?.getAttribute("aria-label")).toBe("Paste Name");
    click(s, "OK");
    expect(s.editor()?.value).toBe("=100*Rate");
  });
});
