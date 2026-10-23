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
  const press = (key: string, shiftKey = false) => {
    const event = new KeyboardEvent("keydown", { key, shiftKey, bubbles: true, cancelable: true, composed: true });
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
  return { press, type, goTo, editor, active: () => nameBox.value };
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
