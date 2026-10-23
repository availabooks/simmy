import { afterEach, describe, expect, it, vi } from "vitest";
import { revealInBox } from "../src/ui/reveal.ts";

afterEach(() => {
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

/** happy-dom lays nothing out, so give an element the box a browser would. */
function place(el: Element, top: number, height: number) {
  vi.spyOn(el, "getBoundingClientRect").mockReturnValue({ top, height, bottom: top + height } as DOMRect);
}

/**
 * A narration box 100px tall at y=600 on the page, below a 500px viewport, holding a 30px line
 * that sits `lineAt` px down the box's visible area.
 */
function narration(lineAt: number) {
  document.body.innerHTML = `<div class="page"><div class="box" style="overflow-y: auto"><div><p>Line</p></div></div></div>`;
  const box = document.querySelector<HTMLElement>(".box") as HTMLElement;
  const line = document.querySelector<HTMLElement>("p") as HTMLElement;
  vi.spyOn(box, "clientHeight", "get").mockReturnValue(100);
  place(box, 600, 100);
  place(line, 600 + lineAt, 30);
  box.scrollTop = 0;
  const scrollPage = vi.spyOn(window, "scrollTo");
  const scrollBy = vi.spyOn(window, "scrollBy");
  return { box, line, pageScrolled: () => scrollPage.mock.calls.length + scrollBy.mock.calls.length + window.scrollY };
}

describe("revealing the active narration line", () => {
  it("scrolls its box down to a line below the visible area, and leaves the page alone", () => {
    const { box, line, pageScrolled } = narration(150);
    revealInBox(line);
    // The line's bottom (180) lines up with the box's (100).
    expect(box.scrollTop).toBe(80);
    expect(pageScrolled()).toBe(0);
  });

  it("scrolls its box up to a line above the visible area", () => {
    const { box, line } = narration(-40);
    box.scrollTop = 200;
    revealInBox(line);
    expect(box.scrollTop).toBe(160);
  });

  it("leaves a line that is already in view where it is", () => {
    const { box, line } = narration(20);
    revealInBox(line);
    expect(box.scrollTop).toBe(0);
  });

  it("does nothing when no box inside it scrolls, rather than scroll the page", () => {
    document.body.innerHTML = `<div><p>Line</p></div>`;
    const scrollPage = vi.spyOn(window, "scrollTo");
    revealInBox(document.querySelector("p") as HTMLElement);
    expect(scrollPage).not.toHaveBeenCalled();
    expect(document.documentElement.scrollTop).toBe(0);
  });
});
