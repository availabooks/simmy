import { afterEach, describe, expect, it, vi } from "vitest";
import { defineSimmy, type SimmyElement } from "../src/element.ts";
import { moduleText } from "./fixture.ts";

defineSimmy();

const settle = () => new Promise((resolve) => setTimeout(resolve, 50));

/** Wait for the element to load its module and render. */
async function ready(el: SimmyElement): Promise<void> {
  for (let i = 0; i < 40 && !el.module && !el.shadowRoot?.querySelector(".xd-error"); i++) {
    await settle();
  }
}

afterEach(() => {
  document.body.innerHTML = "";
  vi.unstubAllGlobals();
});

describe("<simmy-app>", () => {
  it("renders inline YAML into its shadow root", async () => {
    const el = document.createElement("simmy-app");
    el.setAttribute("autoplay", "false");
    el.setAttribute("fonts", "none");
    const script = document.createElement("script");
    script.type = "text/yaml";
    // Indented, the way it sits in a page.
    script.textContent = moduleText()
      .split("\n")
      .map((line) => `      ${line}`)
      .join("\n");
    el.append(script);
    document.body.append(el);
    await ready(el);

    expect(el.module?.workbook.sheet).toBe("Referencing Cells");
    const shadow = el.shadowRoot as ShadowRoot;
    expect(shadow.querySelector(".xd-pane-title")?.textContent).toBe("Excel Educator");
    expect(shadow.textContent).toContain("Pool Cover");
    // The page's own DOM stays untouched.
    expect(document.body.querySelector(".xd")).toBeNull();
  });

  it("fetches a module from src", async () => {
    const fetch = vi.fn(async () => new Response(moduleText()));
    vi.stubGlobal("fetch", fetch);
    const el = document.createElement("simmy-app");
    el.setAttribute("fonts", "none");
    el.setAttribute("autoplay", "false");
    el.setAttribute("src", "/hosted-modules/references-and-calculations/config.yaml");
    document.body.append(el);
    await ready(el);

    expect(fetch).toHaveBeenCalledOnce();
    expect(String((fetch.mock.calls[0] as unknown[])[0])).toMatch(/\/hosted-modules\/references-and-calculations\/config\.yaml$/);
    expect(el.module?.tasks).toHaveLength(4);
    expect(el.module?.autopilot[0]).toMatchObject({
      audio: new URL("/hosted-modules/references-and-calculations/audio/1.mp3", document.baseURI).href,
    });
  });

  it("shows the problems and fires simmy:error for a bad module", async () => {
    const el = document.createElement("simmy-app");
    el.setAttribute("fonts", "none");
    const errors: Event[] = [];
    el.addEventListener("simmy:error", (event) => errors.push(event));
    document.body.append(el);
    el.module = { workbook: { name: "B" }, assignment: { title: "T", text: "X" }, tasks: [] };
    await settle();

    expect(errors).toHaveLength(1);
    const box = el.shadowRoot?.querySelector(".xd-error");
    expect(box?.textContent).toContain("tasks: needs at least one task");
  });

  it("takes a module object set from script", async () => {
    const el = document.createElement("simmy-app");
    el.setAttribute("fonts", "none");
    el.setAttribute("autoplay", "false");
    const seen: unknown[] = [];
    el.addEventListener("simmy:ready", (event) => seen.push((event as CustomEvent).detail));
    document.body.append(el);
    el.module = {
      workbook: { name: "Book1" },
      assignment: { title: "Totals", text: "Add them." },
      sheet: { cells: { A1: 2, A2: 3 } },
      tasks: [{ text: "Total in A3.", range: "A3", solution: "=A1+A2", rules: [] }],
    };
    await settle();

    expect(seen).toHaveLength(1);
    expect(el.shadowRoot?.textContent).toContain("Total in A3.");
  });

  it("projects a slotted heading", async () => {
    const el = document.createElement("simmy-app");
    el.setAttribute("fonts", "none");
    el.setAttribute("autoplay", "false");
    el.innerHTML = '<h2 slot="head">Try it here</h2>';
    document.body.append(el);
    el.module = {
      workbook: { name: "Book1" },
      assignment: { title: "T", text: "X" },
      tasks: [{ text: "One", range: "A1", solution: "=1", rules: [] }],
    };
    await settle();
    expect(el.shadowRoot?.querySelector('slot[name="head"]')).not.toBeNull();
  });

  it("opens in movie mode behind a Play screen, and plays a kiosk without one", async () => {
    const el = document.createElement("simmy-app");
    el.setAttribute("fonts", "none");
    document.body.append(el);
    el.module = {
      workbook: { name: "Book1" },
      assignment: { title: "T", text: "X" },
      tasks: [{ text: "One", range: "A1", solution: "=1", rules: [] }],
    };
    await settle();
    const shadow = el.shadowRoot as ShadowRoot;
    expect(shadow.querySelector(".xd-intro")).not.toBeNull();

    el.setAttribute("mode", "kiosk");
    await settle();
    expect(shadow.querySelector(".xd-intro")).toBeNull();
  });

  it("takes the web platform from the attribute", async () => {
    const el = document.createElement("simmy-app");
    el.setAttribute("fonts", "none");
    el.setAttribute("platform", "web");
    document.body.append(el);
    el.module = {
      workbook: { name: "Book1" },
      assignment: { title: "T", text: "X" },
      tasks: [{ text: "One", range: "A1", solution: "=1", rules: [] }],
    };
    await settle();
    const shadow = el.shadowRoot as ShadowRoot;
    expect(shadow.querySelector(".xd")?.getAttribute("data-platform")).toBe("web");
  });

  it("measures a sticky header for chrome-offset", async () => {
    const header = document.createElement("header");
    header.className = "site-header";
    Object.defineProperty(header, "offsetHeight", { value: 72 });
    document.body.append(header);
    const el = document.createElement("simmy-app");
    el.setAttribute("fonts", "none");
    el.setAttribute("chrome-offset", ".site-header");
    document.body.append(el);
    expect(el.style.getPropertyValue("--chrome-h")).toBe("72px");
    el.setAttribute("chrome-offset", "40");
    expect(el.style.getPropertyValue("--chrome-h")).toBe("40px");
  });
});
