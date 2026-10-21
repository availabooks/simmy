/**
 * <simmy-app>: the simulator as a custom element. It renders into a shadow root,
 * so the host page's styles don't reach in and the simulator's don't leak out.
 *
 *   <simmy-app src="/hosted-modules/references-and-calculations/config.yaml"></simmy-app>
 *   <simmy-app><script type="text/yaml">…</script></simmy-app>
 *   el.module = { …already-parsed object… }
 *
 * Attributes: src, platform (mac | win | web | auto), mode (movie | kiosk; movie waits for
 * Play and stops at the end, kiosk plays muted and loops), ribbon (open | collapsed),
 * autoplay ("false" starts in control), history, fit-screen, chrome-offset (a CSS length,
 * or a selector for a sticky header to measure), fonts ("none" skips loading Carlito and Inter).
 *
 * Events: simmy:ready (detail: module), simmy:graded (detail: grade
 * report), simmy:error (detail: error).
 */
import { h, render } from "preact";
import type { GradeReport } from "./module/grade.ts";
import { resolveModuleAudio } from "./module/audio.ts";
import { loadModule, ModuleError, parseModule } from "./module/parse.ts";
import type { SimModule } from "./module/types.ts";
import { SimContext } from "./ui/context.ts";
import { Simmy } from "./ui/Simmy.tsx";
import type { Platform } from "./ui/Ribbon.tsx";
import css from "./ui/styles.css?inline";

const FONTS_HREF = "https://fonts.googleapis.com/css2?family=Carlito:wght@400;700&family=Inter:wght@400;500;600;700&display=swap";

let sheet: CSSStyleSheet | null = null;

function styleSheet(): CSSStyleSheet {
  if (!sheet) {
    sheet = new CSSStyleSheet();
    sheet.replaceSync(css);
  }
  return sheet;
}

/** @font-face rules don't work inside a shadow root, so the fonts load once into the page. */
function loadFonts(): void {
  if (document.querySelector(`link[href="${FONTS_HREF}"], link[data-simmy-fonts]`)) {
    return;
  }
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = FONTS_HREF;
  link.dataset.simmyFonts = "";
  document.head.append(link);
}

/** Strip the indentation every line shares, so YAML can sit indented inside a <script> tag. */
function dedent(text: string): string {
  const lines = text.replace(/^\s*\n/, "").split("\n");
  const indent = Math.min(
    ...lines.filter((line) => line.trim() !== "").map((line) => /^[ \t]*/.exec(line)?.[0].length ?? 0),
  );
  return Number.isFinite(indent) ? lines.map((line) => line.slice(indent)).join("\n") : text;
}

const PLATFORMS: readonly Platform[] = ["mac", "win", "web"];

// Importing the package on a server (to reach parseModule, say) must not trip over the missing DOM.
const Base = (typeof HTMLElement === "undefined" ? class {} : HTMLElement) as typeof HTMLElement;

export class SimmyElement extends Base {
  static observedAttributes = ["src", "platform", "mode", "ribbon", "autoplay", "history", "fit-screen", "chrome-offset"];

  readonly #root: ShadowRoot;
  #module: SimModule | null = null;
  #explicit = false;
  #error: Error | null = null;
  #load = 0;
  #connected = false;
  #chromeObserver: ResizeObserver | null = null;

  constructor() {
    super();
    this.#root = this.attachShadow({ mode: "open" });
    this.#root.adoptedStyleSheets = [styleSheet()];
  }

  /** The resolved module, once loaded. Set it to an object to configure the element from script. */
  get module(): SimModule | null {
    return this.#module;
  }

  set module(value: unknown) {
    this.#explicit = value !== null && value !== undefined;
    if (!this.#explicit) {
      void this.#reload();
      return;
    }
    try {
      this.#setModule(parseModule(value));
    } catch (error) {
      this.#fail(error as Error);
    }
  }

  connectedCallback(): void {
    this.#connected = true;
    if (this.getAttribute("fonts") !== "none") {
      loadFonts();
    }
    this.#watchChrome();
    if (!this.#explicit && !this.#module) {
      void this.#reload();
    } else {
      this.#render();
    }
  }

  disconnectedCallback(): void {
    this.#connected = false;
    this.#chromeObserver?.disconnect();
    this.#chromeObserver = null;
    render(null, this.#root);
  }

  attributeChangedCallback(name: string, old: string | null, value: string | null): void {
    if (!this.#connected || old === value) {
      return;
    }
    if (name === "src" && !this.#explicit) {
      void this.#reload();
    } else if (name === "chrome-offset") {
      this.#watchChrome();
    } else {
      this.#render();
    }
  }

  /** Inline YAML wins over src; a module set from script wins over both. */
  async #reload(): Promise<void> {
    const load = ++this.#load;
    const inline = this.querySelector<HTMLScriptElement>(
      'script[type="text/yaml"], script[type="application/yaml"], script[type="application/json"]',
    );
    const src = this.getAttribute("src");
    if (!inline && !src) {
      return;
    }
    try {
      // A module loaded by URL can name clips as files in its audio folder.
      // Inline YAML has no folder, so those names stay as written.
      let text: string;
      let moduleUrl: URL | null = null;
      if (inline) {
        text = dedent(inline.textContent ?? "");
      } else {
        moduleUrl = new URL(src as string, document.baseURI);
        const response = await fetch(moduleUrl);
        if (!response.ok) {
          throw new Error(`Couldn't load ${src} (${response.status} ${response.statusText})`);
        }
        text = await response.text();
      }
      const parsed = await loadModule(text);
      const mod = moduleUrl ? resolveModuleAudio(parsed, moduleUrl) : parsed;
      if (load === this.#load) {
        this.#setModule(mod);
      }
    } catch (error) {
      if (load === this.#load) {
        this.#fail(error as Error);
      }
    }
  }

  #setModule(mod: SimModule): void {
    this.#module = mod;
    this.#error = null;
    this.#render();
    this.dispatchEvent(new CustomEvent("simmy:ready", { detail: mod, bubbles: true, composed: true }));
  }

  #fail(error: Error): void {
    this.#module = null;
    this.#error = error;
    this.#render();
    this.dispatchEvent(new CustomEvent("simmy:error", { detail: error, bubbles: true, composed: true }));
  }

  /** Keep --chrome-h equal to the host's sticky header, so fit-screen sizes the desktop under it. */
  #watchChrome(): void {
    this.#chromeObserver?.disconnect();
    this.#chromeObserver = null;
    const offset = this.getAttribute("chrome-offset");
    if (!offset) {
      this.style.removeProperty("--chrome-h");
      return;
    }
    // A length ("72px", "4.5rem", "72"); anything else is a selector, which may well start with a dot.
    if (/^\.?\d/.test(offset)) {
      this.style.setProperty("--chrome-h", /^[\d.]+$/.test(offset) ? `${offset}px` : offset);
      return;
    }
    const chrome = document.querySelector<HTMLElement>(offset);
    if (!chrome) {
      return;
    }
    const sync = () => this.style.setProperty("--chrome-h", `${chrome.offsetHeight}px`);
    sync();
    this.#chromeObserver = new ResizeObserver(sync);
    this.#chromeObserver.observe(chrome);
  }

  #render(): void {
    if (!this.#connected) {
      return;
    }
    if (this.#error) {
      const problems = this.#error instanceof ModuleError ? this.#error.problems : [this.#error.message];
      render(
        h(
          "div",
          { class: "xd-error", role: "alert" },
          h("strong", null, "This Excel simulator couldn't load its module."),
          h("ul", null, ...problems.map((problem) => h("li", null, problem))),
        ),
        this.#root,
      );
      return;
    }
    const mod = this.#module;
    if (!mod) {
      render(null, this.#root);
      return;
    }
    const platform = this.getAttribute("platform");
    const onGraded = (report: GradeReport) =>
      this.dispatchEvent(new CustomEvent("simmy:graded", { detail: report, bubbles: true, composed: true }));
    render(
      h(
        SimContext.Provider,
        { value: { module: mod, onGraded } },
        // A new module is a new workbook: key on it so every bit of sheet state starts over.
        h(Simmy, {
          key: mod,
          platform: PLATFORMS.includes(platform as Platform) ? (platform as Platform) : "auto",
          mode: this.getAttribute("mode") === "kiosk" ? "kiosk" : "movie",
          ribbon: this.getAttribute("ribbon") === "collapsed" ? "collapsed" : "open",
          autoplay: this.getAttribute("autoplay") !== "false",
          history: this.hasAttribute("history"),
          fitScreen: this.hasAttribute("fit-screen"),
          head: this.querySelector(':scope > [slot="head"]') ? h("slot", { name: "head" }) : undefined,
        }),
      ),
      this.#root,
    );
  }
}

/** Register <simmy-app> (or another tag name). Safe to call more than once. */
export function defineSimmy(tag = "simmy-app"): void {
  if (typeof customElements === "undefined" || customElements.get(tag)) {
    return;
  }
  customElements.define(tag, class extends SimmyElement {});
}

declare global {
  interface HTMLElementTagNameMap {
    "simmy-app": SimmyElement;
  }
}
