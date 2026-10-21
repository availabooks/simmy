/**
 * The simulator's auto mode: a virtual cursor that works the assignment the way a
 * student would. It moves a cursor over the real UI, then drives the sheet through
 * `WorkbookApi` and the task pane through ordinary clicks.
 */
import { addrName, parseAddr } from "../engine/sheet.ts";
import type { PilotStep } from "../module/types.ts";

/** What the autopilot may do to the sheet. Positions are client (viewport) pixels. */
export type WorkbookApi = {
  /** Scroll the grid so a cell is on screen. */
  reveal(ref: string): void;
  cellPoint(ref: string): { x: number; y: number } | null;
  fillHandlePoint(): { x: number; y: number } | null;
  select(ref: string): void;
  /** Show `text` in the active cell's editor, as if it had just been typed. */
  setTyping(text: string): void;
  commit(): void;
  /** Show the dashed fill outline from `source` to `target`, or hide it with null. */
  previewFill(source: string, target: string | null): void;
  applyFill(source: string, target: string): void;
  cancelAll(): void;
};

export type PilotContext = {
  /** The steps to play, from the module. */
  plan: PilotStep[];
  desktop: HTMLElement;
  cursor: HTMLElement;
  api: WorkbookApi;
  cancelled: () => boolean;
  /** Resolves once the demo is on screen, so it doesn't run while nobody can see it. */
  visible: () => Promise<void>;
};

class Cancelled extends Error {}

type Point = { x: number; y: number };

const jitter = (ms: number) => ms * (0.85 + Math.random() * 0.3);

/** Runs one pass of the demo. Resolves when it finishes or is cancelled. */
export async function runAutopilot(ctx: PilotContext): Promise<void> {
  const { desktop, cursor, api } = ctx;

  const sleep = async (ms: number) => {
    await new Promise((resolve) => window.setTimeout(resolve, ms));
    await ctx.visible();
    if (ctx.cancelled()) {
      throw new Cancelled();
    }
  };
  const query = <T extends Element>(selector: string, root: ParentNode = desktop) =>
    root.querySelector<T>(selector);
  const need = <T extends Element>(selector: string, root: ParentNode = desktop): T => {
    const el = query<T>(selector, root);
    if (!el) {
      throw new Cancelled();
    }
    return el;
  };

  /* ---------------------------------------------------------- the cursor */

  let at: Point = { x: desktop.clientWidth * 0.55, y: desktop.clientHeight * 0.55 };
  const place = (p: Point, ms: number) => {
    cursor.style.transitionDuration = `${ms}ms`;
    cursor.style.transform = `translate(${p.x}px, ${p.y}px)`;
    at = p;
  };
  const local = (client: Point): Point => {
    const box = desktop.getBoundingClientRect();
    return { x: client.x - box.left, y: client.y - box.top };
  };
  const centerOf = (el: Element): Point => {
    const box = el.getBoundingClientRect();
    return { x: box.left + box.width / 2, y: box.top + box.height / 2 };
  };

  const moveTo = async (client: Point, quick = false) => {
    const to = local(client);
    const distance = Math.hypot(to.x - at.x, to.y - at.y);
    const ms = quick ? 140 : Math.min(1100, Math.max(320, 260 + distance * 1.05));
    place(to, ms);
    await sleep(ms + (quick ? 0 : 60));
  };
  const press = async () => {
    cursor.classList.add("is-down");
    await sleep(110);
    cursor.classList.remove("is-down");
  };
  const clickEl = async (el: HTMLElement) => {
    await moveTo(centerOf(el));
    await press();
    el.click();
  };

  /* ------------------------------------------------------------ the pane */

  const scrollPaneTo = async (el: Element, where: "top" | "middle") => {
    const box = query<HTMLElement>(".ee-scroll");
    if (!box) {
      return;
    }
    const inner = el.getBoundingClientRect().top - box.getBoundingClientRect().top;
    const offset = where === "top" ? 6 : box.clientHeight / 2 - el.getBoundingClientRect().height / 2;
    box.scrollTo({ top: box.scrollTop + inner - offset, behavior: "smooth" });
    await sleep(520);
  };

  const taskEl = (id: string) => need<HTMLElement>(`.ee-task[data-task="${id}"]`);

  const openTask = async (id: string) => {
    const task = taskEl(id);
    await scrollPaneTo(task, "top");
    if (task.classList.contains("is-collapsed")) {
      await clickEl(need<HTMLElement>(".ee-task-text", task));
      await sleep(jitter(500));
      await scrollPaneTo(taskEl(id), "top");
    }
    await sleep(jitter(1100));
  };

  const finishTask = async (id: string) => {
    const task = taskEl(id);
    await scrollPaneTo(task, "top");
    await clickEl(need<HTMLElement>(".ee-pill", task));
    await sleep(jitter(700));
  };

  /* ------------------------------------------------------------ the sheet */

  const pointAtCell = async (ref: string, quick = false) => {
    api.reveal(ref);
    await sleep(40);
    const p = api.cellPoint(ref);
    if (p) {
      await moveTo({ x: p.x + (quick ? 0 : (Math.random() - 0.5) * 10), y: p.y }, quick);
    }
  };

  const typeFormula = async (text: string) => {
    for (let i = 1; i <= text.length; i++) {
      api.setTyping(text.slice(0, i));
      const last = text[i - 1];
      await sleep(jitter(last === "(" || last === "," || last === ":" ? 150 : 82));
    }
  };

  /** Drag the fill handle from the cell out to the end of `fill`. */
  const dragFill = async (source: string, fill: string) => {
    const [, end = source] = fill.split(":");
    const from = parseAddr(source);
    const to = parseAddr(end);
    if (!from || !to) {
      return;
    }
    api.select(source);
    api.reveal(end);
    await sleep(120);
    const handle = api.fillHandlePoint();
    if (!handle) {
      return;
    }
    await moveTo(handle);
    cursor.classList.add("is-down");
    await sleep(140);
    const across = to.col > from.col;
    const steps = across ? to.col - from.col : to.row - from.row;
    for (let i = 1; i <= steps; i++) {
      const stop = addrName(across ? { col: from.col + i, row: from.row } : { col: from.col, row: from.row + i });
      const p = api.cellPoint(stop);
      if (p) {
        await moveTo(p, true);
      }
      api.previewFill(source, `${source}:${stop}`);
      await sleep(jitter(190));
    }
    await sleep(260);
    cursor.classList.remove("is-down");
    api.applyFill(source, fill);
    api.previewFill(source, null);
    await sleep(jitter(700));
  };

  const doByHand = async (step: Extract<PilotStep, { kind: "type" }>) => {
    await pointAtCell(step.cell);
    await press();
    api.select(step.cell);
    await sleep(jitter(450));
    await typeFormula(step.formula);
    await sleep(jitter(650));
    api.commit();
    await sleep(jitter(800));
    if (step.fill) {
      await dragFill(step.cell, step.fill);
    }
  };

  const doItButton = (id: string) => need<HTMLElement>('[data-pilot="doit"]', taskEl(id));

  const wand = async (id: string) => {
    const task = taskEl(id);
    await scrollPaneTo(need<HTMLElement>('[data-pilot="doit"]', task), "middle");
    await clickEl(doItButton(id));
    await sleep(jitter(1300));
  };

  /* ------------------------------------------------------------- the show */

  try {
    place(at, 0);
    cursor.classList.add("is-on");
    await sleep(900);

    for (const step of ctx.plan) {
      await openTask(step.taskId);
      if (step.kind === "type") {
        await doByHand(step);
      } else {
        await wand(step.taskId);
        if (step.showUndo) {
          await sleep(jitter(700));
          await wand(step.taskId); // Undo it
          await sleep(jitter(500));
          await wand(step.taskId); // and Do it again
        }
      }
      await finishTask(step.taskId);
    }

    await clickEl(need<HTMLElement>(".ee-nav .ee-btn-primary"));
    await sleep(jitter(1100));
    await clickEl(need<HTMLElement>(".ee-submit-body .ee-btn-primary"));

    for (let waited = 0; !query(".ee-results"); waited += 150) {
      if (waited > 20000) {
        return;
      }
      await sleep(150);
    }
    await sleep(1200);
    const scroller = query<HTMLElement>(".ee-scroll");
    for (const stop of [0.35, 0.7, 1]) {
      if (scroller) {
        scroller.scrollTo({
          top: (scroller.scrollHeight - scroller.clientHeight) * stop,
          behavior: "smooth",
        });
      }
      await sleep(1800);
    }
    await sleep(5000);
  } catch (error) {
    if (!(error instanceof Cancelled)) {
      throw error;
    }
  } finally {
    cursor.classList.remove("is-down");
  }
}
