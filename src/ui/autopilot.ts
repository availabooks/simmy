/**
 * The simulator's auto mode: a virtual cursor that works the assignment the way a
 * student would. It moves a cursor over the real UI, then drives the sheet through
 * `WorkbookApi` and the task pane through ordinary clicks.
 */
import { addrName, parseAddr } from "../engine/sheet.ts";
import type { PilotStep } from "../module/types.ts";
import { Playback } from "./playback.ts";

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

/** Plays one narration clip per autopilot entry. The next clip is loaded just before it starts. */
export type PilotClock = {
  /** Load and play this entry's clip. Resolves with its length in seconds. */
  play(index: number): Promise<number>;
  /** Resolves when the clip that is playing ends. Returns early when the pass is cancelled. */
  untilEnded(): Promise<void>;
  /** Which transcript line the cursor is on. */
  onLine(index: number): void;
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
  /** Set when the module has a recording. Each entry then waits for its clock window. */
  clock?: PilotClock;
  /**
   * Where playback should be. A function is read continuously, so dragging the
   * time bar can move the cursor and the sheet while a pass is already running.
   * `rush` finishes the target step's actions quickly so they catch the new time.
   */
  skipTo?:
    | { index: number; offset: number; rush?: boolean }
    | (() => { index: number; offset: number; rush?: boolean });
  /** Pauses sleeps while the visitor has paused or taken control. Omitting it never waits. */
  playback?: Playback;
  /** Fires as each step starts, including ones that are skipped over. */
  onStep?: (index: number) => void;
};

class Cancelled extends Error {}

type Point = { x: number; y: number };

const jitter = (ms: number) => ms * (0.85 + Math.random() * 0.3);

/** Pauses the timed demo already uses. A timed entry multiplies them to fit its window. */
const SCROLL = 520;
const PRESS = 110;
const OPEN_AFTER_CLICK = 500;
const OPEN_PAUSE = 1100;
const FINISH_PAUSE = 700;
const POINT_PAUSE = 40;
const SELECT_PAUSE = 450;
const AFTER_TYPE = 650;
const AFTER_COMMIT = 800;
const FILL_REVEAL = 120;
const FILL_HOLD = 140;
const FILL_STEP = 190;
const FILL_SETTLE = 260;
const FILL_AFTER = 700;
const WAND_PAUSE = 1300;
const UNDO_GAP = 700;
const REDO_GAP = 500;
/** A cursor move, before we know the real distance. Long enough that a short window finishes early and waits. */
const MOVE_ESTIMATE = 710;
const MOVE_SETTLE = 60;
const QUICK_MOVE = 140;
const CHAR_PAUSE = 82;
const CHAR_PAUSE_SLOW = 150;
const MIN_SCALE = 0.45;
const MAX_SCALE = 2.5;

const charPause = (character: string) =>
  character === "(" || character === "," || character === ":" ? CHAR_PAUSE_SLOW : CHAR_PAUSE;

/** How many cells a fill drag steps through. */
function fillCount(source: string, fill: string): number {
  const [, end = source] = fill.split(":");
  const from = parseAddr(source);
  const to = parseAddr(end);
  if (!from || !to) {
    return 0;
  }
  return Math.max(0, to.col > from.col ? to.col - from.col : to.row - from.row);
}

const clickEstimate = MOVE_ESTIMATE + MOVE_SETTLE + PRESS;

/** How long an entry takes at the demo's normal speed, with each random pause at its midpoint. */
const fillDragMs = (source: string, fill: string) => {
  const dragged = fillCount(source, fill);
  return (
    FILL_REVEAL +
    MOVE_ESTIMATE +
    MOVE_SETTLE +
    FILL_HOLD +
    dragged * (QUICK_MOVE + FILL_STEP) +
    FILL_SETTLE +
    FILL_AFTER
  );
};

function nominalMs(step: PilotStep): number {
  const open = SCROLL + clickEstimate + OPEN_AFTER_CLICK + SCROLL + OPEN_PAUSE;
  const finish = SCROLL + clickEstimate + FINISH_PAUSE;
  const wand = SCROLL + clickEstimate + WAND_PAUSE;
  const done = (step.kind === "type" || step.kind === "fill" || step.kind === "doit") && step.finish !== false ? finish : 0;
  switch (step.kind) {
    case "hold":
      return step.cell ? POINT_PAUSE + MOVE_ESTIMATE + MOVE_SETTLE : 1;
    case "open":
      return open;
    case "submit":
      return clickEstimate + 1100 + clickEstimate;
    case "fill":
      return fillDragMs(step.cell, step.fill) + done;
    case "doit":
      return open + wand * (step.showUndo ? 3 : 1) + (step.showUndo ? UNDO_GAP + REDO_GAP : 0) + done;
    case "type": {
      const typing = [...step.formula].reduce((sum, character) => sum + charPause(character), 0);
      const fill = step.fill ? fillDragMs(step.cell, step.fill) : 0;
      return (
        open + POINT_PAUSE + MOVE_ESTIMATE + MOVE_SETTLE + PRESS + SELECT_PAUSE + typing + AFTER_TYPE + AFTER_COMMIT + fill + done
      );
    }
    default: {
      const unreachable: never = step;
      return unreachable;
    }
  }
}

/** Runs one pass of the demo. Resolves when it finishes or is cancelled. */
export async function runAutopilot(ctx: PilotContext): Promise<void> {
  const { desktop, cursor, api } = ctx;
  const playback = ctx.playback ?? new Playback();
  let scale = 1;
  let pacing = false;
  let turbo = false;
  let stepIndex = 0;
  const readSkip = () => {
    const value = typeof ctx.skipTo === "function" ? ctx.skipTo() : ctx.skipTo;
    return value ?? { index: 0, offset: 0 };
  };
  const showCursor = () => {
    if (playback.hold !== "control") {
      cursor.classList.add("is-on");
    }
  };

  const sleep = async (ms: number, steady = false) => {
    if (ctx.cancelled()) {
      throw new Cancelled();
    }
    const catchingUp = () => {
      const skip = readSkip();
      return stepIndex < skip.index || (skip.rush === true && stepIndex === skip.index);
    };
    if (turbo || catchingUp()) {
      showCursor();
      await new Promise((resolve) => window.setTimeout(resolve, 16));
      if (ctx.cancelled()) {
        throw new Cancelled();
      }
      return;
    }
    let budget = pacing ? ms * scale : steady ? ms : jitter(ms);
    while (budget > 0) {
      if (catchingUp()) {
        showCursor();
        return;
      }
      await playback.gate(ctx.cancelled);
      if (ctx.cancelled()) {
        throw new Cancelled();
      }
      const rate = playback.rate > 0 ? playback.rate : 1;
      const slice = Math.min(budget / rate, 80);
      if (!(slice > 0)) {
        break;
      }
      const started = performance.now();
      await new Promise((resolve) => window.setTimeout(resolve, slice));
      budget -= (performance.now() - started) * rate;
    }
    await ctx.visible();
    await playback.gate(ctx.cancelled);
    if (ctx.cancelled()) {
      throw new Cancelled();
    }
  };
  const ended = async () => {
    if (!ctx.clock) {
      return;
    }
    // The clip's promise stays pending until the audio ends. A skip to a later
    // step has to leave without waiting that out, or the buttons and the bar do nothing.
    let finished = false;
    const pending = ctx.clock.untilEnded().then(() => {
      finished = true;
    });
    while (!finished) {
      if (ctx.cancelled()) {
        throw new Cancelled();
      }
      if (stepIndex < readSkip().index) {
        return;
      }
      await new Promise((resolve) => window.setTimeout(resolve, 16));
    }
    await pending;
    await ctx.visible();
    await playback.gate(ctx.cancelled);
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
    const ms = quick ? QUICK_MOVE : Math.min(1100, Math.max(320, 260 + distance * 1.05));
    const extra = quick ? 0 : MOVE_SETTLE;
    place(to, pacing ? ms * scale : ms);
    await sleep(ms + extra, true);
  };
  const press = async () => {
    cursor.classList.add("is-down");
    await sleep(PRESS, true);
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
    await sleep(SCROLL, true);
  };

  const taskEl = (id: string) => need<HTMLElement>(`.ee-task[data-task="${id}"]`);

  const openTask = async (id: string) => {
    const task = taskEl(id);
    await scrollPaneTo(task, "top");
    if (task.classList.contains("is-collapsed")) {
      await clickEl(need<HTMLElement>(".ee-task-text", task));
      await sleep(OPEN_AFTER_CLICK);
      await scrollPaneTo(taskEl(id), "top");
    }
    await sleep(OPEN_PAUSE);
  };

  const finishTask = async (id: string) => {
    const task = taskEl(id);
    await scrollPaneTo(task, "top");
    await clickEl(need<HTMLElement>(".ee-pill", task));
    await sleep(FINISH_PAUSE);
  };

  /* ------------------------------------------------------------ the sheet */

  const pointAtCell = async (ref: string, quick = false) => {
    api.reveal(ref);
    await sleep(POINT_PAUSE, true);
    const p = api.cellPoint(ref);
    if (p) {
      await moveTo({ x: p.x + (quick ? 0 : (Math.random() - 0.5) * 10), y: p.y }, quick);
    }
  };

  const typeFormula = async (text: string) => {
    for (let i = 1; i <= text.length; i++) {
      api.setTyping(text.slice(0, i));
      await sleep(charPause(text[i - 1] ?? ""));
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
    await sleep(FILL_REVEAL, true);
    const handle = api.fillHandlePoint();
    if (!handle) {
      return;
    }
    await moveTo(handle);
    cursor.classList.add("is-down");
    await sleep(FILL_HOLD, true);
    const across = to.col > from.col;
    const steps = across ? to.col - from.col : to.row - from.row;
    for (let i = 1; i <= steps; i++) {
      const stop = addrName(across ? { col: from.col + i, row: from.row } : { col: from.col, row: from.row + i });
      const p = api.cellPoint(stop);
      if (p) {
        await moveTo(p, true);
      }
      api.previewFill(source, `${source}:${stop}`);
      await sleep(FILL_STEP);
    }
    await sleep(FILL_SETTLE, true);
    cursor.classList.remove("is-down");
    api.applyFill(source, fill);
    api.previewFill(source, null);
    await sleep(FILL_AFTER);
  };

  const doByHand = async (step: Extract<PilotStep, { kind: "type" }>) => {
    await pointAtCell(step.cell);
    await press();
    api.select(step.cell);
    await sleep(SELECT_PAUSE);
    await typeFormula(step.formula);
    await sleep(AFTER_TYPE);
    api.commit();
    await sleep(AFTER_COMMIT);
    if (step.fill) {
      await dragFill(step.cell, step.fill);
    }
  };

  const doItButton = (id: string) => need<HTMLElement>('[data-pilot="doit"]', taskEl(id));

  const wand = async (id: string) => {
    const task = taskEl(id);
    await scrollPaneTo(need<HTMLElement>('[data-pilot="doit"]', task), "middle");
    await clickEl(doItButton(id));
    await sleep(WAND_PAUSE);
  };

  /* ------------------------------------------------------------- the show */

  let submitted = false;
  const playStep = async (step: PilotStep) => {
    switch (step.kind) {
      case "hold":
        if (step.cell) {
          await pointAtCell(step.cell);
        }
        return;
      case "open":
        await openTask(step.taskId);
        return;
      case "type":
        await openTask(step.taskId);
        await doByHand(step);
        if (step.finish !== false) {
          await finishTask(step.taskId);
        }
        return;
      case "fill":
        await dragFill(step.cell, step.fill);
        if (step.finish !== false) {
          await finishTask(step.taskId);
        }
        return;
      case "doit":
        await openTask(step.taskId);
        await wand(step.taskId);
        if (step.showUndo) {
          await sleep(UNDO_GAP);
          await wand(step.taskId);
          await sleep(REDO_GAP);
          await wand(step.taskId);
        }
        if (step.finish !== false) {
          await finishTask(step.taskId);
        }
        return;
      case "submit":
        await clickEl(need<HTMLElement>(".ee-nav .ee-btn-primary"));
        await sleep(1100);
        await clickEl(need<HTMLElement>(".ee-submit-body .ee-btn-primary"));
        submitted = true;
        return;
      default: {
        const unreachable: never = step;
        return unreachable;
      }
    }
  };

  try {
    place(at, 0);
    showCursor();
    if (!ctx.clock && readSkip().index === 0) {
      await sleep(900, true);
    }

    for (let index = 0; index < ctx.plan.length; index++) {
      const step = ctx.plan[index];
      if (!step) {
        continue;
      }
      stepIndex = index;
      const skip = readSkip();
      ctx.onStep?.(index);
      if (index < skip.index) {
        turbo = true;
        showCursor();
        ctx.clock?.onLine(index);
        await playStep(step);
        turbo = false;
        continue;
      }
      await playback.gate(ctx.cancelled);
      if (ctx.cancelled()) {
        throw new Cancelled();
      }
      showCursor();
      const clock = ctx.clock;
      if (clock && step.audio !== undefined) {
        const duration = await clock.play(index);
        if (ctx.cancelled()) {
          throw new Cancelled();
        }
        clock.onLine(index);
        const live = readSkip();
        const offset = index === live.index ? Math.min(live.offset, duration) : 0;
        const windowMs = Math.max(0, duration - offset) * 1000;
        if (!live.rush && step.kind !== "hold" && windowMs > 0) {
          scale = Math.min(MAX_SCALE, Math.max(MIN_SCALE, windowMs / nominalMs(step)));
          pacing = true;
        }
        turbo = live.rush === true && index === live.index;
        await playStep(step);
        turbo = false;
        pacing = false;
        scale = 1;
        await ended();
        continue;
      }
      await playStep(step);
    }

    if (!submitted) {
      await clickEl(need<HTMLElement>(".ee-nav .ee-btn-primary"));
      await sleep(1100);
      await clickEl(need<HTMLElement>(".ee-submit-body .ee-btn-primary"));
    }

    for (let waited = 0; !query(".ee-results"); waited += 150) {
      if (waited > 20000) {
        return;
      }
      await sleep(150, true);
    }
    await sleep(1200, true);
    const scroller = query<HTMLElement>(".ee-scroll");
    for (const stop of [0.35, 0.7, 1]) {
      if (scroller) {
        scroller.scrollTo({
          top: (scroller.scrollHeight - scroller.clientHeight) * stop,
          behavior: "smooth",
        });
      }
      await sleep(1800, true);
    }
    await sleep(5000, true);
  } catch (error) {
    if (!(error instanceof Cancelled)) {
      throw error;
    }
  } finally {
    cursor.classList.remove("is-down");
  }
}
