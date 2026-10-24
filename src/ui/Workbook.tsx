import type { TargetedEvent, TargetedKeyboardEvent, TargetedPointerEvent } from "preact";
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "preact/hooks";
import {
  type Addr,
  addrName,
  checkFormula,
  cycleReference,
  type Clip,
  colName,
  copyRange,
  evaluateBook,
  fillWrites,
  formatValue,
  formulaRanges,
  FUNCTIONS,
  type Inputs,
  inRange,
  isError,
  normalizeFormula,
  normalizeRange,
  parseAddr,
  parseRange,
  pasteWrites,
  type Range,
  rangeAddrs,
  rangeName,
  rangeText,
  suggestFunctions,
  textPasteWrites,
  type Value,
} from "../engine/sheet.ts";
import {
  absoluteRef,
  guessLabelSides,
  lookupName,
  nameForRange,
  nameProblem,
  nameRange,
  namesFromSelection,
  refersToText,
  renameWrites,
  suggestName,
} from "../engine/names.ts";
import { applyChange, type Change, nextSeed, type WorkbookState } from "../engine/workbook.ts";
import { evalOptions, initialState } from "../module/grade.ts";
import { Menu } from "./Menu.tsx";
import {
  CreateNamesDialog,
  DefineNameDialog,
  GoToDialog,
  NameManagerDialog,
  PasteNameDialog,
} from "./NameDialogs.tsx";
import type { CellStyle, SimModule } from "../module/types.ts";
import type { WorkbookApi } from "./autopilot.ts";
import { useSim } from "./context.ts";
import { type Platform, Ribbon, type RibbonTab, RibbonTabs, Svg } from "./Ribbon.tsx";
import { TaskPane, type TaskPaneProps } from "./TaskPane.tsx";

const ROW_H = 20;
const HEAD_H = 20;
const ROWHEAD_W = 34;
const REF_COLORS = ["#2f6fdb", "#d6322b", "#8a3ab9", "#1e8a4c", "#c55a11", "#0f8b8d"];
const POINTABLE = /[=(,+\-*/^&<>:]\s*$/;

/** Each function's arguments, for the tooltip shown while typing inside its parentheses. */
const SIGNATURES: Record<string, string> = Object.fromEntries(FUNCTIONS.map((fn) => [fn.name, fn.args]));

type Selection = { anchor: Addr; focus: Addr };
/** An entry in the formula autocomplete list: a function (inserted with its "(") or a defined name. */
type Suggestion = { name: string; summary: string; call: boolean };
/** An open dialog. Dialogs stack: Name Manager can open New Name over itself. */
type DialogState =
  | { kind: "goto" }
  | { kind: "define"; editing?: string }
  | { kind: "create" }
  | { kind: "manager" }
  | { kind: "paste" };
/** Excel's message box. `formula` adds its note about formulas; `from` is where focus goes back. */
type Alert = { message: string; formula?: boolean; from: "editor" | "namebox" };
type Point = { start: number; end: number; anchor: Addr; focus: Addr };
type Edit = { addr: Addr; text: string; mode: "enter" | "edit"; origin: "cell" | "bar"; point: Point | null };
/** The workbook, with the states Undo and Redo go back and forth through. */
type Book = { state: WorkbookState; past: WorkbookState[]; future: WorkbookState[] };
type Drag = { kind: "select" } | { kind: "point" } | { kind: "fill"; source: Range } | null;

/** The module's sheet in the shapes the grid draws from. */
function sheetLayout(mod: SimModule) {
  const widths = mod.sheet.widths;
  const colLeft = widths.reduce<number[]>((acc, w, i) => [...acc, (acc[i] ?? 0) + w], [0]);
  return {
    widths,
    cols: widths.length,
    colLeft,
    sheetW: colLeft[widths.length],
    rowCount: mod.sheet.rows,
    styles: mod.sheet.styles,
    outlines: mod.sheet.outlines,
    startAddr: parseAddr(mod.sheet.start) ?? { col: 0, row: 0 },
  };
}

function formatName(style: CellStyle | undefined): string {
  switch (style?.format) {
    case "0":
    case "0.00":
      return "Number";
    case "$#,##0.00":
      return "Currency";
    case "accounting":
      return "Accounting";
    case "0.00%":
      return "Percentage";
    default:
      return "General";
  }
}

function statusNumber(n: number): string {
  return formatValue(n, "General", 120);
}

/** The function call the caret is inside, for the argument tooltip. */
function activeCall(text: string, caret: number): string | null {
  const stack: (string | null)[] = [];
  const before = text.slice(0, caret);
  const re = /([A-Za-z][A-Za-z0-9.]*)?\(|\)/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(before))) {
    if (match[0] === ")") {
      stack.pop();
    } else {
      stack.push(match[1] ? match[1].toUpperCase() : null);
    }
  }
  for (let i = stack.length - 1; i >= 0; i--) {
    const name = stack[i];
    if (name && SIGNATURES[name]) {
      return name;
    }
  }
  return null;
}

function typedFunctionPrefix(text: string, caret: number): string | null {
  if (!text.startsWith("=")) {
    return null;
  }
  const match = /(^=|[=(,+\-*/^&<>:\s])([A-Za-z_\\][A-Za-z0-9._]*)$/.exec(text.slice(0, caret));
  return match ? match[2] : null;
}

export type WorkbookProps = {
  platform: Platform;
  resetKey: number;
  /** Start with the ribbon collapsed to its tabs. The visitor can still expand it. */
  ribbonCollapsed?: boolean;
  maximized: boolean;
  onTitlePointerDown: (event: TargetedPointerEvent<HTMLElement>) => void;
  onMinimize: () => void;
  onMaximize: () => void;
  onClose: () => void;
  onUnsupported: TaskPaneProps["onUnsupported"];
  /** Hands the autopilot its controls once the sheet is mounted. */
  onApi: (api: WorkbookApi | null) => void;
};

export function Workbook(props: WorkbookProps) {
  const { platform } = props;
  const { module: mod } = useSim();
  const { widths, cols, colLeft, sheetW, rowCount, styles, outlines, startAddr } = useMemo(
    () => sheetLayout(mod),
    [mod],
  );
  const clampAddr = ({ col, row }: Addr): Addr => ({
    col: Math.max(0, Math.min(cols - 1, col)),
    row: Math.max(0, Math.min(rowCount - 1, row)),
  });
  const [book, setBook] = useState<Book>(() => ({ state: initialState(mod), past: [], future: [] }));
  const { inputs, names } = book.state;
  const [sel, setSel] = useState<Selection>({ anchor: startAddr, focus: startAddr });
  const [edit, setEdit] = useState<Edit | null>(null);
  const [clip, setClip] = useState<Clip | null>(null);
  /** What the last copy put on the system clipboard, to know a paste of it when it comes back. */
  const copiedText = useRef("");
  /** Set when a copy couldn't reach the system clipboard: then a paste uses the copied cells. */
  const internalOnly = useRef(false);
  const [fillTarget, setFillTarget] = useState<Range | null>(null);
  const [alert, setAlert] = useState<Alert | null>(null);
  const [dialogs, setDialogs] = useState<DialogState[]>([]);
  const openDialog = (dialog: DialogState) => setDialogs((open) => [...open, dialog]);
  const closeDialog = () => {
    setDialogs((open) => open.slice(0, -1));
    // Back to the formula being typed, if a dialog (Paste Name) opened over one.
    const editor = editRef.current?.origin === "bar" ? barRef.current : cellEditorRef.current;
    window.requestAnimationFrame(() => {
      (editRef.current ? editor : scrollRef.current)?.focus({ preventScroll: true });
    });
  };
  const [paneOpen, setPaneOpen] = useState(true);
  const [tab, setTab] = useState<RibbonTab>("Home");
  const [ribbonCollapsed, setRibbonCollapsed] = useState(props.ribbonCollapsed ?? false);
  // A collapsed ribbon shows over the sheet while the visitor uses one of its tabs, as in Excel.
  const [ribbonPeek, setRibbonPeek] = useState(false);
  const [suggest, setSuggest] = useState(0);
  const [caret, setCaret] = useState(0);
  const [paneKey, setPaneKey] = useState(0);
  const [selecting, setSelecting] = useState(false);

  const scrollRef = useRef<HTMLDivElement>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const cellEditorRef = useRef<HTMLInputElement>(null);
  const barRef = useRef<HTMLInputElement>(null);
  const alertOkRef = useRef<HTMLButtonElement>(null);
  const nameBoxRef = useRef<HTMLInputElement>(null);
  const drag = useRef<Drag>(null);
  const pointer = useRef<{ x: number; y: number } | null>(null);
  const editRef = useRef<Edit | null>(null);
  editRef.current = edit;

  const sheetEval = useMemo(
    () => evaluateBook(inputs, evalOptions(mod, book.state)),
    [book.state, mod],
  );
  const { values } = sheetEval;
  const selRange = normalizeRange(sel.anchor, sel.focus);
  const active = sel.anchor;

  useEffect(() => {
    setRibbonCollapsed(props.ribbonCollapsed ?? false);
    setRibbonPeek(false);
  }, [props.ribbonCollapsed]);

  useEffect(() => {
    if (props.resetKey === 0) {
      return;
    }
    setBook({ state: initialState(mod), past: [], future: [] });
    setSel({ anchor: startAddr, focus: startAddr });
    setEdit(null);
    setClip(null);
    setAlert(null);
    setPaneOpen(true);
    setPaneKey((k) => k + 1);
    scrollRef.current?.scrollTo({ left: 0, top: 0 });
  }, [props.resetKey]);

  /** Every change to the workbook comes through here, as one step for Undo. */
  const apply = useCallback((change: Change) => {
    setBook((b) => ({
      state: applyChange(b.state, change),
      past: [...b.past.slice(-99), b.state],
      future: [],
    }));
  }, []);
  const write = useCallback((writes: Inputs) => apply({ cells: writes }), [apply]);
  /** Calculate Now: RANDBETWEEN picks new numbers. As in Excel, it isn't a step to undo. */
  const recalculate = () => setBook((b) => ({ ...b, state: { ...b.state, seed: nextSeed(b.state.seed) } }));

  const undo = () =>
    setBook((b) =>
      b.past.length === 0
        ? b
        : { state: b.past[b.past.length - 1], past: b.past.slice(0, -1), future: [b.state, ...b.future] },
    );
  const redo = () =>
    setBook((b) =>
      b.future.length === 0 ? b : { state: b.future[0], past: [...b.past, b.state], future: b.future.slice(1) },
    );

  /* ---------------------------------------------------------- scrolling */

  const ensureVisible = useCallback((addr: Addr) => {
    const el = scrollRef.current;
    if (!el) {
      return;
    }
    const left = colLeft[addr.col];
    const right = colLeft[addr.col + 1];
    const top = addr.row * ROW_H;
    const viewW = el.clientWidth - ROWHEAD_W;
    const viewH = el.clientHeight - HEAD_H;
    if (left < el.scrollLeft) {
      el.scrollLeft = left;
    } else if (right > el.scrollLeft + viewW) {
      el.scrollLeft = Math.min(left, right - viewW);
    }
    if (top < el.scrollTop) {
      el.scrollTop = top;
    } else if (top + ROW_H > el.scrollTop + viewH) {
      el.scrollTop = top + ROW_H - viewH;
    }
  }, [colLeft]);

  useLayoutEffect(() => {
    // Open at the top left, where the whole order table sits.
    const el = scrollRef.current;
    if (el) {
      el.scrollLeft = 0;
      el.scrollTop = 0;
    }
  }, [props.resetKey]);

  const focusGrid = () => scrollRef.current?.focus({ preventScroll: true });

  /* ------------------------------------------------------------ editing */

  const startEdit = (text: string, mode: Edit["mode"], origin: Edit["origin"] = "cell") => {
    setEdit({ addr: active, text, mode, origin, point: null });
    setCaret(text.length);
    setSuggest(0);
    setClip(null);
  };

  // Focus the editor in the same task as the render, so fast typing doesn't land on the grid.
  useLayoutEffect(() => {
    if (!edit) {
      return;
    }
    const input = edit.origin === "cell" ? cellEditorRef.current : barRef.current;
    // Inside a shadow root, document.activeElement is the host, so ask the input's own root.
    if (input && (input.getRootNode() as Document | ShadowRoot).activeElement !== input) {
      input.focus({ preventScroll: true });
      const end = edit.text.length;
      input.setSelectionRange(end, end);
    }
  }, [edit?.origin, edit?.addr.col, edit?.addr.row, edit === null]);

  useLayoutEffect(() => {
    if (alert) {
      alertOkRef.current?.focus({ preventScroll: true });
    }
  }, [alert]);

  const dismissAlert = () => {
    const from = alert?.from;
    setAlert(null);
    const input =
      from === "namebox"
        ? nameBoxRef.current
        : editRef.current?.origin === "bar"
          ? barRef.current
          : cellEditorRef.current;
    window.requestAnimationFrame(() => {
      input?.focus({ preventScroll: true });
      if (from === "namebox") {
        input?.select();
      }
    });
  };

  /** Commit the edit. Returns false when Excel would refuse the formula. */
  const commitEdit = (): boolean => {
    const current = editRef.current;
    if (!current) {
      return true;
    }
    const text = normalizeFormula(current.text, names);
    if (text.startsWith("=")) {
      try {
        checkFormula(text);
      } catch {
        setAlert({ message: "There's a problem with this formula.", formula: true, from: "editor" });
        return false;
      }
    }
    const key = addrName(current.addr);
    if ((inputs[key] ?? "") !== text) {
      write({ [key]: text });
    }
    setEdit(null);
    return true;
  };

  const cancelEdit = () => {
    setEdit(null);
    focusGrid();
  };

  const moveActive = (dCol: number, dRow: number, extend = false) => {
    const base = extend ? sel.focus : sel.anchor;
    const next = clampAddr({ col: base.col + dCol, row: base.row + dRow });
    setSel(extend ? { anchor: sel.anchor, focus: next } : { anchor: next, focus: next });
    ensureVisible(next);
  };

  /** Commit the edit and move from the edited cell. Returns false when the formula was refused. */
  const commitAndMove = (dCol: number, dRow: number): boolean => {
    if (!commitEdit()) {
      return false;
    }
    const next = clampAddr({ col: active.col + dCol, row: active.row + dRow });
    setSel({ anchor: next, focus: next });
    ensureVisible(next);
    focusGrid();
    return true;
  };

  /**
   * Excel's Tab-then-Enter: after Tab has moved across a row, Enter goes to the next row in the
   * column the Tabs started from, so a table fills in row by row. `at` is where the last Tab
   * landed. Any other move leaves the selection somewhere else, which ends the run.
   */
  const tabRun = useRef<{ startCol: number; at: Addr } | null>(null);

  /** How far Enter or Tab moves from the active cell, and the Tab run to keep once it has. */
  const enterTabStep = (key: "Enter" | "Tab", shift: boolean) => {
    const run = tabRun.current;
    const inRun = run !== null && run.at.col === active.col && run.at.row === active.row;
    if (key === "Tab") {
      const at = clampAddr({ col: active.col + (shift ? -1 : 1), row: active.row });
      return {
        dCol: shift ? -1 : 1,
        dRow: 0,
        run: shift ? null : { startCol: inRun ? run.startCol : active.col, at },
      };
    }
    if (shift) {
      return { dCol: 0, dRow: -1, run: null };
    }
    return { dCol: inRun ? run.startCol - active.col : 0, dRow: 1, run: null };
  };

  const setEditText = (text: string, nextCaret: number, point: Point | null = null) => {
    setEdit((e) => (e ? { ...e, text, point } : e));
    setCaret(nextCaret);
    setSuggest(0);
  };

  /** Drop a reference into the formula where the caret sits, or replace the one being pointed. */
  const pointAt = (anchor: Addr, focus: Addr) => {
    const current = editRef.current;
    if (!current) {
      return;
    }
    const ref = rangeName(normalizeRange(anchor, focus));
    const start = current.point ? current.point.start : caret;
    const end = current.point ? current.point.end : caret;
    const text = current.text.slice(0, start) + ref + current.text.slice(end);
    setEditText(text, start + ref.length, { start, end: start + ref.length, anchor, focus });
    const input = current.origin === "cell" ? cellEditorRef.current : barRef.current;
    window.requestAnimationFrame(() => {
      input?.setSelectionRange(start + ref.length, start + ref.length);
    });
  };

  const canPoint = (): boolean => {
    const current = editRef.current;
    if (!current || !current.text.startsWith("=")) {
      return false;
    }
    if (current.point && caret === current.point.end) {
      return true;
    }
    return caret === current.text.length && POINTABLE.test(current.text.slice(0, caret));
  };

  /**
   * Put a function call or a name into the formula at the caret, in place of what's been typed
   * of it, or start a formula with it.
   */
  const insertAtCaret = (text: string) => {
    const current = editRef.current;
    if (current && current.text.startsWith("=")) {
      const prefix = typedFunctionPrefix(current.text, caret);
      const start = prefix ? caret - prefix.length : caret;
      const end = start + text.length;
      setEditText(`${current.text.slice(0, start)}${text}${current.text.slice(caret)}`, end);
      const input = current.origin === "cell" ? cellEditorRef.current : barRef.current;
      window.requestAnimationFrame(() => {
        input?.focus({ preventScroll: true });
        input?.setSelectionRange(end, end);
      });
      return;
    }
    startEdit(`=${text}`, "enter");
  };
  const insertFunction = (fn: string) => insertAtCaret(`${fn}(`);
  const insertSuggestion = (item: Suggestion) => insertAtCaret(item.call ? `${item.name}(` : item.name);

  /* -------------------------------------------------------- autopilot api */

  const apiImpl: WorkbookApi = {
    reveal(ref) {
      const addr = parseAddr(ref);
      if (addr) {
        ensureVisible(clampAddr(addr));
      }
    },
    cellPoint(ref) {
      const addr = parseAddr(ref);
      const box = gridRef.current?.getBoundingClientRect();
      if (!addr || !box) {
        return null;
      }
      return {
        x: box.left + ROWHEAD_W + colLeft[addr.col] + widths[addr.col] / 2,
        y: box.top + HEAD_H + addr.row * ROW_H + ROW_H / 2,
      };
    },
    fillHandlePoint() {
      const box = gridRef.current?.querySelector(".xd-fill-handle")?.getBoundingClientRect();
      return box ? { x: box.left + box.width / 2, y: box.top + box.height / 2 } : null;
    },
    select(ref) {
      const range = parseRange(ref);
      if (range) {
        setSel({
          anchor: clampAddr({ col: range.c1, row: range.r1 }),
          focus: clampAddr({ col: range.c2, row: range.r2 }),
        });
        ensureVisible(clampAddr({ col: range.c2, row: range.r2 }));
        ensureVisible(clampAddr({ col: range.c1, row: range.r1 }));
      }
    },
    setTyping(text) {
      setEdit({ addr: sel.anchor, text, mode: "enter", origin: "cell", point: null });
      setCaret(text.length);
      setSuggest(0);
    },
    commit() {
      commitEdit();
    },
    previewFill(_source, target) {
      setFillTarget(target ? parseRange(target) : null);
    },
    applyFill(source, target) {
      const from = parseRange(source);
      const to = parseRange(target);
      if (!from || !to) {
        return;
      }
      write(fillWrites(inputs, from, to));
      setSel({ anchor: { col: to.c1, row: to.r1 }, focus: { col: to.c2, row: to.r2 } });
      setFillTarget(null);
    },
    cancelAll() {
      setEdit(null);
      setFillTarget(null);
    },
  };
  // The autopilot calls through a stable object, so it always reaches the latest render's closures.
  const apiRef = useRef(apiImpl);
  apiRef.current = apiImpl;
  const stableApi = useRef<WorkbookApi>({
    reveal: (ref) => apiRef.current.reveal(ref),
    cellPoint: (ref) => apiRef.current.cellPoint(ref),
    fillHandlePoint: () => apiRef.current.fillHandlePoint(),
    select: (ref) => apiRef.current.select(ref),
    setTyping: (text) => apiRef.current.setTyping(text),
    commit: () => apiRef.current.commit(),
    previewFill: (source, target) => apiRef.current.previewFill(source, target),
    applyFill: (source, target) => apiRef.current.applyFill(source, target),
    cancelAll: () => apiRef.current.cancelAll(),
  });
  useEffect(() => {
    props.onApi(stableApi.current);
    return () => props.onApi(null);
  }, []);

  /* ------------------------------------------------------- hit testing */

  const cellAt = (clientX: number, clientY: number): Addr => {
    const rect = gridRef.current?.getBoundingClientRect();
    if (!rect) {
      return active;
    }
    const x = clientX - rect.left - ROWHEAD_W;
    const y = clientY - rect.top - HEAD_H;
    let col = colLeft.findIndex((left, i) => i < cols && x >= left && x < colLeft[i + 1]);
    if (col === -1) {
      col = x < 0 ? 0 : cols - 1;
    }
    return clampAddr({ col, row: Math.floor(y / ROW_H) });
  };

  const fillRangeFor = (source: Range, to: Addr): Range => {
    if (inRange(source, to)) {
      return source;
    }
    const down = to.row > source.r2 ? to.row - source.r2 : 0;
    const up = to.row < source.r1 ? source.r1 - to.row : 0;
    const right = to.col > source.c2 ? colLeft[to.col + 1] - colLeft[source.c2 + 1] : 0;
    const left = to.col < source.c1 ? colLeft[source.c1] - colLeft[to.col] : 0;
    // Compare in pixels so a wide column doesn't win just by being a column.
    if (Math.max(down, up) * ROW_H >= Math.max(right, left)) {
      return { ...source, r1: Math.min(source.r1, to.row), r2: Math.max(source.r2, to.row) };
    }
    return { ...source, c1: Math.min(source.c1, to.col), c2: Math.max(source.c2, to.col) };
  };

  const dragTo = (clientX: number, clientY: number) => {
    const mode = drag.current;
    if (!mode) {
      return;
    }
    const addr = cellAt(clientX, clientY);
    if (mode.kind === "select") {
      setSel((s) =>
        s.focus.col === addr.col && s.focus.row === addr.row ? s : { anchor: s.anchor, focus: addr },
      );
    } else if (mode.kind === "point") {
      const point = editRef.current?.point;
      if (point && (point.focus.col !== addr.col || point.focus.row !== addr.row)) {
        pointAt(point.anchor, addr);
      }
    } else {
      setFillTarget(fillRangeFor(mode.source, addr));
    }
  };

  // Scroll the grid while a drag sits past its edge, like Excel's auto-scroll.
  useEffect(() => {
    let frame = 0;
    const tick = () => {
      frame = window.requestAnimationFrame(tick);
      const el = scrollRef.current;
      const p = pointer.current;
      if (!drag.current || !el || !p) {
        return;
      }
      const rect = el.getBoundingClientRect();
      const dx = p.x > rect.right - 6 ? 12 : p.x < rect.left + ROWHEAD_W ? -12 : 0;
      const dy = p.y > rect.bottom - 6 ? 10 : p.y < rect.top + HEAD_H ? -10 : 0;
      if (dx || dy) {
        el.scrollLeft += dx;
        el.scrollTop += dy;
        dragTo(p.x, p.y);
      }
    };
    frame = window.requestAnimationFrame(tick);
    return () => window.cancelAnimationFrame(frame);
  });

  const onGridPointerDown = (event: TargetedPointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) {
      return;
    }
    const target = event.target as HTMLElement;
    // Clicks on the scroller itself are its scrollbars.
    if (target === scrollRef.current || target.closest(".xd-cell-editor, .xd-suggest")) {
      return;
    }
    const head = target.closest<HTMLElement>("[data-col],[data-row],[data-corner]");
    if (head) {
      if (!commitEdit()) {
        return;
      }
      event.preventDefault();
      focusGrid();
      if (head.dataset.corner !== undefined) {
        setSel({ anchor: { col: 0, row: 0 }, focus: { col: cols - 1, row: rowCount - 1 } });
      } else if (head.dataset.col !== undefined) {
        const col = Number(head.dataset.col);
        setSel({ anchor: { col, row: 0 }, focus: { col, row: rowCount - 1 } });
      } else {
        const row = Number(head.dataset.row);
        setSel({ anchor: { col: 0, row }, focus: { col: cols - 1, row } });
      }
      return;
    }
    const addr = cellAt(event.clientX, event.clientY);
    // A finger scrolls the sheet; a tap only picks the cell.
    if (event.pointerType === "touch" && !editRef.current && !target.closest(".xd-fill-handle")) {
      setSel({ anchor: addr, focus: addr });
      return;
    }
    pointer.current = { x: event.clientX, y: event.clientY };
    (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId);

    if (target.closest(".xd-fill-handle") && !editRef.current) {
      event.preventDefault();
      drag.current = { kind: "fill", source: selRange };
      setFillTarget(selRange);
      return;
    }
    if (editRef.current) {
      if (canPoint()) {
        event.preventDefault();
        drag.current = { kind: "point" };
        if (editRef.current.point && event.shiftKey) {
          pointAt(editRef.current.point.anchor, addr);
        } else {
          pointAt(addr, addr);
        }
        return;
      }
      if (!commitEdit()) {
        event.preventDefault();
        return;
      }
    }
    event.preventDefault();
    focusGrid();
    drag.current = { kind: "select" };
    setSelecting(true);
    setSel(event.shiftKey ? { anchor: sel.anchor, focus: addr } : { anchor: addr, focus: addr });
  };

  const onGridPointerMove = (event: TargetedPointerEvent<HTMLDivElement>) => {
    if (!drag.current) {
      return;
    }
    pointer.current = { x: event.clientX, y: event.clientY };
    dragTo(event.clientX, event.clientY);
  };

  const onGridPointerUp = () => {
    const mode = drag.current;
    drag.current = null;
    pointer.current = null;
    setSelecting(false);
    if (mode?.kind === "fill" && fillTarget) {
      const writes = fillWrites(inputs, mode.source, fillTarget);
      if (Object.keys(writes).length > 0) {
        write(writes);
      }
      setSel({
        anchor: { col: fillTarget.c1, row: fillTarget.r1 },
        focus: { col: fillTarget.c2, row: fillTarget.r2 },
      });
      setFillTarget(null);
    }
  };

  /* ----------------------------------------------------------- keyboard */

  const copy = (cut = false) => {
    setClip(copyRange(inputs, selRange, cut));
    internalOnly.current = false;
    copiedText.current = rangeText(selRange, (addr) => {
      const key = addrName(addr);
      // Wide enough that a number never shows as ####.
      return formatValue(values[key] ?? null, styles[key]?.format ?? "General", 10000);
    });
  };
  /**
   * The ribbon's Copy and Cut, which have no copy event to carry the text, so they write the
   * clipboard themselves. If the page won't allow that, the next paste uses the copied cells.
   */
  const copyToClipboard = (cut = false) => {
    copy(cut);
    const text = copiedText.current;
    const failed = () => {
      internalOnly.current = true;
    };
    try {
      navigator.clipboard.writeText(text).catch(failed);
    } catch {
      failed();
    }
  };
  const paste = () => {
    if (!clip) {
      return;
    }
    write(pasteWrites(clip, { col: selRange.c1, row: selRange.r1 }));
    const h = clip.range.r2 - clip.range.r1;
    const w = clip.range.c2 - clip.range.c1;
    setSel({
      anchor: { col: selRange.c1, row: selRange.r1 },
      focus: clampAddr({ col: selRange.c1 + w, row: selRange.r1 + h }),
    });
    if (clip.cut) {
      setClip(null);
    }
  };
  /**
   * Paste text from the system clipboard, copied from anywhere: one cell, or a table with tabs
   * between cells. Text simmy copied itself pastes as its cells instead, so formulas shift.
   */
  const pasteText = (text: string) => {
    if (clip && (internalOnly.current || text === copiedText.current)) {
      paste();
      return;
    }
    const pasted = textPasteWrites(text, { col: selRange.c1, row: selRange.r1 }, { rows: rowCount, cols });
    if (!pasted) {
      paste();
      return;
    }
    const writes: Inputs = {};
    for (const [key, raw] of Object.entries(pasted.writes)) {
      writes[key] = normalizeFormula(raw, names);
    }
    write(writes);
    setSel({
      anchor: { col: pasted.range.c1, row: pasted.range.r1 },
      focus: { col: pasted.range.c2, row: pasted.range.r2 },
    });
  };
  // Ctrl+V lets the browser's paste event bring the clipboard's text. Should no event come,
  // the cells simmy copied are pasted.
  const pasteFallback = useRef<ReturnType<typeof setTimeout> | null>(null);
  const onGridPaste = (event: ClipboardEvent) => {
    if (editRef.current || (event.target as Element).closest("input, textarea")) {
      return;
    }
    event.preventDefault();
    if (pasteFallback.current) {
      clearTimeout(pasteFallback.current);
      pasteFallback.current = null;
    }
    pasteText(event.clipboardData?.getData("text/plain") ?? "");
  };
  const onGridCopy = (event: ClipboardEvent, cut = false) => {
    if (editRef.current || (event.target as Element).closest("input, textarea")) {
      return;
    }
    event.preventDefault();
    copy(cut);
    event.clipboardData?.setData("text/plain", copiedText.current);
  };
  const clear = () => {
    const writes: Inputs = {};
    for (const addr of rangeAddrs(selRange)) {
      if (inputs[addrName(addr)] !== undefined) {
        writes[addrName(addr)] = "";
      }
    }
    if (Object.keys(writes).length > 0) {
      write(writes);
    }
  };
  const fill = (direction: "down" | "right") => {
    const source = direction === "down" ? { ...selRange, r2: selRange.r1 } : { ...selRange, c2: selRange.c1 };
    write(fillWrites(inputs, source, selRange));
  };

  const onGridKeyDown = (event: TargetedKeyboardEvent<HTMLDivElement>) => {
    // The cell editor sits inside the grid, so its keys bubble here, but they are its own
    // (onEditorKeyDown). By now its edit may already be committed and re-rendered, so acting
    // on them again moved two cells after typing and Enter, and put typed characters at the
    // end of the text instead of at the cursor.
    if (alert || (event.target as Element).closest(".xd-cell-editor")) {
      return;
    }
    const pending = editRef.current;
    if (pending) {
      // A key that beat the editor's focus still belongs to the edit.
      if (event.key.length === 1 && !event.ctrlKey && !event.metaKey) {
        event.preventDefault();
        setEditText(pending.text + event.key, pending.text.length + 1);
      }
      return;
    }
    const mod = event.ctrlKey || event.metaKey;
    const key = event.key;
    const handled = () => event.preventDefault();
    if (mod) {
      const lower = key.toLowerCase();
      if (lower === "z" && !event.shiftKey) {
        handled();
        undo();
      } else if (lower === "y" || (lower === "z" && event.shiftKey)) {
        handled();
        redo();
      } else if (lower === "c" || lower === "x" || lower === "v") {
        // Left to the browser's copy, cut and paste events, which reach the system clipboard.
        // The cells are copied now too, in case a browser sends no copy event.
        if (lower !== "v") {
          copy(lower === "x");
          internalOnly.current = true;
        } else {
          pasteFallback.current = setTimeout(() => {
            pasteFallback.current = null;
            paste();
          });
        }
      } else if (lower === "d") {
        handled();
        fill("down");
      } else if (lower === "r") {
        handled();
        fill("right");
      } else if (lower === "a") {
        handled();
        setSel({ anchor: { col: 0, row: 0 }, focus: { col: cols - 1, row: rowCount - 1 } });
      } else if (lower === "g") {
        handled();
        openDialog({ kind: "goto" });
      } else if (key === "F3") {
        // Ctrl+F3 is the Name Manager; Ctrl+Shift+F3 is Create from Selection.
        handled();
        openDialog({ kind: event.shiftKey ? "create" : "manager" });
      }
      return;
    }
    switch (key) {
      case "ArrowUp":
        handled();
        moveActive(0, -1, event.shiftKey);
        return;
      case "ArrowDown":
        handled();
        moveActive(0, 1, event.shiftKey);
        return;
      case "ArrowLeft":
        handled();
        moveActive(-1, 0, event.shiftKey);
        return;
      case "ArrowRight":
        handled();
        moveActive(1, 0, event.shiftKey);
        return;
      case "Enter":
      case "Tab": {
        handled();
        const step = enterTabStep(key, event.shiftKey);
        moveActive(step.dCol, step.dRow);
        tabRun.current = step.run;
        return;
      }
      case "Delete":
      case "Backspace":
        handled();
        clear();
        return;
      case "F2":
        handled();
        startEdit(inputs[addrName(active)] ?? "", "edit");
        return;
      case "F3":
        // Paste Name, when there are names to paste.
        handled();
        if (Object.keys(names).length > 0) {
          openDialog({ kind: "paste" });
        }
        return;
      case "F5":
        // Go To, and not the browser's reload.
        handled();
        openDialog({ kind: "goto" });
        return;
      case "F9":
        // F9, Shift+F9 and Ctrl+Alt+F9 all recalculate; with one sheet they do the same.
        handled();
        recalculate();
        return;
      case "Escape":
        setClip(null);
        return;
      default:
        if (key.length === 1 && !event.altKey) {
          handled();
          startEdit(key, "enter");
        }
    }
  };

  const prefix = edit ? (typedFunctionPrefix(edit.text, caret) ?? "") : "";
  const suggestions: Suggestion[] = prefix
    ? [
        ...suggestFunctions(prefix).map((fn) => ({ name: fn.name, summary: fn.summary, call: true })),
        ...Object.keys(names)
          .filter((name) => name.toUpperCase().startsWith(prefix.toUpperCase()))
          .map((name) => ({ name, summary: `Defined name: ${names[name]}`, call: false })),
      ].sort((a, b) => a.name.localeCompare(b.name))
    : [];
  const callName = edit && suggestions.length === 0 ? activeCall(edit.text, caret) : null;

  const onEditorKeyDown = (event: TargetedKeyboardEvent<HTMLInputElement>) => {
    const current = editRef.current;
    if (!current) {
      return;
    }
    const key = event.key;
    if (key === "F3" && !event.ctrlKey && !event.metaKey) {
      event.preventDefault();
      if (Object.keys(names).length > 0) {
        openDialog({ kind: "paste" });
      }
      return;
    }
    if (key === "F4") {
      // Step the reference at the caret (or the selected ones) through $A$1, A$1 and $A1.
      const input = event.currentTarget;
      const from = input.selectionStart ?? caret;
      const to = input.selectionEnd ?? caret;
      const next = cycleReference(current.text, from, to);
      if (!next) {
        return;
      }
      event.preventDefault();
      setEditText(next.text, next.end);
      // With text selected, keep it selected; otherwise the caret stays after the reference,
      // so another F4 steps the same one.
      const [start, end] = from === to ? [next.end, next.end] : [next.start, next.end];
      window.requestAnimationFrame(() => input.setSelectionRange(start, end));
      return;
    }
    if (suggestions.length > 0) {
      if (key === "ArrowDown" || key === "ArrowUp") {
        event.preventDefault();
        setSuggest((i) => (i + (key === "ArrowDown" ? 1 : suggestions.length - 1)) % suggestions.length);
        return;
      }
      if (key === "Tab") {
        event.preventDefault();
        insertSuggestion(suggestions[Math.min(suggest, suggestions.length - 1)]);
        return;
      }
    }
    if (key === "Enter" || key === "Tab") {
      event.preventDefault();
      const step = enterTabStep(key, event.shiftKey);
      if (commitAndMove(step.dCol, step.dRow)) {
        tabRun.current = step.run;
      }
    } else if (key === "Escape") {
      event.preventDefault();
      cancelEdit();
    } else if (key === "F2") {
      event.preventDefault();
      setEdit({ ...current, mode: current.mode === "enter" ? "edit" : "enter" });
    } else if (key.startsWith("Arrow") && current.mode === "enter" && current.origin === "cell") {
      event.preventDefault();
      const dCol = key === "ArrowLeft" ? -1 : key === "ArrowRight" ? 1 : 0;
      const dRow = key === "ArrowUp" ? -1 : key === "ArrowDown" ? 1 : 0;
      if (canPoint()) {
        const from = current.point
          ? event.shiftKey
            ? current.point.focus
            : current.point.anchor
          : current.addr;
        const next = clampAddr({ col: from.col + dCol, row: from.row + dRow });
        if (current.point && event.shiftKey) {
          pointAt(current.point.anchor, next);
        } else {
          pointAt(next, next);
        }
        ensureVisible(next);
      } else {
        commitAndMove(dCol, dRow);
      }
    }
  };

  const onEditorInput = (event: TargetedEvent<HTMLInputElement>) => {
    const input = event.currentTarget;
    setEditText(input.value, input.selectionStart ?? input.value.length);
  };

  const syncCaret = (event: TargetedEvent<HTMLInputElement>) => {
    const input = event.currentTarget;
    setCaret(input.selectionStart ?? input.value.length);
  };

  /* ------------------------------------------------------ name box / bar */

  const [nameDraft, setNameDraft] = useState<string | null>(null);
  const selectRange = (range: Range) => {
    const start = clampAddr({ col: range.c1, row: range.r1 });
    const end = clampAddr({ col: range.c2, row: range.r2 });
    setSel({ anchor: start, focus: end });
    ensureVisible(end);
    ensureVisible(start);
  };

  /** The range a reference, a spill (A1#) or a defined name stands for, or null. */
  const resolveReference = (text: string): Range | null => {
    const trimmed = text.trim();
    const spill = /^(\$?[A-Za-z]{1,3}\$?\d{1,7})#$/.exec(trimmed);
    if (spill) {
      const anchor = parseAddr(spill[1]);
      return anchor ? (sheetEval.spills[addrName(anchor)] ?? null) : null;
    }
    const named = lookupName(names, trimmed);
    return named ? nameRange(named.ref) : parseRange(trimmed.toUpperCase());
  };

  const goTo = (text: string) => {
    const range = resolveReference(text);
    if (range) {
      selectRange(range);
    }
    setNameDraft(null);
    focusGrid();
  };

  /** Enter in the Name Box: go to a reference or a name, or give the selection a new name. */
  const nameBoxEnter = (text: string) => {
    const trimmed = text.trim();
    const range = resolveReference(trimmed);
    if (range) {
      goTo(trimmed);
      return;
    }
    if (nameProblem(trimmed) === null) {
      apply({ names: { [trimmed]: absoluteRef(selRange) } });
      setNameDraft(null);
      focusGrid();
      return;
    }
    setAlert({
      message: "You must enter a valid reference you want to go to, or type a valid name for the selection.",
      from: "namebox",
    });
  };

  const selectFromPane = (ref: string) => {
    if (editRef.current && !commitEdit()) {
      return;
    }
    goTo(ref);
  };

  /**
   * The Name Manager's Value column: a cell's value, or for several cells Excel's array form,
   * {"a",1;"b",2}, cut short when it runs long.
   */
  const nameValue = (ref: string): string => {
    const range = nameRange(ref);
    if (!range) {
      return "";
    }
    const show = (value: Value | undefined) =>
      typeof value === "string" ? `"${value}"` : formatValue(value ?? null, "General", 200);
    if (range.c1 === range.c2 && range.r1 === range.r2) {
      return formatValue(values[addrName({ col: range.c1, row: range.r1 })] ?? null, "General", 200);
    }
    const rows = [];
    for (let row = range.r1; row <= range.r2; row++) {
      const line = [];
      for (let col = range.c1; col <= range.c2; col++) {
        line.push(show(values[addrName({ col, row })]));
      }
      rows.push(line.join(","));
    }
    const text = `{${rows.join(";")}}`;
    return text.length > 60 ? `${text.slice(0, 59)}…` : text;
  };

  /* ------------------------------------------------------------- render */

  const refRanges = edit ? formulaRanges(edit.text, names) : [];
  const activeKey = addrName(active);
  // A cell filled by a spill shows its anchor's formula, greyed, as Excel does.
  const spillAnchor = sheetEval.owner[activeKey];
  const barText = edit ? edit.text : (inputs[activeKey] ?? (spillAnchor ? inputs[spillAnchor] : ""));
  const ghost = !edit && spillAnchor !== undefined;
  // The active cell's spill: its range while it fills, or where it would go while blocked.
  const spillRange = sheetEval.spills[spillAnchor ?? activeKey];
  const blockedRange = sheetEval.blocked[activeKey];
  const mode = alert
    ? "Ready"
    : edit
      ? edit.point || drag.current?.kind === "point"
        ? "Point"
        : edit.mode === "edit"
          ? "Edit"
          : "Enter"
      : "Ready";

  const stats = useMemo(() => {
    if (selRange.c1 === selRange.c2 && selRange.r1 === selRange.r2) {
      return null;
    }
    let count = 0;
    const nums: number[] = [];
    for (const addr of rangeAddrs(selRange)) {
      const value = values[addrName(addr)];
      if (value === null || value === undefined || value === "") {
        continue;
      }
      count++;
      if (typeof value === "number") {
        nums.push(value);
      }
    }
    if (count === 0) {
      return null;
    }
    if (nums.length === 0) {
      return { count };
    }
    const sum = nums.reduce((a, b) => a + b, 0);
    return { count, sum, average: sum / nums.length };
  }, [selRange.c1, selRange.c2, selRange.r1, selRange.r2, values]);

  // Overlays sit on the gridlines: one pixel up and left, one pixel wider and taller than the cells.
  const box = (range: Range) => ({
    left: ROWHEAD_W + colLeft[range.c1] - 1,
    top: HEAD_H + range.r1 * ROW_H - 1,
    width: colLeft[range.c2 + 1] - colLeft[range.c1] + 1,
    height: (range.r2 - range.r1 + 1) * ROW_H + 1,
  });

  const rows = [];
  for (let row = 0; row < rowCount; row++) {
    const cells = [];
    for (let col = 0; col < cols; col++) {
      const addr = { col, row };
      const key = addrName(addr);
      const style = styles[key];
      const value: Value = values[key] ?? null;
      const text = formatValue(value, style?.format ?? "General", widths[col]);
      const numeric = typeof value === "number";
      // Accounting pins the $ to the cell's left edge and the amount to the right.
      const accounting = numeric && style?.format === "accounting" && text.startsWith("$");
      const align =
        style?.align === "center"
          ? "c"
          : numeric
            ? "r"
            : isError(value) || typeof value === "boolean"
              ? "c"
              : "l";
      const selected = inRange(selRange, addr) && !(col === active.col && row === active.row);
      // Text runs into the next cell only when that cell is empty, as in Excel.
      const next = values[addrName({ col: col + 1, row })];
      const overflows = text !== "" && !numeric && align === "l" && (next === null || next === undefined);
      const className = [
        "xd-cell",
        `xd-a-${align}`,
        style?.bold ? "is-b" : "",
        style?.italic ? "is-i" : "",
        style?.lined ? "is-lined" : "",
        style?.fill ? "is-filled" : "",
        selected ? "is-sel" : "",
        overflows ? "is-overflow" : "",
        accounting ? "is-acct" : "",
      ]
        .filter(Boolean)
        .join(" ");
      cells.push(
        <div
          key={key}
          class={className}
          style={{ width: widths[col], background: style?.fill, color: style?.color }}
        >
          {accounting ? (
            <>
              <span>$</span>
              <span>{text.slice(1).trimStart()}</span>
            </>
          ) : (
            text
          )}
        </div>,
      );
    }
    const rowSel = row >= selRange.r1 && row <= selRange.r2;
    rows.push(
      <div class="xd-row-line" key={row}>
        <div class={`xd-rowhead${rowSel ? " is-sel" : ""}`} data-row={row}>
          {row + 1}
        </div>
        {cells}
      </div>,
    );
  }

  const editBox = edit ? box(normalizeRange(edit.addr, edit.addr)) : null;
  const editStyle = edit ? styles[addrName(edit.addr)] : undefined;

  const toggleRibbon = () => {
    setRibbonCollapsed(!ribbonCollapsed);
    setRibbonPeek(false);
  };
  const chooseTab = (next: RibbonTab) => {
    setTab(next);
    if (ribbonCollapsed) {
      // A second click on the tab that is showing puts the ribbon away again.
      setRibbonPeek(!(ribbonPeek && next === tab));
    }
  };
  /** A ribbon command. One run from a ribbon shown over the sheet puts it away, as in Excel. */
  const command =
    <A extends unknown[]>(run: (...args: A) => void) =>
    (...args: A) => {
      setRibbonPeek(false);
      run(...args);
    };
  const ribbonShown = !ribbonCollapsed || ribbonPeek;

  return (
    <div
      class={`xd-excel xd-${platform}`}
      onPointerDownCapture={(event) => {
        if (ribbonPeek && !(event.target instanceof Element && event.target.closest(".xd-ribbon-area"))) {
          setRibbonPeek(false);
        }
      }}
      onKeyDownCapture={(event) => {
        if (event.key === "F1" && (event.ctrlKey || event.metaKey)) {
          event.preventDefault();
          toggleRibbon();
        } else if (event.key === "Escape" && ribbonPeek) {
          setRibbonPeek(false);
        }
      }}
    >
      <TitleBar
        {...props}
        canUndo={book.past.length > 0}
        canRedo={book.future.length > 0}
        onUndo={undo}
        onRedo={redo}
      />
      <div class={`xd-ribbon-area${ribbonCollapsed ? " is-collapsed" : ""}${ribbonPeek ? " is-peek" : ""}`}>
        <RibbonTabs
          platform={platform}
          tab={ribbonShown ? tab : null}
          onTab={chooseTab}
          onToggleRibbon={toggleRibbon}
        />
        {ribbonShown && (
          <Ribbon
            platform={platform}
            tab={tab}
            onTab={setTab}
            numberFormat={formatName(styles[activeKey])}
            paneOpen={paneOpen}
            collapsed={ribbonCollapsed}
            onToggleRibbon={toggleRibbon}
            onTogglePane={command(() => setPaneOpen(!paneOpen))}
            onInsertFunction={command(insertFunction)}
            onStartFormula={command(() => (edit ? undefined : startEdit("=", "enter")))}
            onCopy={command(() => copyToClipboard())}
            onCut={command(() => copyToClipboard(true))}
            onPaste={command(paste)}
            onClear={command(clear)}
            onFillDown={command(() => fill("down"))}
            onCalculateNow={command(recalculate)}
            onGoTo={command(() => openDialog({ kind: "goto" }))}
            names={Object.keys(names)}
            editing={edit !== null}
            onNameManager={command(() => openDialog({ kind: "manager" }))}
            onDefineName={command(() => openDialog({ kind: "define" }))}
            onCreateFromSelection={command(() => openDialog({ kind: "create" }))}
            onUseInFormula={(name) => insertAtCaret(name)}
            onPasteNames={() => openDialog({ kind: "paste" })}
          />
        )}
      </div>
      <div class="xd-formula-bar">
        <div class="xd-namebox">
          <input
            ref={nameBoxRef}
            aria-label="Name Box"
            value={
              nameDraft ??
              (selecting && stats
                ? `${selRange.r2 - selRange.r1 + 1}R x ${selRange.c2 - selRange.c1 + 1}C`
                : (nameForRange(names, selRange) ?? activeKey))
            }
            onFocus={(event) => {
              setNameDraft(event.currentTarget.value);
              event.currentTarget.select();
            }}
            onInput={(event) => setNameDraft(event.currentTarget.value)}
            onBlur={() => setNameDraft(null)}
            onKeyDown={(event) => {
              if (event.key === "Enter") {
                event.preventDefault();
                nameBoxEnter(event.currentTarget.value);
              } else if (event.key === "Escape") {
                setNameDraft(null);
                focusGrid();
              }
            }}
          />
          <Menu
            trigger={(open, toggle) => (
              <button
                type="button"
                class={`xd-namebox-arrow${open ? " is-on" : ""}`}
                title="Name Box"
                onMouseDown={(event) => event.preventDefault()}
                onClick={toggle}
              >
                <Svg w={12}>
                  <path d="M5 8l5 5 5-5" fill="none" stroke="#616161" stroke-width="1.6" />
                </Svg>
              </button>
            )}
          >
            {(close) =>
              Object.keys(names)
                .sort((a, b) => a.localeCompare(b))
                .map((name) => (
                  <button
                    key={name}
                    type="button"
                    class="xd-menu-item"
                    onClick={() => {
                      close();
                      goTo(name);
                    }}
                  >
                    {name}
                  </button>
                ))
            }
          </Menu>
        </div>
        <span class="xd-fb-dots">⋮</span>
        <button
          type="button"
          class="xd-fb-btn"
          title="Cancel"
          disabled={!edit}
          onMouseDown={(e) => e.preventDefault()}
          onClick={cancelEdit}
        >
          <Svg w={14}>
            <path d="M5 5l10 10M15 5L5 15" stroke="currentColor" stroke-width="1.6" />
          </Svg>
        </button>
        <button
          type="button"
          class="xd-fb-btn"
          title="Enter"
          disabled={!edit}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => commitAndMove(0, 0)}
        >
          <Svg w={14}>
            <path d="M4 10.5l4 4 8-9" fill="none" stroke="currentColor" stroke-width="1.6" />
          </Svg>
        </button>
        <button
          type="button"
          class="xd-fb-btn xd-fx"
          title="Insert Function"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => (edit ? undefined : startEdit("=", "enter"))}
        >
          fx
        </button>
        <input
          ref={barRef}
          class={`xd-fb-input${ghost ? " is-ghost" : ""}`}
          aria-label="Formula Bar"
          value={barText}
          spellcheck={false}
          autocomplete="off"
          onFocus={() => {
            if (!editRef.current) {
              startEdit(inputs[activeKey] ?? "", "edit", "bar");
            } else if (editRef.current.origin !== "bar") {
              setEdit({ ...editRef.current, origin: "bar", mode: "edit" });
            }
          }}
          onInput={onEditorInput}
          onKeyDown={onEditorKeyDown}
          onKeyUp={syncCaret}
          onClick={syncCaret}
        />
      </div>
      <div class="xd-body">
        <div class="xd-sheet-area">
          <div
            class="xd-grid-scroll"
            ref={scrollRef}
            tabIndex={0}
            role="grid"
            aria-label={`${mod.workbook.sheet} worksheet`}
            aria-rowcount={rowCount}
            aria-colcount={cols}
            onPointerDown={onGridPointerDown}
            onPointerMove={onGridPointerMove}
            onPointerUp={onGridPointerUp}
            onPointerCancel={onGridPointerUp}
            onDblClick={(event) => {
              const target = event.target as HTMLElement;
              if (target.closest("[data-col],[data-row],.xd-cell-editor")) {
                return;
              }
              if (!editRef.current) {
                startEdit(inputs[activeKey] ?? "", "edit");
              }
            }}
            onKeyDown={onGridKeyDown}
            onCopy={(event) => onGridCopy(event)}
            onCut={(event) => onGridCopy(event, true)}
            onPaste={onGridPaste}
          >
            <div
              class="xd-grid"
              ref={gridRef}
              style={{ width: ROWHEAD_W + sheetW, height: HEAD_H + rowCount * ROW_H }}
            >
              <div class="xd-colheads">
                <div class="xd-corner" data-corner="">
                  <svg viewBox="0 0 10 10" width="10" height="10" aria-hidden="true">
                    <path d="M10 0v10H0z" fill="#b9b9b9" />
                  </svg>
                </div>
                {widths.map((width, col) => (
                  <div
                    key={col}
                    data-col={col}
                    class={`xd-colhead${col >= selRange.c1 && col <= selRange.c2 ? " is-sel" : ""}`}
                    style={{ width }}
                  >
                    {colName(col)}
                  </div>
                ))}
              </div>
              {rows}
              <div class="xd-layer">
                {outlines.map((outline, i) => (
                  <div
                    key={i}
                    class="xd-outline"
                    style={{ ...box(outline.range), borderWidth: outline.weight, borderColor: outline.color }}
                  />
                ))}
                {refRanges.map((range, i) => (
                  <div
                    key={`ref-${i}`}
                    class="xd-refbox"
                    style={{ ...box(range), "--ref": REF_COLORS[i % REF_COLORS.length] }}
                  />
                ))}
                {spillRange && <div class="xd-spillbox" style={box(spillRange)} />}
                {blockedRange && <div class="xd-spillbox is-blocked" style={box(blockedRange)} />}
                {clip && <div class="xd-ants" style={box(clip.range)} />}
                {!edit && (
                  <div class="xd-selbox" style={box(selRange)}>
                    <span class="xd-fill-handle" />
                  </div>
                )}
                {fillTarget && <div class="xd-fillbox" style={box(fillTarget)} />}
              </div>
              {edit && editBox && (
                <>
                  <input
                    ref={cellEditorRef}
                    class={`xd-cell-editor${editStyle?.bold ? " is-b" : ""}`}
                    value={edit.text}
                    spellcheck={false}
                    autocomplete="off"
                    aria-label={`Edit ${addrName(edit.addr)}`}
                    readOnly={edit.origin === "bar"}
                    style={{
                      left: editBox.left,
                      top: editBox.top,
                      minWidth: editBox.width + 1,
                      width: Math.max(editBox.width + 1, edit.text.length * 7.6 + 14),
                      height: ROW_H + 1,
                      background: editStyle?.fill ?? "#fff",
                    }}
                    onFocus={() => {
                      if (editRef.current?.origin === "bar") {
                        setEdit({ ...editRef.current, origin: "cell" });
                      }
                    }}
                    onInput={onEditorInput}
                    onKeyDown={onEditorKeyDown}
                    onKeyUp={syncCaret}
                    onClick={syncCaret}
                  />
                  {suggestions.length > 0 && (
                    <div class="xd-suggest" style={{ left: editBox.left, top: editBox.top + ROW_H + 2 }}>
                      <ul>
                        {suggestions.map((fn, i) => (
                          <li key={fn.name}>
                            <button
                              type="button"
                              class={i === suggest ? "is-on" : ""}
                              onMouseDown={(event) => event.preventDefault()}
                              onClick={() => insertSuggestion(fn)}
                            >
                              <span class="xd-fn-icon">{fn.call ? "ƒ" : "▭"}</span>
                              {fn.name}
                            </button>
                          </li>
                        ))}
                      </ul>
                      <p class="xd-suggest-tip">
                        {suggestions[Math.min(suggest, suggestions.length - 1)].summary}
                      </p>
                    </div>
                  )}
                  {callName && (
                    <div class="xd-argtip" style={{ left: editBox.left, top: editBox.top + ROW_H + 3 }}>
                      <b>{callName}</b>({SIGNATURES[callName]})
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
          <SheetTabs name={mod.workbook.sheet} />
        </div>
        {paneOpen && (
          <TaskPane
            key={paneKey}
            state={{ ...book.state, selection: rangeName(selRange) }}
            onSelect={selectFromPane}
            onChange={apply}
            onClose={() => setPaneOpen(false)}
            onUnsupported={props.onUnsupported}
          />
        )}
      </div>
      <div class="xd-status">
        <span class="xd-mode">{mode}</span>
        <span class="xd-a11y">
          <Svg w={14}>
            <circle cx="10" cy="4" r="1.6" fill="currentColor" />
            <path d="M4 7h12M10 7v5M7 18l3-6 3 6" fill="none" stroke="currentColor" stroke-width="1.4" />
          </Svg>
          Accessibility: Good to go
        </span>
        <span class="xd-status-spacer" />
        {stats && (
          <span class="xd-stats">
            {stats.average !== undefined && <span>Average: {statusNumber(stats.average)}</span>}
            <span>Count: {stats.count}</span>
            {stats.sum !== undefined && <span>Sum: {statusNumber(stats.sum)}</span>}
          </span>
        )}
        <span class="xd-views is-idle" aria-hidden="true">
          <span class="is-on">
            <Svg w={14}>
              <rect x="3" y="4" width="14" height="12" fill="none" stroke="currentColor" />
              <path d="M3 8h14M8 4v12M13 4v12" stroke="currentColor" />
            </Svg>
          </span>
          <span>
            <Svg w={14}>
              <rect x="5" y="3" width="10" height="14" fill="none" stroke="currentColor" />
            </Svg>
          </span>
          <span>
            <Svg w={14}>
              <rect
                x="3"
                y="3"
                width="14"
                height="14"
                fill="none"
                stroke="currentColor"
                stroke-dasharray="2 1.5"
              />
            </Svg>
          </span>
        </span>
        <span class="xd-zoom is-idle" aria-hidden="true">
          <span>−</span>
          <span class="xd-zoom-track">
            <span class="xd-zoom-thumb" />
          </span>
          <span>+</span>
          <span class="xd-zoom-pct">100%</span>
        </span>
      </div>
      {dialogs.map((dialog, i) => (
        <div key={i} class="xd-dialog-layer" inert={i < dialogs.length - 1}>
          {dialog.kind === "define" && (
            <DefineNameDialog
              title={dialog.editing ? "Edit Name" : "New Name"}
              initial={
                dialog.editing
                  ? {
                      name: dialog.editing,
                      ref: refersToText(nameRange(names[dialog.editing] ?? "") ?? selRange, mod.workbook.sheet),
                      comment: book.state.nameComments?.[dialog.editing] ?? "",
                    }
                  : {
                      name: suggestName(inputs, selRange),
                      ref: refersToText(selRange, mod.workbook.sheet),
                      comment: "",
                    }
              }
              names={names}
              editing={dialog.editing}
              sheet={mod.workbook.sheet}
              onCancel={closeDialog}
              onSave={(entry) => {
                const old = dialog.editing;
                const renaming = old !== undefined && old !== entry.name;
                // Close first: the dialog mustn't render again for a name that's gone.
                closeDialog();
                apply({
                  // As in Excel, formulas that used the old name now use the new one.
                  cells: renaming ? renameWrites(inputs, old, entry.name) : undefined,
                  names: { ...(renaming ? { [old]: null } : {}), [entry.name]: entry.ref },
                  nameComments: { [entry.name]: entry.comment },
                });
              }}
            />
          )}
          {dialog.kind === "create" && (
            <CreateNamesDialog
              initial={guessLabelSides(inputs, selRange)}
              onCancel={closeDialog}
              onCreate={(sides) => {
                closeDialog();
                apply({ names: namesFromSelection(inputs, selRange, sides) });
              }}
            />
          )}
          {dialog.kind === "manager" && (
            <NameManagerDialog
              names={names}
              comments={book.state.nameComments ?? {}}
              valueOf={nameValue}
              sheet={mod.workbook.sheet}
              onNew={() => openDialog({ kind: "define" })}
              onEdit={(name) => openDialog({ kind: "define", editing: name })}
              onDelete={(name) => apply({ names: { [name]: null } })}
              onClose={closeDialog}
            />
          )}
          {dialog.kind === "paste" && (
            <PasteNameDialog
              names={names}
              onCancel={closeDialog}
              onPaste={(name) => {
                closeDialog();
                insertAtCaret(name);
              }}
            />
          )}
          {dialog.kind === "goto" && (
            <GoToDialog
              names={names}
              onCancel={closeDialog}
              onGo={(text) => {
                const range = resolveReference(text);
                if (range) {
                  selectRange(range);
                  closeDialog();
                }
                return range !== null;
              }}
            />
          )}
        </div>
      ))}
      {alert && (
        <div class="xd-alert-veil">
          <div class="xd-alert" role="alertdialog" aria-label="Microsoft Excel">
            <div class="xd-alert-title">Microsoft Excel</div>
            <div class="xd-alert-body">
              <span class="xd-alert-icon">!</span>
              <div>
                <p>{alert.message}</p>
                {alert.formula && (
                  <p class="xd-alert-help">
                    Not trying to type a formula? When the first character is an equal (=) or minus (-) sign,
                    Excel thinks it's a formula.
                  </p>
                )}
              </div>
            </div>
            <div class="xd-alert-actions">
              <button
                type="button"
                class="xd-alert-ok"
                ref={alertOkRef}
                onClick={dismissAlert}
                onKeyDown={(event) => {
                  if (event.key === "Escape" || event.key === "Enter") {
                    event.preventDefault();
                    dismissAlert();
                  }
                }}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function SheetTabs({ name }: { name: string }) {
  return (
    <div class="xd-sheet-tabs">
      <span class="xd-sheet-nav is-idle" aria-hidden="true">
        <Svg w={12}>
          <path d="M12 5l-5 5 5 5" fill="none" stroke="#9b9b9b" stroke-width="1.6" />
        </Svg>
        <Svg w={12}>
          <path d="M8 5l5 5-5 5" fill="none" stroke="#9b9b9b" stroke-width="1.6" />
        </Svg>
      </span>
      <span class="xd-sheet-tab is-active">{name}</span>
      <span class="xd-sheet-add is-idle" title="New sheet (not available here)" aria-hidden="true">
        <Svg w={14}>
          <circle cx="10" cy="10" r="7.5" fill="none" stroke="#616161" />
          <path d="M10 6.5v7M6.5 10h7" stroke="#616161" />
        </Svg>
      </span>
    </div>
  );
}

type TitleBarProps = WorkbookProps & {
  canUndo: boolean;
  canRedo: boolean;
  onUndo: () => void;
  onRedo: () => void;
};

function TitleBar(props: TitleBarProps) {
  const { platform } = props;
  const workbookName = useSim().module.workbook.name;
  const stop = (event: Event) => event.stopPropagation();
  const undoIcon = (
    <Svg w={16}>
      <path d="M7 5L3 9l4 4M3 9h9a5 5 0 010 10h-2" fill="none" stroke="currentColor" stroke-width="1.4" />
    </Svg>
  );
  const redoIcon = (
    <Svg w={16}>
      <path d="M13 5l4 4-4 4M17 9H8a5 5 0 000 10h2" fill="none" stroke="currentColor" stroke-width="1.4" />
    </Svg>
  );
  const quick = (
    <span class="xd-quick" onPointerDown={stop}>
      <span class="xd-autosave is-idle">
        AutoSave
        <span class="xd-switch" />
      </span>
      <span class="xd-qbtn is-idle" title="Save (not available here)">
        <Svg w={16}>
          <path
            d="M4 3h10l3 3v11H4z M7 3v4h6V3 M7 17v-5h7v5"
            fill="none"
            stroke="currentColor"
            stroke-width="1.2"
          />
        </Svg>
      </span>
      <button type="button" class="xd-qbtn" title="Undo" disabled={!props.canUndo} onClick={props.onUndo}>
        {undoIcon}
      </button>
      <button type="button" class="xd-qbtn" title="Redo" disabled={!props.canRedo} onClick={props.onRedo}>
        {redoIcon}
      </button>
    </span>
  );
  const undoRedo = (
    <span class="xd-quick" onPointerDown={stop}>
      <button type="button" class="xd-qbtn" title="Undo" disabled={!props.canUndo} onClick={props.onUndo}>
        {undoIcon}
      </button>
      <button type="button" class="xd-qbtn" title="Redo" disabled={!props.canRedo} onClick={props.onRedo}>
        {redoIcon}
      </button>
    </span>
  );
  // The web view fills the stage, so there is no window to move, shrink or close: just the app's
  // name, the workbook and Undo, nothing an OS or a Microsoft account adds.
  if (platform === "web") {
    return (
      <div class="xd-titlebar">
        <span class="xd-app-glyph" aria-hidden="true">
          <ExcelGlyph size={16} />
        </span>
        {undoRedo}
        <span class="xd-title">{workbookName}</span>
      </div>
    );
  }
  if (platform === "mac") {
    return (
      <div class="xd-titlebar" onPointerDown={props.onTitlePointerDown} onDblClick={props.onMaximize}>
        <span class="xd-lights" onPointerDown={stop}>
          <button type="button" class="xd-light is-close" title="Close" onClick={props.onClose} />
          <button type="button" class="xd-light is-min" title="Minimize" onClick={props.onMinimize} />
          <button type="button" class="xd-light is-max" title="Zoom" onClick={props.onMaximize} />
        </span>
        {quick}
        <span class="xd-title">{workbookName}</span>
        <span class="xd-search is-idle" onPointerDown={stop}>
          <Svg w={13}>
            <circle cx="8" cy="8" r="5" fill="none" stroke="currentColor" stroke-width="1.5" />
            <path d="M12 12l5 5" stroke="currentColor" stroke-width="1.5" />
          </Svg>
          Search (Cmd + Ctrl + U)
        </span>
      </div>
    );
  }
  return (
    <div class="xd-titlebar" onPointerDown={props.onTitlePointerDown} onDblClick={props.onMaximize}>
      <span class="xd-app-glyph" aria-hidden="true">
        <ExcelGlyph size={16} />
      </span>
      {quick}
      <span class="xd-title">
        {workbookName} • Saved to this PC
        <Svg w={10}>
          <path d="M5 8l5 5 5-5" fill="none" stroke="currentColor" stroke-width="1.6" />
        </Svg>
      </span>
      <span class="xd-search is-idle" onPointerDown={stop}>
        <Svg w={13}>
          <circle cx="8" cy="8" r="5" fill="none" stroke="currentColor" stroke-width="1.5" />
          <path d="M12 12l5 5" stroke="currentColor" stroke-width="1.5" />
        </Svg>
        Search (Alt+Q)
      </span>
      <span class="xd-avatar" aria-hidden="true">
        EE
      </span>
      <span class="xd-caption" onPointerDown={stop}>
        <button type="button" title="Minimize" onClick={props.onMinimize}>
          <Svg w={10}>
            <path d="M2 10h16" stroke="currentColor" stroke-width="1.4" />
          </Svg>
        </button>
        <button
          type="button"
          title={props.maximized ? "Restore Down" : "Maximize"}
          onClick={props.onMaximize}
        >
          <Svg w={10}>
            {props.maximized ? (
              <path
                d="M2.5 6.5h11v11h-11zM6.5 6V2.5h11v11H14"
                fill="none"
                stroke="currentColor"
                stroke-width="1.4"
              />
            ) : (
              <rect
                x="2.5"
                y="2.5"
                width="15"
                height="15"
                fill="none"
                stroke="currentColor"
                stroke-width="1.4"
              />
            )}
          </Svg>
        </button>
        <button type="button" class="is-close" title="Close" onClick={props.onClose}>
          <Svg w={10}>
            <path d="M2 2l16 16M18 2L2 18" stroke="currentColor" stroke-width="1.4" />
          </Svg>
        </button>
      </span>
    </div>
  );
}

/** A generic spreadsheet app tile: green, a grid, an X. */
export function ExcelGlyph({ size = 32 }: { size?: number }) {
  return (
    <svg viewBox="0 0 32 32" width={size} height={size} aria-hidden="true">
      <rect x="8" y="3" width="21" height="26" rx="2.5" fill="#21a366" />
      <path d="M8 11.5h21M8 20.5h21M18.5 3v26" stroke="#107c41" stroke-width="1.2" />
      <rect x="2.5" y="8.5" width="15" height="15" rx="2" fill="#107c41" />
      <path d="M6.5 12.5l7 7M13.5 12.5l-7 7" stroke="#fff" stroke-width="2.2" stroke-linecap="round" />
    </svg>
  );
}
