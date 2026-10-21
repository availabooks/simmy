import type { ComponentChildren, CSSProperties, JSX, Ref, TargetedPointerEvent } from "preact";
import { useEffect, useMemo, useRef, useState } from "preact/hooks";
import { addrName, evaluateSheet, formatValue, isError, rangeAddrs } from "../engine/sheet.ts";
import { solvedInputs } from "../module/grade.ts";
import type { NoteMessage, SimModule, UnsupportedTool } from "../module/types.ts";
import { runAutopilot, type WorkbookApi } from "./autopilot.ts";
import { useSim } from "./context.ts";
import type { Platform } from "./Ribbon.tsx";
import { ExcelGlyph, Workbook } from "./Workbook.tsx";

type WinId = "excel" | "notes";
type Win = {
  x: number;
  y: number;
  w: number;
  h: number;
  open: boolean;
  min: boolean;
  max: boolean;
};
type Wins = Record<WinId, Win>;

const MIN_SIZE: Record<WinId, { w: number; h: number }> = {
  excel: { w: 640, h: 420 },
  notes: { w: 170, h: 130 },
};

/** Space the menu bar and dock (or the taskbar) take from the desktop. */
function insets(platform: Platform) {
  return platform === "mac" ? { top: 26, bottom: 74 } : { top: 0, bottom: 48 };
}

function initialWins(platform: Platform, width: number, height: number): Wins {
  const { top, bottom } = insets(platform);
  const excelH = height - top - bottom - 20;
  return {
    excel: {
      x: 20,
      y: top + 12,
      w: width - 40,
      h: excelH,
      open: true,
      min: false,
      max: false,
    },
    // Over the empty columns right of the order table, clear of the task pane.
    notes: {
      x: Math.max(34, Math.min(420, width - 330 - 236 - 40)),
      y: top + excelH - 236,
      w: 236,
      h: 210,
      open: true,
      min: false,
      max: false,
    },
  };
}

export function detectPlatform(nav: {
  platform?: string;
  userAgent?: string;
  userAgentData?: { platform?: string };
}): Platform {
  const hint = `${nav.userAgentData?.platform ?? ""} ${nav.platform ?? ""} ${nav.userAgent ?? ""}`;
  return /mac|iphone|ipad|ipod/i.test(hint) ? "mac" : "win";
}

function formatClock(date: Date, platform: Platform): { time: string; date: string } {
  const time = date.toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
  });
  if (platform === "mac") {
    const day = date.toLocaleDateString("en-US", {
      weekday: "short",
      month: "short",
      day: "numeric",
    });
    return { time, date: day.replace(",", "") };
  }
  return { time, date: date.toLocaleDateString("en-US") };
}

/** Set inline styles on a page element and hand back a function that puts the old ones back. */
function setPageStyles(el: HTMLElement, styles: Record<string, string>): () => void {
  const old = Object.fromEntries(Object.keys(styles).map((key) => [key, el.style.getPropertyValue(key)]));
  for (const [key, value] of Object.entries(styles)) {
    el.style.setProperty(key, value);
  }
  return () => {
    for (const [key, value] of Object.entries(old)) {
      el.style.setProperty(key, value);
    }
  };
}

/** While a window drags, the page behind it shouldn't select text. */
function lockSelection(): () => void {
  return setPageStyles(document.body, { "user-select": "none", "-webkit-user-select": "none" });
}

/** The phone layout: a small screen, or a landscape phone too short for a desktop. */
const PHONE_QUERY = "(max-width: 759px), (max-height: 520px) and (pointer: coarse)";

export type SimmyProps = {
  /** The desktop to show. "auto" (the default) matches the visitor's computer. */
  platform?: Platform | "auto";
  /** Run the auto demo until the visitor takes control. Off for people who ask for less motion. */
  autoplay?: boolean;
  /** On a phone, let the browser's Back button close the full-screen view. */
  history?: boolean;
  /** Size the desktop to the screen, less the host's sticky header (the `--chrome-h` property). */
  fitScreen?: boolean;
  /** The host's own heading, shown above the desktop beside the toolbar. */
  head?: ComponentChildren;
};

export function Simmy({ platform: platformProp = "auto", autoplay = true, history = false, fitScreen = false, head }: SimmyProps) {
  const { module: mod } = useSim();
  const [platform, setPlatform] = useState<Platform>(platformProp === "auto" ? "win" : platformProp);
  const [size, setSize] = useState<{ w: number; h: number } | null>(null);
  const [wins, setWins] = useState<Wins | null>(null);
  const [order, setOrder] = useState<WinId[]>(["excel", "notes"]);
  const [resetKey, setResetKey] = useState(0);
  const [pinned, setPinned] = useState(true);
  const [note, setNote] = useState<NoteMessage | null>(null);
  const [clock, setClock] = useState<Date | null>(null);
  const [auto, setAuto] = useState(true);
  const [apiReady, setApiReady] = useState(false);
  const [phone, setPhone] = useState(false);
  const [open, setOpen] = useState(false);
  const [calm, setCalm] = useState(false);
  const desktopRef = useRef<HTMLDivElement>(null);
  const stageRef = useRef<HTMLDivElement>(null);
  const posterRef = useRef<HTMLButtonElement>(null);
  const cursorRef = useRef<HTMLDivElement>(null);
  const apiRef = useRef<WorkbookApi | null>(null);
  const onScreen = useRef(true);

  useEffect(() => {
    setPlatform(
      platformProp === "auto"
        ? detectPlatform(navigator as Navigator & { userAgentData?: { platform?: string } })
        : platformProp,
    );
  }, [platformProp]);

  useEffect(() => {
    setClock(new Date());
    const timer = window.setInterval(() => setClock(new Date()), 30_000);
    return () => window.clearInterval(timer);
  }, []);

  // Auto mode is the default. People who ask for less motion start in control.
  useEffect(() => {
    const calm = !autoplay || window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    setCalm(calm);
    setAuto(!calm);
  }, [autoplay]);

  useEffect(() => {
    const query = window.matchMedia(PHONE_QUERY);
    const sync = () => setPhone(query.matches);
    sync();
    query.addEventListener("change", sync);
    return () => query.removeEventListener("change", sync);
  }, []);

  // On a phone the demo is a full-screen view: lock the page behind it and let Back close it.
  useEffect(() => {
    if (!phone || !open) {
      return;
    }
    const unlockRoot = setPageStyles(document.documentElement, { overflow: "hidden" });
    const unlockBody = setPageStyles(document.body, { overflow: "hidden" });
    const onPop = () => setOpen(false);
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        closeDemo();
      }
    };
    window.addEventListener("popstate", onPop);
    window.addEventListener("keydown", onKey);
    return () => {
      unlockRoot();
      unlockBody();
      window.removeEventListener("popstate", onPop);
      window.removeEventListener("keydown", onKey);
    };
  }, [phone, open]);

  const openDemo = () => {
    if (history) {
      window.history.pushState({ xdDemo: true }, "");
    }
    setNote(null);
    setResetKey((k) => k + 1);
    setAuto(!calm);
    setOpen(true);
  };
  const closeDemo = () => {
    if (history && window.history.state?.xdDemo) {
      window.history.back();
    } else {
      setOpen(false);
    }
    window.setTimeout(() => posterRef.current?.focus({ preventScroll: true }), 0);
  };
  useEffect(() => {
    const el = desktopRef.current;
    if (!el || phone) {
      return;
    }
    const watcher = new IntersectionObserver(
      ([entry]) => {
        onScreen.current = entry.isIntersecting;
      },
      { threshold: 0.25 },
    );
    watcher.observe(el);
    const observer = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    observer.observe(el);
    return () => {
      observer.disconnect();
      watcher.disconnect();
    };
  }, [phone]);

  // Each pass of the autopilot ends by resetting the workbook, which starts the next pass.
  const laidOut = phone ? open : wins !== null;
  useEffect(() => {
    const desktop = phone ? stageRef.current : desktopRef.current;
    const cursor = cursorRef.current;
    const api = apiRef.current;
    if (!auto || !laidOut || !apiReady || !desktop || !cursor || !api) {
      return;
    }
    let stopped = false;
    const waitUntilVisible = async () => {
      while (!phone && !onScreen.current && !stopped) {
        await new Promise((resolve) => window.setTimeout(resolve, 300));
      }
    };
    void runAutopilot({
      plan: mod.autopilot,
      desktop,
      cursor,
      api,
      cancelled: () => stopped,
      visible: waitUntilVisible,
    }).then(() => {
      if (!stopped) {
        cursor.classList.remove("is-on");
        window.setTimeout(() => !stopped && setResetKey((k) => k + 1), 600);
      }
    });
    return () => {
      stopped = true;
      cursor.classList.remove("is-on", "is-down");
      api.cancelAll();
    };
  }, [auto, laidOut, apiReady, resetKey, phone, mod]);

  const takeControl = () => {
    setAuto(false);
    setNote(null);
    update("notes", { open: true, min: false });
    apiRef.current?.cancelAll();
  };
  const watchAgain = () => {
    setResetKey((k) => k + 1);
    setAuto(true);
  };

  // Lay the windows out once the desktop has a size, and again when the platform or desktop changes.
  useEffect(() => {
    if (size) {
      setWins(initialWins(platform, size.w, size.h));
    }
  }, [platform, size?.w, size?.h]);

  const showUnsupported = (kind: UnsupportedTool, task: { name: string }) => {
    const message = mod.copy.unsupported[kind];
    setNote({ title: message.title, text: `${message.text} (${task.name})` });
    update("notes", { open: true, min: false });
    focus("notes");
  };
  const closeNote = () => {
    update("notes", { open: false });
    setNote(null);
  };

  const focus = (id: WinId) => setOrder((o) => [...o.filter((x) => x !== id), id]);
  const update = (id: WinId, patch: Partial<Win>) =>
    setWins((w) => (w ? { ...w, [id]: { ...w[id], ...patch } } : w));

  const compact = () => window.matchMedia("(max-width: 759px)").matches;

  const startMove = (id: WinId, event: TargetedPointerEvent<HTMLElement>) => {
    focus(id);
    if (event.button !== 0 || !wins || !size || wins[id].max || compact()) {
      return;
    }
    event.preventDefault();
    const start = {
      px: event.clientX,
      py: event.clientY,
      x: wins[id].x,
      y: wins[id].y,
    };
    const { top, bottom } = insets(platform);
    const w = wins[id].w;
    const onMove = (ev: PointerEvent) => {
      update(id, {
        x: Math.max(80 - w, Math.min(size.w - 80, start.x + ev.clientX - start.px)),
        y: Math.max(top, Math.min(size.h - bottom - 32, start.y + ev.clientY - start.py)),
      });
    };
    const unlock = lockSelection();
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      unlock();
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  };

  const startResize = (id: WinId, event: TargetedPointerEvent<HTMLElement>) => {
    if (event.button !== 0 || !wins || !size || wins[id].max || compact()) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    focus(id);
    const start = {
      px: event.clientX,
      py: event.clientY,
      w: wins[id].w,
      h: wins[id].h,
    };
    const { x, y } = wins[id];
    const { bottom } = insets(platform);
    const onMove = (ev: PointerEvent) => {
      update(id, {
        w: Math.max(MIN_SIZE[id].w, Math.min(size.w - x, start.w + ev.clientX - start.px)),
        h: Math.max(MIN_SIZE[id].h, Math.min(size.h - bottom - y, start.h + ev.clientY - start.py)),
      });
    };
    const unlock = lockSelection();
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      unlock();
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  };

  const showApp = (id: WinId) => {
    if (!wins) {
      return;
    }
    const win = wins[id];
    const onTop = order[order.length - 1] === id;
    if (id === "notes" && win.open && !win.min && onTop) {
      update(id, { open: false });
      return;
    }
    if (id === "excel" && win.open && !win.min && onTop) {
      update(id, { min: true });
      return;
    }
    update(id, { open: true, min: false });
    focus(id);
  };

  const { top, bottom } = insets(platform);
  const geometry = (id: WinId): CSSProperties => {
    if (!wins || !size) {
      return {};
    }
    const win = wins[id];
    const zIndex = (id === "notes" && pinned ? 50 : 10) + order.indexOf(id);
    if (win.max) {
      return {
        left: 0,
        top,
        width: size.w,
        height: size.h - top - bottom,
        zIndex,
      };
    }
    return { left: win.x, top: win.y, width: win.w, height: win.h, zIndex };
  };
  const winClass = (id: WinId) => {
    const win = wins?.[id];
    const focused = order[order.length - 1] === id;
    return [
      "xd-window",
      `xd-window-${id}`,
      win?.min ? "is-min" : "",
      (id === "notes" && auto) || (win != null && !win.open) ? "is-closed" : "",
      win?.max ? "is-max" : "",
      focused ? "" : "is-inactive",
    ]
      .filter(Boolean)
      .join(" ");
  };

  const now = clock ? formatClock(clock, platform) : null;
  const excelOpen = wins ? wins.excel.open && !wins.excel.min : true;

  return (
    <div class={`xd${fitScreen ? " is-fit" : ""}`} data-platform={platform}>
      <div class="xd-head">
        <div class="xd-head-slot">{head}</div>
        <div class="xd-toolbar">
          <div class="xd-platform" role="radiogroup" aria-label="Desktop style">
            {(["mac", "win"] as const).map((p) => (
              <button
                key={p}
                type="button"
                role="radio"
                aria-checked={platform === p}
                class={platform === p ? "is-on" : ""}
                onClick={() => setPlatform(p)}
              >
                {p === "mac" ? "macOS" : "Windows"}
              </button>
            ))}
          </div>
          <button type="button" class="xd-reset" onClick={() => setResetKey((k) => k + 1)}>
            Start over
          </button>
          {auto ? (
            <span class="xd-auto-note">Auto demo running. Hover the window to take control.</span>
          ) : (
            <button type="button" class="xd-reset xd-watch" onClick={watchAgain}>
              Watch it run
            </button>
          )}
        </div>
      </div>
      <div class="xd-compact">
        {note ? (
          <div class="xd-compact-note">
            <strong>{note.title}</strong> {note.text}{" "}
            <button type="button" class="xd-note-back" onClick={() => setNote(null)}>
              Back to the steps
            </button>
          </div>
        ) : (
          <ol class="xd-compact-steps">
            {mod.note.steps.map((step) => (
              <li key={step}>{step}</li>
            ))}
          </ol>
        )}
      </div>
      {!phone && (
        <div class={`xd-desktop xd-desktop-${platform}${wins ? " is-ready" : ""}`} ref={desktopRef}>
          {platform === "mac" ? (
            <div class="xd-menubar" aria-hidden="true">
              <span class="xd-menubar-app">Excel</span>
              {["File", "Edit", "View", "Insert", "Format", "Tools", "Data", "Window", "Help"].map((item) => (
                <span key={item}>{item}</span>
              ))}
              <span class="xd-menubar-spacer" />
              <span class="xd-menubar-icon">
                <svg viewBox="0 0 20 14" width="17" height="12">
                  <path
                    d="M10 13l2.2-2.6a3.2 3.2 0 00-4.4 0zM5.3 7.8a7 7 0 019.4 0l1.6-1.8a9.4 9.4 0 00-12.6 0zM1.8 4.4a12 12 0 0116.4 0L20 2.6a14.6 14.6 0 00-20 0z"
                    fill="currentColor"
                  />
                </svg>
              </span>
              <span class="xd-menubar-icon">
                <svg viewBox="0 0 26 12" width="24" height="11">
                  <rect
                    x=".5"
                    y=".5"
                    width="22"
                    height="11"
                    rx="3"
                    fill="none"
                    stroke="currentColor"
                    opacity=".5"
                  />
                  <rect x="2" y="2" width="16" height="8" rx="1.6" fill="currentColor" />
                  <path d="M24 4v4" stroke="currentColor" opacity=".5" stroke-width="1.5" />
                </svg>
              </span>
              {now && (
                <span>
                  {now.date}&nbsp;&nbsp;{now.time}
                </span>
              )}
            </div>
          ) : null}

          <section
            class={winClass("excel")}
            inert={auto}
            style={geometry("excel")}
            aria-label="Microsoft Excel window"
            onPointerDown={() => focus("excel")}
          >
            <Workbook
              platform={platform}
              resetKey={resetKey}
              maximized={wins?.excel.max ?? false}
              onTitlePointerDown={(event) => startMove("excel", event)}
              onMinimize={() => update("excel", { min: true })}
              onMaximize={() => wins && update("excel", { max: !wins.excel.max })}
              onClose={() => update("excel", { min: true })}
              onUnsupported={showUnsupported}
              onApi={(api) => {
                apiRef.current = api;
                setApiReady(api !== null);
              }}
            />
            <span
              class="xd-resize"
              onPointerDown={(event) => startResize("excel", event)}
              aria-hidden="true"
            />
          </section>

          <section
            class={winClass("notes")}
            inert={auto}
            style={geometry("notes")}
            aria-label="Sticky note"
            onPointerDown={() => focus("notes")}
          >
            <div class="xd-note-head" onPointerDown={(event) => startMove("notes", event)}>
              {platform === "mac" ? (
                <>
                  <button
                    type="button"
                    class="xd-note-close"
                    title="Close"
                    onPointerDown={(event) => event.stopPropagation()}
                    onClick={closeNote}
                  />
                  <span class="xd-note-spacer" />
                  <button
                    type="button"
                    class={`xd-note-pin${pinned ? " is-on" : ""}`}
                    title={pinned ? "Unpin: let other windows cover this note" : "Pin: keep this note on top"}
                    aria-pressed={pinned}
                    onPointerDown={(event) => event.stopPropagation()}
                    onClick={() => setPinned(!pinned)}
                  >
                    <svg viewBox="0 0 16 16" aria-hidden="true">
                      <path
                        d="M10 1.5l4.5 4.5-1.6.6-2.4 2.4.2 3-1 1-2.7-2.7-4 4-.7-.7 4-4L3.6 6.9l1-1 3 .2 2.4-2.4z"
                        fill="currentColor"
                      />
                    </svg>
                  </button>
                </>
              ) : (
                <>
                  <span class="xd-note-spacer" />
                  <button
                    type="button"
                    class={`xd-note-pin${pinned ? " is-on" : ""}`}
                    title={pinned ? "Unpin: let other windows cover this note" : "Pin: keep this note on top"}
                    aria-pressed={pinned}
                    onPointerDown={(event) => event.stopPropagation()}
                    onClick={() => setPinned(!pinned)}
                  >
                    <svg viewBox="0 0 16 16" aria-hidden="true">
                      <path
                        d="M10 1.5l4.5 4.5-1.6.6-2.4 2.4.2 3-1 1-2.7-2.7-4 4-.7-.7 4-4L3.6 6.9l1-1 3 .2 2.4-2.4z"
                        fill="currentColor"
                      />
                    </svg>
                  </button>

                  <button
                    type="button"
                    class="xd-note-x"
                    title="Close"
                    onPointerDown={(event) => event.stopPropagation()}
                    onClick={closeNote}
                  >
                    ×
                  </button>
                </>
              )}
            </div>
            <div class="xd-note-body">
              {note ? (
                <>
                  <strong>{note.title}</strong>
                  <p>{note.text}</p>
                  <button type="button" class="xd-note-back" onClick={() => setNote(null)}>
                    Back to the steps
                  </button>
                </>
              ) : (
                <>
                  <div class="xd-note-title">
                    <strong>{mod.note.title}</strong>
                    <button type="button" class="xd-note-btn" onClick={() => setResetKey((k) => k + 1)}>
                      Start over
                    </button>
                  </div>
                  <ol>
                    {mod.note.steps.map((step) => (
                      <li key={step}>{step}</li>
                    ))}
                  </ol>
                </>
              )}
            </div>
            <span
              class="xd-resize"
              onPointerDown={(event) => startResize("notes", event)}
              aria-hidden="true"
            />
          </section>

          {platform === "mac" ? (
            <div class="xd-dock" inert={auto}>
              <DockItem
                label="Microsoft Excel"
                running={wins?.excel.open ?? true}
                onClick={() => showApp("excel")}
              >
                <span class="xd-tile xd-tile-excel">
                  <ExcelGlyph size={34} />
                </span>
              </DockItem>
              <DockItem
                label="Stickies"
                running={(wins?.notes.open ?? false) && !auto}
                onClick={() => showApp("notes")}
              >
                <span class="xd-tile xd-tile-notes" />
              </DockItem>
              <span class="xd-dock-sep" />
              <DockItem label="Trash" running={false} onClick={() => undefined}>
                <span class="xd-tile xd-tile-trash" />
              </DockItem>
            </div>
          ) : (
            <div class="xd-taskbar" inert={auto}>
              <span class="xd-taskbar-spacer" />
              <span class="xd-start" title="Start" aria-hidden="true">
                <svg viewBox="0 0 20 20" width="20" height="20">
                  <rect x="1" y="1" width="8.5" height="8.5" rx="1" fill="#3b8ee8" />
                  <rect x="10.5" y="1" width="8.5" height="8.5" rx="1" fill="#3b8ee8" />
                  <rect x="1" y="10.5" width="8.5" height="8.5" rx="1" fill="#3b8ee8" />
                  <rect x="10.5" y="10.5" width="8.5" height="8.5" rx="1" fill="#3b8ee8" />
                </svg>
              </span>
              <span class="xd-tb-search" aria-hidden="true">
                <svg viewBox="0 0 20 20" width="14" height="14">
                  <circle cx="8" cy="8" r="5.5" fill="none" stroke="currentColor" stroke-width="1.6" />
                  <path d="M12.5 12.5l5 5" stroke="currentColor" stroke-width="1.6" />
                </svg>
                Search
              </span>
              <button
                type="button"
                class={`xd-tb-app${wins?.excel.open ? " is-running" : ""}${excelOpen && order[order.length - 1] === "excel" ? " is-active" : ""}`}
                title="Excel"
                onClick={() => showApp("excel")}
              >
                <ExcelGlyph size={24} />
              </button>
              <button
                type="button"
                class={`xd-tb-app${wins?.notes.open && !auto ? " is-running" : ""}`}
                title="Sticky Notes"
                onClick={() => showApp("notes")}
              >
                <span class="xd-tile-notes xd-tile-notes-sm" />
              </button>
              <span class="xd-taskbar-spacer" />
              <span class="xd-tray" aria-hidden="true">
                <svg viewBox="0 0 20 14" width="15" height="11">
                  <path
                    d="M10 13l2.2-2.6a3.2 3.2 0 00-4.4 0zM5.3 7.8a7 7 0 019.4 0l1.6-1.8a9.4 9.4 0 00-12.6 0zM1.8 4.4a12 12 0 0116.4 0L20 2.6a14.6 14.6 0 00-20 0z"
                    fill="currentColor"
                  />
                </svg>
                <svg viewBox="0 0 20 20" width="14" height="14">
                  <path d="M3 8h3l4-4v12l-4-4H3z" fill="currentColor" />
                  <path
                    d="M13 7a4 4 0 010 6M15 5a7 7 0 010 10"
                    fill="none"
                    stroke="currentColor"
                    stroke-width="1.3"
                  />
                </svg>
                {now && (
                  <span class="xd-tray-clock">
                    <span>{now.time}</span>
                    <span>{now.date}</span>
                  </span>
                )}
              </span>
            </div>
          )}

          <div class={`xd-cursor${platform === "mac" ? " is-mac" : ""}`} ref={cursorRef} aria-hidden="true">
            <svg viewBox="0 0 24 24" width="24" height="24">
              <path
                d="M4 2.5v17l4.6-4.2 3 6.4 3.1-1.4-3-6.3h6.3z"
                fill="#fff"
                stroke="#111"
                stroke-width="1.4"
                stroke-linejoin="round"
              />
            </svg>
            <span class="xd-cursor-ring" />
          </div>
          {auto && (
            <div class="xd-autoveil">
              <button type="button" class="xd-take" onClick={takeControl}>
                <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
                  <path
                    d="M4 2.5v17l4.6-4.2 3 6.4 3.1-1.4-3-6.3h6.3z"
                    fill="currentColor"
                    stroke="currentColor"
                    stroke-width="1.2"
                    stroke-linejoin="round"
                  />
                </svg>
                Take control
              </button>
            </div>
          )}
        </div>
      )}

      <Poster module={mod} buttonRef={posterRef} onOpen={openDemo} />
      {phone && open && (
        <div class="xd-phone" role="dialog" aria-modal="true" aria-label="Excel Educator demo">
          <div class="xd-phone-bar">
            <span>Try the add-in</span>
            <button type="button" class="xd-phone-close" onClick={closeDemo}>
              Close
            </button>
          </div>
          {!auto && <p class="xd-phone-hint">Tap a cell, then tap the formula bar to type.</p>}
          <div class="xd-phone-body" ref={stageRef}>
            <div class="xd-phone-excel" inert={auto}>
              <Workbook
                platform={platform}
                resetKey={resetKey}
                maximized={false}
                onTitlePointerDown={() => undefined}
                onMinimize={() => undefined}
                onMaximize={() => undefined}
                onClose={closeDemo}
                onUnsupported={showUnsupported}
                onApi={(api) => {
                  apiRef.current = api;
                  setApiReady(api !== null);
                }}
              />
            </div>
            <div class="xd-cursor" ref={cursorRef} aria-hidden="true">
              <svg viewBox="0 0 24 24" width="24" height="24">
                <path
                  d="M4 2.5v17l4.6-4.2 3 6.4 3.1-1.4-3-6.3h6.3z"
                  fill="#fff"
                  stroke="#111"
                  stroke-width="1.4"
                  stroke-linejoin="round"
                />
              </svg>
              <span class="xd-cursor-ring" />
            </div>
            {auto && (
              <div class="xd-autoveil" onClick={takeControl}>
                <button type="button" class="xd-take">
                  Tap to take control
                </button>
              </div>
            )}
            {note && (
              <div class="xd-phone-toast" role="status">
                <strong>{note.title}</strong> {note.text}{" "}
                <button type="button" class="xd-note-back" onClick={() => setNote(null)}>
                  Dismiss
                </button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/** What a phone sees in the page: a still of the finished sheet that opens the simulator full screen. */
function Poster({
  module: mod,
  onOpen,
  buttonRef,
}: {
  module: SimModule;
  onOpen: () => void;
  buttonRef: Ref<HTMLButtonElement>;
}) {
  const { poster } = mod;
  const { styles, widths } = mod.sheet;
  const values = useMemo(() => evaluateSheet(solvedInputs(mod)), [mod]);
  const columns = widths.slice(poster.c1, poster.c2 + 1);
  const rows = [];
  for (let row = poster.r1; row <= poster.r2; row++) {
    const cells = rangeAddrs({ c1: poster.c1, r1: row, c2: poster.c2, r2: row }).map((addr) => {
      const key = addrName(addr);
      const style = styles[key];
      const value = values[key] ?? null;
      const align =
        style?.align === "center" || isError(value) ? "center" : typeof value === "number" ? "right" : "left";
      return (
        <span
          key={key}
          style={{
            background: style?.fill,
            color: style?.color,
            fontWeight: style?.bold ? 700 : undefined,
            textAlign: align,
          }}
        >
          {formatValue(value, style?.format ?? "General", widths[addr.col])}
        </span>
      );
    });
    rows.push(
      <span class="xd-poster-row" key={row} style={{ gridTemplateColumns: columns.map((w) => `${w}fr`).join(" ") }}>
        {cells}
      </span>,
    );
  }
  const first = mod.tasks[0];
  return (
    <button type="button" class="xd-poster" ref={buttonRef} onClick={onOpen}>
      <span class="xd-poster-sheet" aria-hidden="true">
        {rows}
      </span>
      <span class="xd-poster-pane" aria-hidden="true">
        <b>{mod.copy.paneTitle}</b>
        {first && <span>1.1 {first.text}</span>}
      </span>
      <span class="xd-poster-cta">
        <svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true">
          <path d="M5 3.5v13l11-6.5z" fill="currentColor" />
        </svg>
        Try the add-in
      </span>
      <span class="xd-poster-note">Opens full screen</span>
    </button>
  );
}

function DockItem({
  label,
  running,
  onClick,
  children,
}: {
  label: string;
  running: boolean;
  onClick: () => void;
  children: JSX.Element;
}) {
  return (
    <button type="button" class="xd-dock-item" onClick={onClick} aria-label={label}>
      <span class="xd-dock-label">{label}</span>
      {children}
      <span class={`xd-dock-dot${running ? " is-on" : ""}`} />
    </button>
  );
}
