import type { ComponentChildren, CSSProperties, JSX, Ref, TargetedPointerEvent } from "preact";
import { useEffect, useMemo, useRef, useState } from "preact/hooks";
import { addrName, evaluateSheet, formatValue, isError, rangeAddrs } from "../engine/sheet.ts";
import { solvedInputs } from "../module/grade.ts";
import type { NoteMessage, SimModule, UnsupportedTool } from "../module/types.ts";
import { runAutopilot, type WorkbookApi } from "./autopilot.ts";
import { CursorMark } from "./CursorMark.tsx";
import { useSim } from "./context.ts";
import { Playback, SKIP_SECONDS, elapsedAt, locateTime, sumDurations, type Hold } from "./playback.ts";
import type { Platform } from "./Ribbon.tsx";
import { TransportBar } from "./TransportBar.tsx";
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

/** Height of the playback bar along the bottom of the desktop. */
const TRANSPORT = 46;

/** Space the menu bar, the playback bar, and the dock (or the taskbar) take from the desktop. */
function insets(platform: Platform) {
  return platform === "mac" ? { top: 26, bottom: 74 + TRANSPORT } : { top: 0, bottom: 48 + TRANSPORT };
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
  /** Run the auto demo. Off for people who ask for less motion. The bar can pause it and resume from the same moment. */
  autoplay?: boolean;
  /** On a phone, let the browser's Back button close the full-screen view. */
  history?: boolean;
  /** Size the desktop to the screen, less the host's sticky header (the `--chrome-h` property). */
  fitScreen?: boolean;
  /** The host's own heading, shown above the desktop beside the toolbar. */
  head?: ComponentChildren;
};

function NarrationNote({
  lines,
  line,
  activeRef,
  needsPlay,
  onPlay,
}: {
  lines: string[];
  line: number;
  activeRef?: Ref<HTMLParagraphElement>;
  needsPlay: boolean;
  onPlay: () => void;
}) {
  return (
    <div class="xd-narration">
      <strong>Narration</strong>
      {needsPlay ? (
        <button type="button" class="xd-narration-play" onClick={onPlay}>
          Play narration
        </button>
      ) : null}
      <p class="xd-narration-live" aria-live="polite">
        {line >= 0 ? lines[line] : ""}
      </p>
      {lines.map((text, index) => (
        <p key={index} class={index === line ? "is-on" : ""} ref={index === line ? activeRef : undefined}>
          {text}
        </p>
      ))}
    </div>
  );
}

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
  const [session, setSession] = useState(true);
  const [hold, setHold] = useState<Hold>("play");
  const [rate, setRate] = useState(1);
  const [muted, setMuted] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const [total, setTotal] = useState(0);
  const [timelineReady, setTimelineReady] = useState(false);
  const [scrub, setScrub] = useState<number | null>(null);
  const [skipping, setSkipping] = useState(false);
  const [apiReady, setApiReady] = useState(false);
  const [phone, setPhone] = useState(false);
  const [open, setOpen] = useState(false);
  const [calm, setCalm] = useState(false);
  const [line, setLine] = useState(-1);
  const [needsPlay, setNeedsPlay] = useState(false);
  const desktopRef = useRef<HTMLDivElement>(null);
  const stageRef = useRef<HTMLDivElement>(null);
  const posterRef = useRef<HTMLButtonElement>(null);
  const cursorRef = useRef<HTMLDivElement>(null);
  const apiRef = useRef<WorkbookApi | null>(null);
  const audioRef = useRef<HTMLAudioElement | null>(null);
  const resumeNarration = useRef<(() => void) | null>(null);
  const activeLineRef = useRef<HTMLParagraphElement>(null);
  const onScreen = useRef(true);
  const playbackRef = useRef(new Playback());
  const mutedRef = useRef(false);
  const holdRef = useRef<Hold>("play");
  const launchRef = useRef<{ index: number; offset: number; rush?: boolean } | null>(null);
  const skipRef = useRef({ index: 0, offset: 0, rush: false });
  const scrubWait = useRef(0);
  const stopRef = useRef(() => {});
  const progressRef = useRef({ durations: [] as number[], known: [] as boolean[], index: 0, offset: 0 });

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
    const reduced = !autoplay || window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    setCalm(reduced);
    const next: Hold = reduced ? "control" : "play";
    holdRef.current = next;
    setHold(next);
    setSession(!reduced);
    playbackRef.current.setHold(next);
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
    launchRef.current = null;
    const next: Hold = calm ? "control" : "play";
    holdRef.current = next;
    setHold(next);
    setSession(!calm);
    playbackRef.current.setHold(next);
    stopRef.current();
    setResetKey((k) => k + 1);
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
    activeLineRef.current?.scrollIntoView({ block: "nearest" });
  }, [line]);
  useEffect(() => {
    if (!session) {
      return;
    }
    const tick = () => {
      const progress = progressRef.current;
      const audio = audioRef.current;
      const into =
        audio && Number(audio.dataset.step) === progress.index ? audio.currentTime : progress.offset;
      const ready = progress.known.length > 0 && progress.known.every(Boolean);
      setTimelineReady(ready);
      setTotal(ready ? sumDurations(progress.durations) : 0);
      setElapsed(elapsedAt(progress.durations, progress.index, into));
    };
    tick();
    const timer = window.setInterval(tick, 200);
    return () => window.clearInterval(timer);
  }, [session, resetKey]);
  useEffect(() => {
    const desktop = phone ? stageRef.current : desktopRef.current;
    const cursor = cursorRef.current;
    const api = apiRef.current;
    if (!session || !laidOut || !apiReady || !desktop || !cursor || !api) {
      return;
    }
    let stopped = false;
    let release: (() => void) | null = null;
    const playback = playbackRef.current;
    const launch = launchRef.current ?? { index: 0, offset: 0, rush: false };
    launchRef.current = null;
    skipRef.current = { index: launch.index, offset: launch.offset, rush: launch.rush ?? false };
    stopRef.current = () => {
      stopped = true;
      playback.wake();
    };
    const narrated = mod.autopilot.some((step) => step.audio !== undefined);
    const clips: HTMLAudioElement[] = [];
    let current: HTMLAudioElement | null = null;
    let queued: { index: number; audio: HTMLAudioElement } | null = null;
    const durations = mod.autopilot.map(() => 0);
    const known = mod.autopilot.map((step) => step.audio === undefined);
    progressRef.current = { durations, known, index: launch.index, offset: launch.offset };
    setSkipping(launch.index > 0 || launch.offset > 0);
    if (narrated) {
      setLine(-1);
      setWins((wins) => (wins ? { ...wins, notes: { ...wins.notes, open: true, min: false } } : wins));
    }
    for (let index = 0; index < mod.autopilot.length; index++) {
      const url = mod.autopilot[index]?.audio;
      if (!url) {
        continue;
      }
      const probe = new Audio();
      clips.push(probe);
      probe.preload = "metadata";
      const finish = (seconds: number) => {
        if (stopped) {
          return;
        }
        durations[index] = seconds;
        known[index] = true;
      };
      probe.addEventListener("loadedmetadata", () => finish(Number.isFinite(probe.duration) ? probe.duration : 0), {
        once: true,
      });
      probe.addEventListener("error", () => finish(0), { once: true });
      probe.src = url;
    }
    const waitUntilVisible = async () => {
      while (!phone && !onScreen.current && !stopped) {
        await new Promise((resolve) => window.setTimeout(resolve, 300));
      }
    };
    const prepare = (index: number) => {
      const url = mod.autopilot[index]?.audio;
      if (!url) {
        return null;
      }
      if (queued?.index === index) {
        const audio = queued.audio;
        queued = null;
        return audio;
      }
      const audio = new Audio(url);
      audio.preload = "auto";
      clips.push(audio);
      return audio;
    };
    const clock = narrated
      ? {
          play: async (index: number) => {
            await playback.gate(() => stopped);
            if (stopped) {
              return 0;
            }
            current?.pause();
            const audio = prepare(index);
            if (!audio) {
              return 0;
            }
            current = audio;
            audio.dataset.step = String(index);
            audioRef.current = audio;
            const next = prepare(index + 1);
            if (next) {
              queued = { index: index + 1, audio: next };
            }
            const metadata = new Promise<void>((resolve) => {
              if (audio.readyState >= 1) {
                resolve();
                return;
              }
              audio.addEventListener("loadedmetadata", () => resolve(), { once: true });
              audio.addEventListener("error", () => resolve(), { once: true });
            });
            await metadata;
            if (stopped) {
              return 0;
            }
            const duration = Number.isFinite(audio.duration) ? audio.duration : 0;
            if (durations[index] === 0 && duration > 0) {
              durations[index] = duration;
              known[index] = true;
            }
            const live = skipRef.current;
            const offset = index === live.index ? Math.min(live.offset, Math.max(0, duration - 0.05)) : 0;
            if (offset > 0) {
              audio.currentTime = offset;
            }
            audio.muted = mutedRef.current;
            audio.playbackRate = playback.rate;
            progressRef.current.offset = offset;
            if (!stopped) {
              setSkipping(false);
            }
            const startClip = async () => {
              await playback.gate(() => stopped);
              if (stopped || playback.hold !== "play") {
                return false;
              }
              try {
                await audio.play();
                return true;
              } catch {
                return false;
              }
            };
            if (await startClip()) {
              if (!stopped) {
                setNeedsPlay(false);
              }
            } else if (!stopped && playback.hold === "play") {
              setNeedsPlay(true);
              await new Promise<void>((resolve) => {
                release = resolve;
                resumeNarration.current = resolve;
              });
              if (stopped) {
                return 0;
              }
              await startClip();
              if (!stopped) {
                setNeedsPlay(false);
              }
            }
            if (playback.hold !== "play") {
              audio.pause();
            }
            return duration;
          },
          untilEnded: async () => {
            const audio = current;
            if (!audio) {
              return;
            }
            while (!stopped && !audio.ended && audio.error === null) {
              await playback.gate(() => stopped);
              if (stopped) {
                return;
              }
              if (audio.paused) {
                audio.muted = mutedRef.current;
                audio.playbackRate = playback.rate;
                try {
                  await audio.play();
                } catch {
                  if (stopped) {
                    return;
                  }
                  await new Promise((resolve) => window.setTimeout(resolve, 250));
                  continue;
                }
              }
              await new Promise((resolve) => window.setTimeout(resolve, 50));
              await waitUntilVisible();
            }
          },
          onLine: (index: number) => {
            if (!stopped) {
              setLine(index);
            }
          },
        }
      : undefined;
    const run = async () => {
      await runAutopilot({
        plan: mod.autopilot,
        desktop,
        cursor,
        api,
        cancelled: () => stopped,
        visible: waitUntilVisible,
        clock,
        skipTo: () => skipRef.current,
        playback,
        onStep: (index) => {
          progressRef.current.index = index;
          const live = skipRef.current;
          progressRef.current.offset = index === live.index ? live.offset : 0;
          if (!stopped && index >= live.index) {
            setSkipping(false);
          }
        },
      });
      if (!stopped) {
        cursor.classList.remove("is-on");
        window.setTimeout(() => !stopped && setResetKey((k) => k + 1), 600);
      }
    };
    void run();
    return () => {
      stopped = true;
      playback.wake();
      release?.();
      resumeNarration.current = null;
      setNeedsPlay(false);
      setSkipping(false);
      for (const clip of clips) {
        clip.pause();
        clip.src = "";
      }
      audioRef.current = null;
      cursor.classList.remove("is-on", "is-down");
      api.cancelAll();
    };
  }, [session, laidOut, apiReady, resetKey, phone, mod]);

  const publishHold = (next: Hold) => {
    holdRef.current = next;
    setHold(next);
    playbackRef.current.setHold(next);
  };
  const here = () => {
    const progress = progressRef.current;
    const audio = audioRef.current;
    const into = audio && Number(audio.dataset.step) === progress.index ? audio.currentTime : progress.offset;
    return { ...progress, into, at: elapsedAt(progress.durations, progress.index, into) };
  };
  const relaunch = (at: { index: number; offset: number; rush?: boolean }, keepScrub = false) => {
    launchRef.current = at;
    skipRef.current = { index: at.index, offset: at.offset, rush: at.rush ?? false };
    stopRef.current();
    publishHold("play");
    setSession(true);
    if (!keepScrub) {
      setScrub(null);
    }
    setResetKey((k) => k + 1);
  };
  const seekTo = (seconds: number, keepScrub = false) => {
    const progress = here();
    if (progress.known.length === 0 || !progress.known.every(Boolean)) {
      return;
    }
    const loc = locateTime(progress.durations, seconds);
    const ahead =
      loc.index > progress.index || (loc.index === progress.index && loc.offset >= progress.into - 0.05);
    const at = {
      index: loc.index,
      offset: loc.offset,
      rush: loc.index === progress.index && loc.offset > progress.into + 0.35,
    };
    if (ahead && session) {
      skipRef.current = at;
      const audio = audioRef.current;
      if (loc.index === progress.index && audio && Number(audio.dataset.step) === progress.index) {
        const span = Number.isFinite(audio.duration) ? audio.duration : loc.offset;
        audio.currentTime = Math.min(loc.offset, span);
        progressRef.current.offset = audio.currentTime;
        setElapsed(elapsedAt(progress.durations, progress.index, audio.currentTime));
      }
      if (!keepScrub) {
        setScrub(null);
      }
      return;
    }
    relaunch(at, keepScrub);
  };
  const scrubTo = (seconds: number) => {
    setScrub(seconds);
    window.clearTimeout(scrubWait.current);
    scrubWait.current = window.setTimeout(() => seekTo(seconds, true), 80);
  };
  const play = () => {
    if (!session) {
      launchRef.current = null;
      publishHold("play");
      setSession(true);
      setResetKey((k) => k + 1);
      return;
    }
    publishHold("play");
    const audio = audioRef.current;
    if (audio && !audio.ended) {
      audio.muted = mutedRef.current;
      audio.playbackRate = playbackRef.current.rate;
      void audio.play().then(
        () => {
          setNeedsPlay(false);
          resumeNarration.current?.();
        },
        () => setNeedsPlay(true),
      );
      return;
    }
    resumeNarration.current?.();
  };
  const freezeCursor = () => {
    const cursor = cursorRef.current;
    if (!cursor) {
      return;
    }
    const computed = getComputedStyle(cursor).transform;
    cursor.style.transitionDuration = "0ms";
    if (computed && computed !== "none") {
      cursor.style.transform = computed;
    }
    cursor.classList.remove("is-down");
  };
  const pause = () => {
    if (holdRef.current !== "play") {
      return;
    }
    publishHold("pause");
    audioRef.current?.pause();
    freezeCursor();
  };
  const startNarration = () => {
    play();
  };
  const takeControl = () => {
    if (!session || holdRef.current === "control") {
      return;
    }
    publishHold("control");
    audioRef.current?.pause();
    apiRef.current?.previewFill("", null);
    freezeCursor();
    setNote(null);
    setNeedsPlay(false);
    update("notes", { open: true, min: false });
  };
  const restart = () => {
    launchRef.current = null;
    stopRef.current();
    publishHold("play");
    setSession(true);
    setScrub(null);
    setResetKey((k) => k + 1);
  };
  const toggleMute = () => {
    const next = !mutedRef.current;
    mutedRef.current = next;
    setMuted(next);
    const audio = audioRef.current;
    if (audio) {
      audio.muted = next;
    }
  };
  const chooseRate = (next: number) => {
    playbackRef.current.setRate(next);
    setRate(next);
    const audio = audioRef.current;
    if (audio) {
      audio.playbackRate = next;
    }
  };
  const takeOverFromPointer = (event: { target: EventTarget | null }) => {
    if (holdRef.current === "control" || !session) {
      return;
    }
    const target = event.target;
    if (!(target instanceof Element) || target.closest(".xd-transport")) {
      return;
    }
    takeControl();
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
    if (session && hold !== "control" && mod.autopilot.some((step) => step.audio !== undefined)) {
      return;
    }
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
      if (session && hold !== "control" && mod.autopilot.some((step) => step.audio !== undefined)) {
        return;
      }
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
  const narrating = session && mod.autopilot.some((step) => step.audio !== undefined);
  const handsOff = session && hold !== "control";
  const narrationLines = mod.autopilot.map((step) => step.transcript ?? "");
  const geometry = (id: WinId): CSSProperties => {
    if (!wins || !size) {
      return {};
    }
    const win = wins[id];
    const zIndex = (id === "notes" && pinned ? 50 : 10) + order.indexOf(id) + (id === "notes" && narrating ? 100 : 0);
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
    const shut = id === "notes" && narrating ? false : (id === "notes" && handsOff) || (win != null && !win.open);
    return [
      "xd-window",
      `xd-window-${id}`,
      win?.min ? "is-min" : "",
      shut ? "is-closed" : "",
      win?.max ? "is-max" : "",
      focused ? "" : "is-inactive",
    ]
      .filter(Boolean)
      .join(" ");
  };

  const now = clock ? formatClock(clock, platform) : null;
  const excelOpen = wins ? wins.excel.open && !wins.excel.min : true;
  const transport = (
    <TransportBar
      hold={hold}
      session={session}
      rate={rate}
      elapsed={elapsed}
      total={total}
      ready={timelineReady}
      scrub={scrub}
      onScrub={(seconds) => {
        if (seconds === null) {
          window.clearTimeout(scrubWait.current);
          setScrub(null);
          return;
        }
        scrubTo(seconds);
      }}
      onSeek={(seconds) => {
        window.clearTimeout(scrubWait.current);
        setScrub(null);
        seekTo(seconds);
      }}
      onPlay={play}
      onPause={pause}
      onRestart={restart}
      onRewind={() => seekTo(here().at - SKIP_SECONDS)}
      onFastForward={() => seekTo(here().at + SKIP_SECONDS)}
      onTakeControl={takeControl}
      muted={muted}
      onMute={toggleMute}
      onRate={chooseRate}
    />
  );

  return (
    <div
      class={`xd${fitScreen ? " is-fit" : ""}${narrating ? " is-narrating" : ""}${hold === "control" ? " is-user" : ""}${skipping ? " is-skipping" : ""}`}
      data-platform={platform}
    >
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
          <button type="button" class="xd-reset" onClick={() => (session ? restart() : setResetKey((k) => k + 1))}>
            Start over
          </button>
        </div>
      </div>
      <div class="xd-compact">
        {narrating ? (
          <NarrationNote
            lines={narrationLines}
            line={line}
            activeRef={phone ? activeLineRef : undefined}
            needsPlay={needsPlay}
            onPlay={startNarration}
          />
        ) : note ? (
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
        <div
          class={`xd-desktop xd-desktop-${platform}${wins ? " is-ready" : ""}${needsPlay ? " is-awaiting-play" : ""}`}
          ref={desktopRef}
        >
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
            style={geometry("excel")}
            aria-label="Microsoft Excel window"
            onPointerDownCapture={takeOverFromPointer}
            onKeyDownCapture={takeOverFromPointer}
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
            inert={handsOff && !narrating}
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
              {narrating ? (
                <NarrationNote
                  lines={narrationLines}
                  line={line}
                  activeRef={phone ? undefined : activeLineRef}
                  needsPlay={needsPlay}
                  onPlay={startNarration}
                />
              ) : note ? (
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
                    <button
                      type="button"
                      class="xd-note-btn"
                      onClick={() => (session ? restart() : setResetKey((k) => k + 1))}
                    >
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
            <div class="xd-dock" inert={handsOff}>
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
                running={(wins?.notes.open ?? false) && (!handsOff || narrating)}
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
            <div class="xd-taskbar" inert={handsOff}>
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
                class={`xd-tb-app${wins?.notes.open && (!handsOff || narrating) ? " is-running" : ""}`}
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
            <CursorMark size={24} />
            <span class="xd-cursor-ring" />
          </div>
          {transport}
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
          {hold === "control" && <p class="xd-phone-hint">Tap a cell, then tap the formula bar to type.</p>}
          {transport}
          <div class="xd-phone-body" ref={stageRef}>
            <div class="xd-phone-excel" onPointerDownCapture={takeOverFromPointer} onKeyDownCapture={takeOverFromPointer}>
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
              <CursorMark size={24} />
              <span class="xd-cursor-ring" />
            </div>
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
