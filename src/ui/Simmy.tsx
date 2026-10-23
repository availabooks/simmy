import type { ComponentChildren, CSSProperties, JSX, Ref, TargetedPointerEvent } from "preact";
import { useEffect, useMemo, useRef, useState } from "preact/hooks";
import { addrName, evaluateSheet, formatValue, isError, rangeAddrs } from "../engine/sheet.ts";
import { solvedState } from "../module/grade.ts";
import type { NoteMessage, SimModule, UnsupportedTool } from "../module/types.ts";
import { runAutopilot, type WorkbookApi } from "./autopilot.ts";
import { CursorMark } from "./CursorMark.tsx";
import { useSim } from "./context.ts";
import { Playback, elapsedAt, planSeek, playhead, skipTrack, sumDurations, type Hold } from "./playback.ts";
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

/** The switcher's names for each platform, in the order it lists them. */
const PLATFORM_LABELS: Record<Platform, string> = {
  mac: "macOS",
  win: "Windows",
  web: "Web",
};

/** macOS and Windows put Excel in a window on a desktop. The web view has only Excel. */
const isDesktop = (platform: Platform) => platform === "mac" || platform === "win";

/** Space the menu bar, the playback bar, and the dock (or the taskbar) take from the desktop. */
function insets(platform: Platform) {
  if (platform === "mac") {
    return { top: 26, bottom: 60 + TRANSPORT };
  }
  return { top: 0, bottom: (platform === "win" ? 48 : 0) + TRANSPORT };
}

function initialWins(platform: Platform, width: number, height: number): Wins {
  const { top, bottom } = insets(platform);
  const excelH = height - top - bottom - 20;
  const desktop = isDesktop(platform);
  return {
    // Without a desktop, Excel fills the stage the way a browser tab or a full-screen app does.
    excel: {
      x: 20,
      y: top + 12,
      w: width - 40,
      h: excelH,
      open: true,
      min: false,
      max: !desktop,
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

export type SimmyMode = "movie" | "kiosk";

export type SimmyProps = {
  /** The desktop to show: mac, win or web. "auto" (the default) matches the visitor's computer, mac or win. */
  platform?: Platform | "auto";
  /**
   * "movie" (the default) waits behind a darkened Play / Try it yourself screen and stops at the end.
   * "kiosk" plays by itself, muted, and loops.
   */
  mode?: SimmyMode;
  /** false starts with the visitor in control (and, in movie mode, without the Play screen). */
  autoplay?: boolean;
  /** "open" (the default) starts with the full ribbon; "collapsed" starts with only its tabs. The visitor can switch. */
  ribbon?: "open" | "collapsed";
  /** On a phone, let the browser's Back button close the full-screen view. */
  history?: boolean;
  /** Fit the whole simulator, playback bar included, to the screen less the host's sticky header (`--chrome-h`), down to `--simmy-min-height`. */
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

/** Movie mode's opening screen: the desktop darkened behind a Play button, with Try it yourself beside it. */
function Intro({ onPlay, onTakeControl }: { onPlay: () => void; onTakeControl: () => void }) {
  return (
    <div class="xd-intro" role="group" aria-label="Start the demo">
      <button type="button" class="xd-intro-play" title="Play" aria-label="Play" onClick={onPlay}>
        <svg viewBox="0 0 20 20" aria-hidden="true">
          <path d="M6 3.5v13l11-6.5z" fill="currentColor" />
        </svg>
      </button>
      <button type="button" class="xd-intro-control" onClick={onTakeControl}>
        <CursorMark size={16} />
        Try it yourself
      </button>
    </div>
  );
}

/** Over the desktop while the demo plays: hovering darkens it and invites the visitor to try it. CSS shows it, so it costs nothing until then. */
function TakeControlVeil({ onTake }: { onTake: () => void }) {
  return (
    <div class="xd-veil" onClick={onTake}>
      <button
        type="button"
        class="xd-intro-control"
        onClick={(event) => {
          event.stopPropagation();
          onTake();
        }}
      >
        <CursorMark size={16} />
        Try it yourself
      </button>
    </div>
  );
}

/**
 * A narration clip. It is a <video>, not an <audio>, because Chrome lets muted video start
 * without a click but refuses muted audio, and a kiosk has to start with no one touching the
 * page. It is never attached to the page, so nothing shows.
 */
function narrationClip(url?: string): HTMLMediaElement {
  const clip = document.createElement("video");
  // Keeps iOS from taking a playing video full screen.
  clip.playsInline = true;
  if (url) {
    clip.src = url;
  }
  return clip;
}

export function Simmy({
  platform: platformProp = "auto",
  mode = "movie",
  autoplay = true,
  ribbon = "open",
  history = false,
  fitScreen = false,
  head,
}: SimmyProps) {
  const { module: mod } = useSim();
  const movie = mode === "movie";
  // Movie mode waits for Play; kiosk mode starts on its own unless told not to.
  const starts = !movie && autoplay;
  const offersIntro = movie && autoplay;
  const [platform, setPlatform] = useState<Platform>(platformProp === "auto" ? "win" : platformProp);
  const [size, setSize] = useState<{ w: number; h: number } | null>(null);
  const [wins, setWins] = useState<Wins | null>(null);
  const [order, setOrder] = useState<WinId[]>(["excel", "notes"]);
  const [resetKey, setResetKey] = useState(0);
  const [pinned, setPinned] = useState(true);
  const [note, setNote] = useState<NoteMessage | null>(null);
  const [clock, setClock] = useState<Date | null>(null);
  const [session, setSession] = useState(starts);
  const [hold, setHold] = useState<Hold>(starts ? "play" : "control");
  const [intro, setIntro] = useState(offersIntro);
  const [rate, setRate] = useState(1);
  // A kiosk plays to a room that didn't ask for sound.
  const [muted, setMuted] = useState(!movie);
  const [elapsed, setElapsed] = useState(0);
  const [total, setTotal] = useState(0);
  const [timelineReady, setTimelineReady] = useState(false);
  const [scrub, setScrub] = useState<number | null>(null);
  const [skipping, setSkipping] = useState(false);
  const [apiReady, setApiReady] = useState(false);
  const [phone, setPhone] = useState(false);
  const [open, setOpen] = useState(false);
  const [calm, setCalm] = useState(!starts);
  const [line, setLine] = useState(-1);
  const [needsPlay, setNeedsPlay] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const toolbarRef = useRef<HTMLDivElement>(null);
  const desktopRef = useRef<HTMLDivElement>(null);
  const stageRef = useRef<HTMLDivElement>(null);
  const posterRef = useRef<HTMLButtonElement>(null);
  const cursorRef = useRef<HTMLDivElement>(null);
  const apiRef = useRef<WorkbookApi | null>(null);
  const audioRef = useRef<HTMLMediaElement | null>(null);
  const resumeNarration = useRef<(() => void) | null>(null);
  const activeLineRef = useRef<HTMLParagraphElement>(null);
  const onScreen = useRef(true);
  const playbackRef = useRef(new Playback());
  const mutedRef = useRef(!movie);
  const holdRef = useRef<Hold>(starts ? "play" : "control");
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

  // Movie mode waits for Play. A kiosk starts on its own, except for people who ask for less motion.
  useEffect(() => {
    const reduced = !starts || window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    setCalm(reduced);
    const next: Hold = reduced ? "control" : "play";
    holdRef.current = next;
    setHold(next);
    setSession(!reduced);
    setIntro(offersIntro);
    playbackRef.current.setHold(next);
  }, [starts, offersIntro]);
  useEffect(() => {
    mutedRef.current = !movie;
    setMuted(!movie);
  }, [movie]);

  // With fit-screen, the kiosk's controls don't count against the screen: the desktop gets the
  // height it would have without them, and they sit above it. The space they take is measured,
  // because it depends on the fonts and on whether the host put a heading above them.
  useEffect(() => {
    const root = rootRef.current;
    const toolbar = toolbarRef.current;
    const head = toolbar?.parentElement;
    if (!fitScreen || !root || !toolbar || !head) {
      return;
    }
    const sync = () => {
      // The toolbar, and the space after it: the gap below the host's heading when there is one,
      // otherwise the heading area's margin above the desktop. A hidden toolbar takes none.
      const style = getComputedStyle(head);
      const slot = head.firstElementChild as HTMLElement | null;
      const after = parseFloat(slot && slot.offsetHeight > 0 ? style.rowGap : style.marginBottom) || 0;
      const space = toolbar.offsetHeight > 0 ? toolbar.offsetHeight + after : 0;
      root.style.setProperty("--xd-toolbar-space", `${space}px`);
    };
    sync();
    const observer = new ResizeObserver(sync);
    observer.observe(head);
    return () => {
      observer.disconnect();
      root.style.removeProperty("--xd-toolbar-space");
    };
  }, [fitScreen, movie]);

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
    setIntro(offersIntro);
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
      const audioTime =
        audio && Number(audio.dataset.step) === progress.index && Number.isFinite(audio.currentTime)
          ? audio.currentTime
          : null;
      const into = playhead(progress.offset, audioTime);
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
    const clips: HTMLMediaElement[] = [];
    let current: HTMLMediaElement | null = null;
    let queued: { index: number; audio: HTMLMediaElement } | null = null;
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
      const probe = narrationClip();
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
      const audio = narrationClip(url);
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
            const superseded = () => skipRef.current.index > Number(audio.dataset.step);
            while (!stopped && !audio.ended && audio.error === null) {
              if (superseded()) {
                audio.pause();
                return;
              }
              await playback.gate(() => stopped);
              if (stopped || superseded()) {
                if (superseded()) {
                  audio.pause();
                }
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
          const live = skipRef.current;
          // Catch-up still plays the steps in between. The bar stays on the time the user chose.
          if (index < live.index) {
            return;
          }
          progressRef.current.index = index;
          progressRef.current.offset = index === live.index ? live.offset : 0;
          if (!stopped) {
            setSkipping(false);
          }
        },
      });
      if (stopped) {
        return;
      }
      cursor.classList.remove("is-on");
      if (movie) {
        // A movie ends on the finished sheet and hands it to the visitor. Play starts it over.
        holdRef.current = "control";
        setHold("control");
        playback.setHold("control");
        setSession(false);
      } else {
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
  }, [session, laidOut, apiReady, resetKey, phone, mod, movie]);

  const publishHold = (next: Hold) => {
    holdRef.current = next;
    setHold(next);
    playbackRef.current.setHold(next);
  };
  const here = () => {
    const progress = progressRef.current;
    const audio = audioRef.current;
    const audioTime =
      audio && Number(audio.dataset.step) === progress.index && Number.isFinite(audio.currentTime)
        ? audio.currentTime
        : null;
    const into = playhead(progress.offset, audioTime);
    return { ...progress, into, at: elapsedAt(progress.durations, progress.index, into) };
  };
  const relaunch = (at: { index: number; offset: number; rush?: boolean }, keepScrub = false) => {
    setIntro(false);
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
    const plan = planSeek({
      durations: progress.durations,
      known: progress.known,
      index: progress.index,
      into: progress.into,
      seconds,
      session,
    });
    if (plan.kind === "wait") {
      return;
    }
    const audio = audioRef.current;
    const audioStep = audio ? Number(audio.dataset.step) : Number.NaN;
    progressRef.current.index = plan.index;
    progressRef.current.offset = plan.offset;
    setElapsed(elapsedAt(progress.durations, plan.index, plan.offset));
    if (!keepScrub) {
      setScrub(null);
    }
    if (plan.kind === "relaunch") {
      relaunch({ index: plan.index, offset: plan.offset, rush: plan.rush }, keepScrub);
      return;
    }
    skipRef.current = { index: plan.index, offset: plan.offset, rush: plan.rush };
    if (audio && audioStep === progress.index) {
      if (plan.leaveClip) {
        audio.pause();
      } else {
        const span = Number.isFinite(audio.duration) ? audio.duration : plan.offset;
        audio.currentTime = Math.min(plan.offset, span);
      }
    }
    if (plan.leaveClip) {
      setSkipping(true);
    }
  };
  const scrubTo = (seconds: number) => {
    setScrub(seconds);
    window.clearTimeout(scrubWait.current);
    scrubWait.current = window.setTimeout(() => {
      const progress = here();
      const plan = planSeek({
        durations: progress.durations,
        known: progress.known,
        index: progress.index,
        into: progress.into,
        seconds,
        session,
      });
      // Dragging backward restarts the pass. Wait until the pointer is released
      // so the restart does not fight the thumb on every move.
      if (plan.kind === "jump") {
        seekTo(seconds, true);
      }
    }, 80);
  };
  const play = () => {
    setIntro(false);
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
    setIntro(false);
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
  // A movie shows no sticky note until it plays.
  const noteShown = !intro && (narrating || !handsOff);
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
    // The narration stays up while it plays. Once the visitor has control, its note closes like any other.
    const shut = id === "notes" && narrating && handsOff ? false : (id === "notes" && !noteShown) || (win != null && !win.open);
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

  // The desktop switcher and Start over. A kiosk shows them above the desktop. A movie keeps the
  // page clear and tucks them into a settings panel on the playback bar.
  const controls = (
    <>
      <div class="xd-platform" role="radiogroup" aria-label="Desktop style">
        {(Object.keys(PLATFORM_LABELS) as Platform[]).map((p) => (
          <button
            key={p}
            type="button"
            role="radio"
            aria-checked={platform === p}
            class={platform === p ? "is-on" : ""}
            onClick={() => setPlatform(p)}
          >
            {PLATFORM_LABELS[p]}
          </button>
        ))}
      </div>
      <button type="button" class="xd-reset" onClick={() => (session ? restart() : setResetKey((k) => k + 1))}>
        Start over
      </button>
    </>
  );
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
      onRewind={() => {
        const place = here();
        seekTo(skipTrack(place.durations, place.index, "back"));
      }}
      onFastForward={() => {
        const place = here();
        seekTo(skipTrack(place.durations, place.index, "forward"));
      }}
      onTakeControl={takeControl}
      muted={muted}
      onMute={toggleMute}
      onRate={chooseRate}
      settings={
        movie ? (
          <>
            <span class="xd-settings-label">Desktop</span>
            {controls}
          </>
        ) : undefined
      }
    />
  );

  return (
    <div
      ref={rootRef}
      class={`xd${fitScreen ? " is-fit" : ""}${narrating ? " is-narrating" : ""}${hold === "control" ? " is-user" : ""}${skipping ? " is-skipping" : ""}`}
      data-platform={platform}
    >
      {/* A movie with no heading of the host's has nothing up here, so it leaves no gap above the desktop. */}
      {(head || !movie) && (
        <div class="xd-head">
          <div class="xd-head-slot">{head}</div>
          {!movie && (
            <div class="xd-toolbar" ref={toolbarRef}>
              {controls}
            </div>
          )}
        </div>
      )}
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
            <div key="menubar" class="xd-menubar" aria-hidden="true">
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
            inert={intro}
            onPointerDownCapture={takeOverFromPointer}
            onKeyDownCapture={takeOverFromPointer}
            onPointerDown={() => focus("excel")}
          >
            <Workbook
              platform={platform}
              resetKey={resetKey}
              ribbonCollapsed={ribbon === "collapsed"}
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
            inert={!noteShown}
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
            <div key="dock" class="xd-dock" inert={handsOff || intro}>
              <DockItem
                label="Microsoft Excel"
                running={wins?.excel.open ?? true}
                onClick={() => showApp("excel")}
              >
                <span class="xd-tile xd-tile-excel">
                  <ExcelGlyph size={27} />
                </span>
              </DockItem>
              <DockItem
                label="Stickies"
                running={(wins?.notes.open ?? false) && noteShown}
                onClick={() => showApp("notes")}
              >
                <span class="xd-tile xd-tile-notes" />
              </DockItem>
              <span class="xd-dock-sep" />
              <DockItem label="Trash" running={false} onClick={() => undefined}>
                <span class="xd-tile xd-tile-trash" />
              </DockItem>
            </div>
          ) : platform === "win" ? (
            <div key="taskbar" class="xd-taskbar" inert={handsOff || intro}>
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
                class={`xd-tb-app${wins?.notes.open && noteShown ? " is-running" : ""}`}
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
          ) : (
            // No dock or taskbar to bring a closed note back from, so the stage offers it.
            wins &&
            !wins.notes.open &&
            !intro &&
            !handsOff && (
              <button key="note-reopen" type="button" class="xd-note-reopen" onClick={() => showApp("notes")}>
                <span class="xd-tile-notes xd-tile-notes-sm" />
                Show note
              </button>
            )
          )}

          {/*
            Keyed, like the platform's own furniture above, so switching platforms never hands the
            cursor's element (and the transform the demo left on it) to the dock or the taskbar.
          */}
          <div key="cursor" class={`xd-cursor${platform === "mac" ? " is-mac" : ""}`} ref={cursorRef} aria-hidden="true">
            <CursorMark size={24} />
            <span class="xd-cursor-ring" />
          </div>
          {intro && <Intro onPlay={play} onTakeControl={() => setIntro(false)} />}
          {session && hold === "play" && !intro && <TakeControlVeil onTake={takeControl} />}
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
          {hold === "control" && !intro && <p class="xd-phone-hint">Tap a cell, then tap the formula bar to type.</p>}
          {transport}
          <div class="xd-phone-body" ref={stageRef}>
            <div class="xd-phone-excel" inert={intro} onPointerDownCapture={takeOverFromPointer} onKeyDownCapture={takeOverFromPointer}>
              <Workbook
                platform={platform}
                resetKey={resetKey}
                ribbonCollapsed={ribbon === "collapsed"}
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
            {intro && <Intro onPlay={play} onTakeControl={() => setIntro(false)} />}
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
  const values = useMemo(() => {
    const solved = solvedState(mod);
    return evaluateSheet(solved.inputs, { names: solved.names, seed: solved.seed });
  }, [mod]);
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
