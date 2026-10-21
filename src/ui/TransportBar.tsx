import type { ComponentChildren } from "preact";
import { useEffect, useRef, useState } from "preact/hooks";
import { CursorMark } from "./CursorMark.tsx";
import type { Hold } from "./playback.ts";
import { SPEEDS, formatStamp } from "./playback.ts";

export type TransportBarProps = {
  hold: Hold;
  /** A pass is in progress, paused, or waiting while the visitor has control. */
  session: boolean;
  rate: number;
  elapsed: number;
  total: number;
  /** Durations are loaded, so the scrubber and skip buttons can move through the timeline. */
  ready: boolean;
  scrub: number | null;
  onScrub: (seconds: number | null) => void;
  onSeek: (seconds: number) => void;
  onPlay: () => void;
  onPause: () => void;
  onRestart: () => void;
  onRewind: () => void;
  onFastForward: () => void;
  onTakeControl: () => void;
  muted: boolean;
  onMute: () => void;
  onRate: (rate: number) => void;
  /** Controls for a gear at the end of the bar, shown in a panel above it. Without them there is no gear. */
  settings?: ComponentChildren;
};

const shownOf = (seconds: number, ready: boolean) => (ready ? formatStamp(seconds) : "–:––");

export function TransportBar(props: TransportBarProps) {
  const shown = props.scrub ?? props.elapsed;
  const playing = props.session && props.hold === "play";
  const playLabel = playing ? "Pause" : props.session ? "Resume" : "Start";
  return (
    <div class="xd-transport" role="group" aria-label="Playback">
      <button type="button" class="xd-transport-btn" title="Restart" aria-label="Restart" onClick={props.onRestart}>
        <RestartIcon />
      </button>
      <button
        type="button"
        class="xd-transport-btn"
        title="Previous sentence"
        aria-label="Previous sentence"
        disabled={!props.ready || shown <= 0}
        onClick={props.onRewind}
      >
        <SkipIcon direction="back" />
      </button>
      <button
        type="button"
        class={`xd-transport-btn${playing ? "" : " is-primary"}`}
        title={playLabel}
        aria-label={playLabel}
        onClick={playing ? props.onPause : props.onPlay}
      >
        {playing ? <PauseIcon /> : <PlayIcon />}
      </button>
      <button
        type="button"
        class="xd-transport-btn"
        title="Next sentence"
        aria-label="Next sentence"
        disabled={!props.ready || props.total <= 0 || shown >= props.total}
        onClick={props.onFastForward}
      >
        <SkipIcon direction="forward" />
      </button>
      <span class="xd-transport-tipwrap">
        <button
          type="button"
          class={`xd-transport-btn${props.hold === "control" ? " is-on" : ""}`}
          aria-label="Take control"
          aria-pressed={props.hold === "control"}
          disabled={!props.session || props.hold === "control"}
          onClick={props.onTakeControl}
        >
          <CursorMark size={16} />
        </button>
        <span class="xd-transport-tip" role="tooltip">
          {props.hold === "control" ? "You have control" : "Take control"}
        </span>
      </span>
      <input
        class="xd-transport-range"
        type="range"
        min={0}
        max={props.ready && props.total > 0 ? props.total : 0}
        step={0.1}
        value={props.ready ? Math.min(shown, props.total) : 0}
        disabled={!props.ready || props.total <= 0}
        aria-label="Playback position"
        aria-valuetext={props.ready ? `${formatStamp(shown)} of ${formatStamp(props.total)}` : "Loading"}
        onInput={(event) => props.onScrub(Number(event.currentTarget.value))}
        onChange={(event) => props.onSeek(Number(event.currentTarget.value))}
      />
      <span class="xd-transport-time">
        {shownOf(shown, props.ready)}
        <span> / {shownOf(props.total, props.ready)}</span>
      </span>
      <span class="xd-transport-tipwrap">
        <button
          type="button"
          class={`xd-transport-btn${props.muted ? " is-on" : ""}`}
          aria-label={props.muted ? "Unmute" : "Mute"}
          aria-pressed={props.muted}
          onClick={props.onMute}
        >
          <MuteIcon muted={props.muted} />
        </button>
        <span class="xd-transport-tip" role="tooltip">
          {props.muted ? "Unmute" : "Mute"}
        </span>
      </span>
      <select
        class="xd-transport-speed"
        aria-label="Playback speed"
        value={props.rate}
        onChange={(event) => props.onRate(Number(event.currentTarget.value))}
      >
        {SPEEDS.map((speed) => (
          <option key={speed} value={speed}>
            {speed}×
          </option>
        ))}
      </select>
      {props.settings && <Settings>{props.settings}</Settings>}
    </div>
  );
}

/** The gear and its panel. A click outside the panel, or Escape, closes it. */
function Settings({ children }: { children: ComponentChildren }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLSpanElement>(null);
  useEffect(() => {
    if (!open) {
      return;
    }
    // composedPath sees through a shadow root, where event.target is retargeted to the host.
    const onDown = (event: PointerEvent) => {
      if (ref.current && !event.composedPath().includes(ref.current)) {
        setOpen(false);
      }
    };
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setOpen(false);
      }
    };
    document.addEventListener("pointerdown", onDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("pointerdown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);
  return (
    <span class="xd-transport-tipwrap xd-settings" ref={ref}>
      <button
        type="button"
        class={`xd-transport-btn${open ? " is-on" : ""}`}
        aria-label="Settings"
        aria-expanded={open}
        aria-haspopup="true"
        onClick={() => setOpen(!open)}
      >
        <GearIcon />
      </button>
      {!open && (
        <span class="xd-transport-tip" role="tooltip">
          Settings
        </span>
      )}
      {open && (
        <div class="xd-settings-panel" role="group" aria-label="Settings">
          {children}
        </div>
      )}
    </span>
  );
}

function GearIcon() {
  return (
    <svg viewBox="0 0 24 24" width="17" height="17" aria-hidden="true">
      <path
        fill="currentColor"
        fill-rule="evenodd"
        d="M19.14 12.94a7.1 7.1 0 000-1.88l2.03-1.58a.5.5 0 00.12-.61l-1.92-3.32a.5.5 0 00-.59-.22l-2.39.96a7 7 0 00-1.62-.94l-.36-2.54a.48.48 0 00-.48-.41h-3.84a.48.48 0 00-.47.41l-.36 2.54a7.3 7.3 0 00-1.62.94l-2.39-.96a.5.5 0 00-.59.22L2.74 8.87a.48.48 0 00.12.61l2.03 1.58a7.1 7.1 0 000 1.88l-2.03 1.58a.5.5 0 00-.12.61l1.92 3.32a.5.5 0 00.59.22l2.39-.96a7 7 0 001.62.94l.36 2.54a.49.49 0 00.48.41h3.84a.48.48 0 00.47-.41l.36-2.54a7.3 7.3 0 001.62-.94l2.39.96a.5.5 0 00.59-.22l1.92-3.32a.48.48 0 00-.12-.61zM12 15.6a3.6 3.6 0 110-7.2 3.6 3.6 0 010 7.2z"
      />
    </svg>
  );
}

function RestartIcon() {
  return (
    <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
      <path
        fill="currentColor"
        d="M12 5V1L7 6l5 5V7c3.31 0 6 2.69 6 6s-2.69 6-6 6-6-2.69-6-6H4c0 4.42 3.58 8 8 8s8-3.58 8-8-3.58-8-8-8z"
      />
    </svg>
  );
}

function MuteIcon({ muted }: { muted: boolean }) {
  return (
    <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
      <path fill="currentColor" d="M4 9.2h3.4L12 5v14l-4.6-4.2H4z" />
      {muted ? (
        <path
          fill="none"
          stroke="currentColor"
          stroke-width="1.8"
          stroke-linecap="round"
          d="M16 9.2l5 5.6M21 9.2l-5 5.6"
        />
      ) : (
        <path
          fill="none"
          stroke="currentColor"
          stroke-width="1.8"
          stroke-linecap="round"
          d="M16 9a4.2 4.2 0 010 6M18.4 6.6a7.4 7.4 0 010 10.8"
        />
      )}
    </svg>
  );
}

function PlayIcon() {
  return (
    <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
      <path d="M8 5.5v13l11-6.5z" fill="currentColor" />
    </svg>
  );
}

function PauseIcon() {
  return (
    <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
      <path d="M7 5h3.2v14H7zm6.8 0H17v14h-3.2z" fill="currentColor" />
    </svg>
  );
}

function SkipIcon({ direction }: { direction: "back" | "forward" }) {
  const back = direction === "back";
  return (
    <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
      <path d={back ? "M18 6l-8 6 8 6zM8 6H6v12h2z" : "M6 6l8 6-8 6zM16 6h2v12h-2z"} fill="currentColor" />
    </svg>
  );
}
