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
    </div>
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
