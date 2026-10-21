/** Pause, speed, and timeline for a narrated autopilot pass. */

export type Hold = "play" | "pause" | "control";

export const SPEEDS = [0.5, 1, 1.25, 1.5, 2] as const;

/** Shared by the autopilot loop and the control bar. Pausing does not cancel the pass. */
export class Playback {
  hold: Hold = "play";
  rate = 1;
  private waiting: Array<() => void> = [];

  setHold(hold: Hold) {
    this.hold = hold;
    this.wake();
  }

  setRate(rate: number) {
    this.rate = rate;
  }

  /** Release anyone waiting in `gate`. */
  wake() {
    const waiting = this.waiting;
    this.waiting = [];
    for (const resolve of waiting) {
      resolve();
    }
  }

  /** Resolves immediately while playing. Otherwise waits until play, or until cancelled. */
  async gate(cancelled: () => boolean): Promise<void> {
    while (this.hold !== "play") {
      if (cancelled()) {
        return;
      }
      await new Promise<void>((resolve) => {
        this.waiting.push(resolve);
      });
    }
  }
}

/** `m:ss`, rounding down. */
export function formatStamp(seconds: number): string {
  const safe = Number.isFinite(seconds) ? Math.max(0, Math.floor(seconds)) : 0;
  const minutes = Math.floor(safe / 60);
  const remain = safe % 60;
  return `${minutes}:${remain.toString().padStart(2, "0")}`;
}

/** Map a position on the concatenated clips to a step and a second offset into that step. */
export function locateTime(durations: readonly number[], seconds: number): { index: number; offset: number } {
  const safe = Number.isFinite(seconds) ? Math.max(0, seconds) : 0;
  if (durations.length === 0) {
    return { index: 0, offset: 0 };
  }
  let cursor = 0;
  for (let index = 0; index < durations.length; index++) {
    const span = Math.max(0, durations[index] ?? 0);
    const last = index === durations.length - 1;
    if (!last && safe >= cursor + span) {
      cursor += span;
      continue;
    }
    return { index, offset: Math.min(Math.max(0, safe - cursor), span) };
  }
  return { index: durations.length - 1, offset: 0 };
}

export function sumDurations(durations: readonly number[]): number {
  return durations.reduce((sum, span) => sum + (Number.isFinite(span) ? Math.max(0, span) : 0), 0);
}

export type SeekPlan =
  | { kind: "wait" }
  | { kind: "jump"; index: number; offset: number; rush: boolean; leaveClip: boolean }
  | { kind: "relaunch"; index: number; offset: number; rush: boolean };

/**
 * Where the playhead is. A seek moves the bookmark before the clip's clock catches up,
 * and the clip is ahead again once it is actually playing from there.
 */
export function playhead(bookmark: number, audioTime: number | null): number {
  if (audioTime === null || !Number.isFinite(audioTime)) {
    return bookmark;
  }
  return Math.max(bookmark, audioTime);
}

/** What a skip or a drag should do from the current playhead. */
export function planSeek(input: {
  durations: readonly number[];
  known: readonly boolean[];
  index: number;
  into: number;
  seconds: number;
  session: boolean;
}): SeekPlan {
  if (input.known.length === 0 || input.known.some((ready) => !ready)) {
    return { kind: "wait" };
  }
  const seconds = Math.min(Math.max(0, Number.isFinite(input.seconds) ? input.seconds : 0), sumDurations(input.durations));
  const loc = locateTime(input.durations, seconds);
  const ahead = loc.index > input.index || (loc.index === input.index && loc.offset >= input.into - 0.05);
  const rushed = loc.index === input.index ? loc.offset > input.into + 0.35 : loc.offset > 0.35;
  if (ahead && input.session) {
    return {
      kind: "jump",
      index: loc.index,
      offset: loc.offset,
      rush: rushed,
      leaveClip: loc.index > input.index,
    };
  }
  return { kind: "relaunch", index: loc.index, offset: loc.offset, rush: loc.offset > 0.35 };
}

/**
 * One narration clip forward or back. Each clip is one sentence, so this does not
 * jump a block of time that can cover several of them.
 */
export function skipTrack(durations: readonly number[], index: number, direction: "back" | "forward"): number {
  const last = Math.max(0, durations.length - 1);
  const at = Math.min(Math.max(0, index), last);
  if (direction === "forward") {
    if (at >= last) {
      return sumDurations(durations);
    }
    return elapsedAt(durations, at + 1, 0);
  }
  if (at === 0) {
    return 0;
  }
  return elapsedAt(durations, at - 1, 0);
}

/** Seconds from the start of the timeline to `offset` within `index`. */
export function elapsedAt(durations: readonly number[], index: number, offset: number): number {
  let elapsed = 0;
  for (let i = 0; i < index && i < durations.length; i++) {
    elapsed += Math.max(0, durations[i] ?? 0);
  }
  return elapsed + Math.max(0, Number.isFinite(offset) ? offset : 0);
}
