/** Pause, speed, and timeline for a narrated autopilot pass. */

export type Hold = "play" | "pause" | "control";

export const SKIP_SECONDS = 10;

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

/** Seconds from the start of the timeline to `offset` within `index`. */
export function elapsedAt(durations: readonly number[], index: number, offset: number): number {
  let elapsed = 0;
  for (let i = 0; i < index && i < durations.length; i++) {
    elapsed += Math.max(0, durations[i] ?? 0);
  }
  return elapsed + Math.max(0, Number.isFinite(offset) ? offset : 0);
}
