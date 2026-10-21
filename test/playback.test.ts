import { describe, expect, it } from "vitest";
import { Playback, elapsedAt, formatStamp, locateTime, planSeek, playhead, skipTrack, sumDurations } from "../src/ui/playback.ts";

describe("timeline", () => {
  const durations = [10, 20, 5];

  it("maps a time onto the clip that contains it", () => {
    expect(locateTime(durations, 0)).toEqual({ index: 0, offset: 0 });
    expect(locateTime(durations, 9.5)).toEqual({ index: 0, offset: 9.5 });
    expect(locateTime(durations, 10)).toEqual({ index: 1, offset: 0 });
    expect(locateTime(durations, 25)).toEqual({ index: 1, offset: 15 });
    expect(locateTime(durations, 40)).toEqual({ index: 2, offset: 5 });
    expect(locateTime([], 4)).toEqual({ index: 0, offset: 0 });
  });

  it("adds up elapsed time and the whole timeline", () => {
    expect(elapsedAt(durations, 1, 5)).toBe(15);
    expect(elapsedAt(durations, 0, 0)).toBe(0);
    expect(sumDurations(durations)).toBe(35);
    expect(formatStamp(0)).toBe("0:00");
    expect(formatStamp(65)).toBe("1:05");
    expect(formatStamp(59.9)).toBe("0:59");
  });
});

describe("planSeek", () => {
  const durations = [10, 20, 5];
  const known = [true, true, true];

  it("leaves the clip that is playing when the target is a later one", () => {
    expect(
      planSeek({ durations, known, index: 0, into: 3, seconds: 15, session: true }),
    ).toEqual({ kind: "jump", index: 1, offset: 5, rush: true, leaveClip: true });
    expect(elapsedAt(durations, 1, 5)).toBe(15);
  });

  it("stays on the current clip when the target is later in it", () => {
    expect(
      planSeek({ durations, known, index: 0, into: 2, seconds: 8, session: true }),
    ).toEqual({ kind: "jump", index: 0, offset: 8, rush: true, leaveClip: false });
  });

  it("restarts from an earlier time", () => {
    expect(
      planSeek({ durations, known, index: 1, into: 5, seconds: 4, session: true }),
    ).toEqual({ kind: "relaunch", index: 0, offset: 4, rush: true });
  });

  it("does not seek until every clip length is known", () => {
    expect(
      planSeek({ durations, known: [true, false, true], index: 0, into: 1, seconds: 12, session: true }),
    ).toEqual({ kind: "wait" });
  });

  it("moves a second skip from the time the first one already chose", () => {
    const first = planSeek({ durations: [30, 30], known: [true, true], index: 0, into: 2, seconds: 12, session: true });
    expect(first).toMatchObject({ kind: "jump", index: 0, offset: 12 });
    const stuckClock = 2;
    const from = playhead(first.kind === "jump" ? first.offset : 0, stuckClock);
    const second = planSeek({
      durations: [30, 30],
      known: [true, true],
      index: 0,
      into: from,
      seconds: from + 10,
      session: true,
    });
    expect(second).toMatchObject({ kind: "jump", offset: 22 });
  });
});

describe("skipTrack", () => {
  const durations = [4, 3, 8];

  it("moves one sentence, not a block of time", () => {
    expect(skipTrack(durations, 0, "forward")).toBe(4);
    expect(skipTrack(durations, 1, "forward")).toBe(7);
    expect(skipTrack(durations, 1, "back")).toBe(0);
    expect(skipTrack(durations, 2, "forward")).toBe(15);
    expect(skipTrack(durations, 0, "back")).toBe(0);
  });
});

describe("playhead", () => {
  it("keeps a seek the clip has not reached, then follows the clip", () => {
    expect(playhead(12, 2)).toBe(12);
    expect(playhead(12, 14)).toBe(14);
    expect(playhead(12, null)).toBe(12);
  });
});

describe("Playback", () => {
  it("holds the pass until play", async () => {
    const playback = new Playback();
    playback.setHold("pause");
    let released = false;
    const pending = playback.gate(() => false).then(() => {
      released = true;
    });
    await Promise.resolve();
    expect(released).toBe(false);
    playback.setHold("control");
    await Promise.resolve();
    expect(released).toBe(false);
    playback.setHold("play");
    await pending;
    expect(released).toBe(true);
  });

  it("lets a cancelled pass leave the gate", async () => {
    const playback = new Playback();
    playback.setHold("control");
    let cancelled = false;
    const pending = playback.gate(() => cancelled);
    await Promise.resolve();
    cancelled = true;
    playback.wake();
    await pending;
    expect(playback.hold).toBe("control");
  });
});
