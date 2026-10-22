/**
 * A kiosk has to start its narrated demo with no one touching the page. Browsers only allow
 * that for muted media, and Chrome only for muted *video*: a muted `<audio>` (or `new Audio()`)
 * is still refused until the visitor interacts. Checked in Chrome 2026-10-22 on the dev page:
 * new Audio(clip) with muted set rejects with NotAllowedError, and a muted <video> with the
 * same clip plays. These tests give play() that policy and check the kiosk gets going anyway.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { parse } from "yaml";
import { defineSimmy, type SimmyElement } from "../src/element.ts";
import { moduleText } from "./fixture.ts";

defineSimmy();

const settle = (ms = 50) => new Promise((resolve) => setTimeout(resolve, ms));

/** What each play() call was, and whether the policy let it start. */
let plays: { tag: string; muted: boolean; allowed: boolean }[] = [];
/** Whether the visitor has interacted with the page (Chrome's "sticky activation"). */
let activated = false;

function chromeAutoplayPolicy(this: HTMLMediaElement): Promise<void> {
  const allowed = activated || (this instanceof HTMLVideoElement && this.muted);
  plays.push({ tag: this.tagName.toLowerCase(), muted: this.muted, allowed });
  return allowed
    ? Promise.resolve()
    : Promise.reject(new DOMException("play() failed because the user didn't interact with the document first.", "NotAllowedError"));
}

beforeEach(() => {
  plays = [];
  activated = false;
  vi.spyOn(HTMLMediaElement.prototype, "play").mockImplementation(chromeAutoplayPolicy);
  vi.spyOn(HTMLMediaElement.prototype, "pause").mockImplementation(() => {});
  // happy-dom loads no media, so report every clip as loaded and three seconds long.
  vi.spyOn(HTMLMediaElement.prototype, "readyState", "get").mockReturnValue(4);
  vi.spyOn(HTMLMediaElement.prototype, "duration", "get").mockReturnValue(3);
  // happy-dom never lays anything out, so tell the desktop its size the way a browser would.
  vi.stubGlobal(
    "ResizeObserver",
    class {
      constructor(private callback: ResizeObserverCallback) {}
      observe() {
        queueMicrotask(() => this.callback([], this as unknown as ResizeObserver));
      }
      unobserve() {}
      disconnect() {}
    },
  );
});

afterEach(() => {
  document.body.innerHTML = "";
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

/** A kiosk playing the narrated fixture module, given time to reach its first clip. */
async function kiosk(): Promise<ShadowRoot> {
  const el = document.createElement("simmy-app") as SimmyElement;
  el.setAttribute("mode", "kiosk");
  el.setAttribute("fonts", "none");
  document.body.append(el);
  el.module = parse(moduleText());
  for (let i = 0; i < 40 && plays.length === 0; i++) await settle();
  await settle(200);
  return el.shadowRoot as ShadowRoot;
}

describe("a kiosk's narrated demo", () => {
  it("plays its first clip once the visitor has interacted (checks the harness)", async () => {
    activated = true;
    const shadow = await kiosk();
    expect(plays.length).toBeGreaterThan(0);
    expect(plays[0]).toMatchObject({ muted: true, allowed: true });
    expect(shadow.querySelector(".xd-narration-play")).toBeNull();
  });

  it("starts by itself, before anyone touches the page", async () => {
    const shadow = await kiosk();
    expect(plays.length).toBeGreaterThan(0);
    // It plays muted, so the browser should let it start without a click.
    expect(plays[0].muted).toBe(true);
    expect(plays[0].allowed).toBe(true);
    // A kiosk is never clicked: asking for one stalls it on the first line for good.
    expect(shadow.querySelector(".xd-narration-play")).toBeNull();
  });
});
