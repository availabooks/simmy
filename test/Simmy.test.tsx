import { render } from "preact-render-to-string";
import { describe, expect, it } from "vitest";
import { SimContext } from "../src/ui/context.ts";
import { detectPlatform, Simmy, type SimmyProps } from "../src/ui/Simmy.tsx";
import { loadFixture } from "./fixture.ts";

const renderSimmy = (props: SimmyProps = {}) =>
  render(
    <SimContext.Provider value={{ module: loadFixture() }}>
      <Simmy {...props} />
    </SimContext.Provider>,
  );

describe("Simmy", () => {
  const html = renderSimmy({ mode: "kiosk" });

  it("renders Excel with the module's sheet and the task pane", () => {
    expect(html).toContain("Referencing Cells");
    expect(html).toContain("Tax Rate");
    expect(html).toContain("Pool Cover");
    expect(html).toContain("Excel Educator");
    expect(html).toContain("Calculate the sale price in cell ");
    expect(html).toContain("Complete this task for me");
    expect(html).toContain("Assessment Section 1");
    expect(html).toContain("Next");
  });

  it("shows the narration and a poster of the finished sheet", () => {
    expect(html).toContain("Narration");
    expect(html).toContain("Peter's Pool Palace is having an end-of-season pool sale.");
    expect(html).toContain("1.1 Calculate the &quot;Discount&quot; in cell D6.");
  });

  it("offers both desktops and keeps MyEducator out", () => {
    expect(html).toContain("macOS");
    expect(html).toContain("Windows");
    expect(html).not.toMatch(/MyEducator/i);
  });

  it("renders a playback bar", () => {
    expect(html).toContain("Playback");
    expect(html).toContain("Restart");
    expect(html).toContain("Pause");
    expect(html).toContain("Take control");
    expect(html).toContain("Playback speed");
    expect(html).toContain("Next sentence");
    expect(html).toContain("Previous sentence");
  });
});

describe("Simmy modes", () => {
  it("opens a movie behind a Play button and Try it yourself, waiting for the visitor", () => {
    const html = renderSimmy();
    expect(html).toContain('class="xd-intro"');
    expect(html).toContain('aria-label="Play"');
    expect(html).toMatch(/class="xd-intro-control".*Try it yourself/);
    // Nothing is playing yet, so the bar offers Start rather than Pause, with the sound on.
    expect(html).toContain('aria-label="Start"');
    expect(html).not.toContain('aria-label="Pause"');
    expect(html).toContain('aria-label="Mute"');
    // No sticky note until it plays.
    expect(html).toMatch(/class="xd-window xd-window-notes[^"]*is-closed/);
  });

  it("starts a movie in control, with no Play screen, when autoplay is off", () => {
    const html = renderSimmy({ autoplay: false });
    expect(html).not.toContain('class="xd-intro"');
    expect(html).not.toMatch(/class="xd-window xd-window-notes[^"]*is-closed/);
  });

  it("tucks a movie's desktop switcher and Start over into a settings gear on the playback bar", () => {
    const html = renderSimmy();
    expect(html).not.toContain('class="xd-toolbar"');
    // Nothing is left above the desktop, so there is no empty row there.
    expect(html).not.toContain('class="xd-head"');
    expect(html).toContain('aria-label="Settings"');
    expect(html).toContain('aria-expanded="false"');
  });

  it("shows a kiosk's desktop switcher and Start over above the desktop, with no gear", () => {
    const html = renderSimmy({ mode: "kiosk" });
    expect(html).toContain('class="xd-toolbar"');
    expect(html).toContain("Start over");
    expect(html).not.toContain('aria-label="Settings"');
  });

  it("invites the visitor to try it over the desktop while a demo plays, and not before", () => {
    expect(renderSimmy({ mode: "kiosk" })).toMatch(/class="xd-veil".*Try it yourself/);
    expect(renderSimmy()).not.toContain('class="xd-veil"');
    expect(renderSimmy({ mode: "kiosk", autoplay: false })).not.toContain('class="xd-veil"');
  });

  it("plays a kiosk by itself, muted, with no Play screen", () => {
    const html = renderSimmy({ mode: "kiosk" });
    expect(html).not.toContain('class="xd-intro"');
    expect(html).toContain('aria-label="Pause"');
    expect(html).toContain('aria-label="Unmute"');
  });
});

describe("Simmy platforms", () => {
  it("offers macOS, Windows and Web in the switcher", () => {
    const html = renderSimmy({ mode: "kiosk" });
    for (const label of ["macOS", "Windows", "Web"]) {
      expect(html).toContain(`>${label}</button>`);
    }
  });

  it("shows the web view as Excel on its own, without OS or account chrome", () => {
    const html = renderSimmy({ platform: "web" });
    expect(html).toContain('class="xd-excel xd-web"');
    expect(html).toContain('title="Undo"');
    for (const absent of ["xd-dock", "xd-taskbar", "xd-menubar", "xd-avatar", "xd-autosave", "xd-search", "xd-share"]) {
      expect(html).not.toContain(`class="${absent}"`);
    }
    expect(html).not.toContain('title="Minimize"');
    expect(html).not.toContain(">File</button>");
    // The sticky note is still there to carry the steps.
    expect(html).toContain('aria-label="Sticky note"');
  });
});

describe("Simmy ribbon", () => {
  it("starts open, with a button to collapse it", () => {
    const html = renderSimmy();
    expect(html).toContain('class="xd-ribbon"');
    expect(html).toContain('aria-label="Collapse the Ribbon (Ctrl+F1)"');
  });

  it("starts collapsed to its tabs, with no tab selected, when asked", () => {
    const html = renderSimmy({ ribbon: "collapsed" });
    expect(html).toContain('class="xd-ribbon-area is-collapsed"');
    expect(html).not.toContain('class="xd-ribbon"');
    expect(html).toContain(">Formulas</button>");
    expect(html).not.toContain("xd-tab is-active");
  });
});

describe("Simmy idle controls", () => {
  it("dims the ribbon buttons that do nothing and leaves the working ones alone", () => {
    const html = renderSimmy();
    expect(html).toMatch(/class="xd-small is-idle" title="Bold \(not available here\)" aria-disabled="true"/);
    expect(html).toMatch(/class="xd-big is-idle" title="Cell Styles \(not available here\)"/);
    expect(html).toMatch(/class="xd-small" title="Copy"/);
    expect(html).toMatch(/class="xd-big" title="Paste"/);
    expect(html).toContain('class="xd-tab is-idle"');
  });
});

describe("detectPlatform", () => {
  it("picks macOS for Apple devices and Windows otherwise", () => {
    expect(detectPlatform({ userAgentData: { platform: "macOS" } })).toBe("mac");
    expect(detectPlatform({ platform: "MacIntel" })).toBe("mac");
    expect(detectPlatform({ platform: "Win32" })).toBe("win");
    expect(detectPlatform({ platform: "Linux x86_64" })).toBe("win");
  });
});
