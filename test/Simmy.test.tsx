import { render } from "preact-render-to-string";
import { describe, expect, it } from "vitest";
import { SimContext } from "../src/ui/context.ts";
import { detectPlatform, Simmy } from "../src/ui/Simmy.tsx";
import { loadFixture } from "./fixture.ts";

describe("Simmy", () => {
  const html = render(
    <SimContext.Provider value={{ module: loadFixture() }}>
      <Simmy />
    </SimContext.Provider>,
  );

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
    expect(html).toContain("Mute");
    expect(html).toContain("Playback speed");
    expect(html).toContain("Fast-forward 10 seconds");
    expect(html).toContain("Rewind 10 seconds");
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
