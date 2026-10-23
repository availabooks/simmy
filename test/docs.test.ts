import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const { version } = JSON.parse(readFileSync("package.json", "utf8")) as { version: string };

describe("the docs site", () => {
  const html = readFileSync("docs/index.html", "utf8");

  it("loads the latest release and shows the version in package.json", () => {
    const pinned = [...html.matchAll(/@availabooks\/simmy@([^/"\s]+)/g)].map((match) => match[1]);
    expect(pinned.length).toBeGreaterThan(0);
    expect(new Set(pinned)).toEqual(new Set(["latest"]));
    expect(html).toContain(`<span class="version">v${version}</span>`);
  });
});
