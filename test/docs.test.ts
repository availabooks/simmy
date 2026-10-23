import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { FUNCTIONS } from "../src/engine/sheet.ts";

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

describe("the function lists", () => {
  const docs = {
    "README.md": readFileSync("README.md", "utf8"),
    "docs/index.html": readFileSync("docs/index.html", "utf8"),
    "skill/simmy/module-format.md": readFileSync("skill/simmy/module-format.md", "utf8"),
  };

  it("name every function the engine has", () => {
    for (const [file, text] of Object.entries(docs)) {
      for (const { name } of FUNCTIONS) {
        expect(text, `${file} lists ${name}`).toMatch(new RegExp(`\\b${name.replace(".", "\\.")}\\b`));
      }
    }
  });
});
