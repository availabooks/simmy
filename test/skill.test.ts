import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { parse } from "yaml";
import { gradeAssignment, initialInputs, solvedInputs } from "../src/module/grade.ts";
import { MODULE, RULE_KEYS, type Shape, unknownKeys } from "../src/module/keys.ts";
import { parseModule } from "../src/module/parse.ts";

const read = (file: string) => readFileSync(`skill/simmy/${file}`, "utf8");

describe("the Claude skill", () => {
  it("has the frontmatter a skill needs", () => {
    const front = /^---\n([\s\S]*?)\n---\n/.exec(read("SKILL.md"));
    expect(front).not.toBeNull();
    const meta = parse(front![1]) as { name: string; description: string };
    expect(meta.name).toBe("simmy");
    expect(meta.description.length).toBeGreaterThan(50);
    expect(meta.description.length).toBeLessThanOrEqual(1024);
  });

  it("pins the version in package.json", () => {
    const { version } = JSON.parse(readFileSync("package.json", "utf8")) as { version: string };
    const pinned = [...read("SKILL.md").matchAll(/@availabooks\/simmy@([^/"\s]+)/g)].map((match) => match[1]);
    expect(pinned.length).toBeGreaterThan(0);
    expect(new Set(pinned)).toEqual(new Set([version]));
  });

  it("starts from an example that loads, has no stray keys, and grades its own solutions", () => {
    const data = parse(read("example-module.yaml"));
    expect(unknownKeys(data)).toEqual([]);
    const mod = parseModule(data);
    expect(gradeAssignment(mod, solvedInputs(mod)).earned).toBe(mod.totalPoints);
    expect(gradeAssignment(mod, initialInputs(mod)).earned).toBe(0);
  });

  it("documents every module key", () => {
    const doc = read("module-format.md");
    const keys = new Set<string>(Object.values(RULE_KEYS).flat());
    const walk = (shape: Shape | Shape[] | true) => {
      if (shape === true || shape === "any-keys") return;
      if (Array.isArray(shape)) return walk(shape[0]);
      for (const [key, child] of Object.entries(shape)) {
        keys.add(key);
        walk(child);
      }
    };
    walk(MODULE);
    for (const key of keys) expect(doc, key).toMatch(new RegExp(`\\b${key}\\b`));
  });
});
