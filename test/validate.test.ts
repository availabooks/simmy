import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { publishHostedModules } from "../src/cli/host.ts";
import { validateModuleFile, validatePath } from "../src/cli/validate.ts";

const dirs: string[] = [];
afterEach(() => {
  for (const dir of dirs) rmSync(dir, { recursive: true, force: true });
  dirs.length = 0;
});

/** A copy of the hosted modules that a test can break. */
function copyModules(): string {
  const dir = mkdtempSync(join(tmpdir(), "simmy-validate-"));
  dirs.push(dir);
  cpSync("hosted-modules", dir, { recursive: true });
  return dir;
}

describe("validatePath", () => {
  it("passes every hosted module", () => {
    const reports = validatePath("hosted-modules");
    expect(reports.map((report) => report.name)).toContain("references-and-calculations");
    for (const report of reports) expect(report.problems, report.name).toEqual([]);
  });

  it("takes a module folder or its config file", () => {
    expect(validatePath("hosted-modules/feature-tour").map((report) => report.name)).toEqual(["feature-tour"]);
    expect(validatePath("hosted-modules/feature-tour/config.yaml").map((report) => report.name)).toEqual([
      "feature-tour",
    ]);
  });

  it("throws for a path that doesn't exist", () => {
    expect(() => validatePath("no-such-folder")).toThrow(/Nothing at/);
  });
});

describe("validateModuleFile", () => {
  it("reports YAML syntax errors", () => {
    const dir = copyModules();
    const config = join(dir, "feature-tour", "config.yaml");
    writeFileSync(config, "title: [unclosed\n");
    expect(validateModuleFile(config).problems[0]).toMatch(/^YAML: /);
  });

  it("reports module problems with their paths", () => {
    const dir = copyModules();
    const config = join(dir, "feature-tour", "config.yaml");
    writeFileSync(config, "- not a mapping\n");
    expect(validateModuleFile(config).problems).toContain("module: must be a mapping");
  });

  it("reports missing clips and clips outside the audio folder", () => {
    const dir = copyModules();
    const config = join(dir, "references-and-calculations", "config.yaml");
    const text = readFileSync(config, "utf8");
    rmSync(join(dir, "references-and-calculations", "audio", "1.mp3"));
    writeFileSync(config, text.replace(/audio: 2\.mp3/, "audio: ../2.mp3"));
    const problems = validateModuleFile(config).problems;
    expect(problems).toContain('audio "1.mp3": missing audio/1.mp3');
    expect(problems).toContain('audio "../2.mp3": leaves the audio folder');
  });

  it("reports a module folder without a config", () => {
    const dir = copyModules();
    mkdirSync(join(dir, "empty"));
    const empty = validatePath(dir).find((report) => report.name === "empty");
    expect(empty?.problems).toEqual(["missing config.yaml"]);
  });
});

describe("publishing", () => {
  it("refuses to build when a module is invalid", () => {
    const dir = copyModules();
    writeFileSync(join(dir, "feature-tour", "config.yaml"), "- not a mapping\n");
    const out = mkdtempSync(join(tmpdir(), "simmy-host-"));
    dirs.push(out);
    expect(() => publishHostedModules(dir, out, "https://modules.example.com")).toThrow(
      /feature-tour: module: must be a mapping/,
    );
  });
});
