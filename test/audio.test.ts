import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { parse } from "yaml";
import { resolveAudioUrl, resolveModuleAudio } from "../src/module/audio.ts";
import { relativeAudioRefs, rewriteHostedModule } from "../src/module/host.ts";
import { parseModule } from "../src/module/parse.ts";
import { publishHostedModules } from "../scripts/host-modules.ts";
import { loadFixture, moduleText } from "./fixture.ts";

const moduleUrl = "https://modules.example.com/references-and-calculations/config.yaml";

describe("resolveAudioUrl", () => {
  it("resolves a clip name from the audio folder beside the module", () => {
    expect(resolveAudioUrl("1.mp3", moduleUrl)).toBe(
      "https://modules.example.com/references-and-calculations/audio/1.mp3",
    );
    expect(resolveAudioUrl("./takes/2.mp3", "http://localhost:5173/hosted-modules/pool/config.yaml")).toBe(
      "http://localhost:5173/hosted-modules/pool/audio/takes/2.mp3",
    );
  });

  it("leaves absolute locations alone", () => {
    const root = "/modules/narration/1.mp3";
    const remote = "https://cdn.example.com/1.mp3";
    expect(resolveAudioUrl(root, moduleUrl)).toBe(root);
    expect(resolveAudioUrl(remote, moduleUrl)).toBe(remote);
  });

  it("rejects a path that leaves the audio folder", () => {
    expect(() => resolveAudioUrl("../secret.mp3", moduleUrl)).toThrow(/audio folder/);
  });
});

describe("rewriteHostedModule", () => {
  it("rewrites relative clips and leaves the rest of the module loadable", () => {
    const text = moduleText();
    expect(relativeAudioRefs(text)).toContain("1.mp3");
    const hosted = rewriteHostedModule(text, moduleUrl);
    const data = parse(hosted) as { autopilot: { audio?: string }[]; sheet: { cells: { C2: string } } };
    expect(data.autopilot[0]?.audio).toBe("https://modules.example.com/references-and-calculations/audio/1.mp3");
    expect(data.autopilot[4]?.audio).toBe("https://modules.example.com/references-and-calculations/audio/5.mp3");
    expect(data.sheet.cells.C2).toBe(0.3);
    const mod = parseModule(parse(hosted));
    expect(mod.tasks).toHaveLength(4);
    expect(resolveModuleAudio(loadFixture(), moduleUrl).autopilot.at(-1)).toMatchObject({
      audio: "https://modules.example.com/references-and-calculations/audio/16.mp3",
    });
  });
});

describe("publishHostedModules", () => {
  const dirs: string[] = [];
  afterEach(() => {
    for (const dir of dirs) rmSync(dir, { recursive: true, force: true });
    dirs.length = 0;
  });

  it("writes absolute clip URLs, the audio files, and a CORS header", () => {
    const out = mkdtempSync(join(tmpdir(), "simmy-host-"));
    dirs.push(out);
    publishHostedModules("hosted-modules", out, "https://modules.example.com");

    const config = readFileSync(join(out, "references-and-calculations", "config.yaml"), "utf8");
    expect(config).toContain("https://modules.example.com/references-and-calculations/audio/1.mp3");
    expect(readFileSync(join(out, "references-and-calculations", "audio", "1.mp3")).byteLength).toBeGreaterThan(0);
    expect(readFileSync(join(out, "_headers"), "utf8")).toContain("Access-Control-Allow-Origin: *");
    expect(readFileSync(join(out, "index.html"), "utf8")).toContain(
      'href="/references-and-calculations/config.yaml"',
    );
  });

  it("puts a site at the root and moves the module list to modules.html", () => {
    const out = mkdtempSync(join(tmpdir(), "simmy-host-"));
    dirs.push(out);
    publishHostedModules("hosted-modules", out, "https://modules.example.com", "docs");

    expect(readFileSync(join(out, "index.html"), "utf8")).toContain("<title>simmy docs</title>");
    expect(readFileSync(join(out, "docs.css"), "utf8").length).toBeGreaterThan(0);
    expect(readFileSync(join(out, "modules.html"), "utf8")).toContain(
      'href="/references-and-calculations/config.yaml"',
    );
  });
});
