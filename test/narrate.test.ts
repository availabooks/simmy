import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { parse } from "yaml";
import { narrateModule } from "../src/cli/narrate.ts";
import { DEFAULT_VOICE, type Synthesize } from "../src/cli/tts.ts";

const dirs: string[] = [];
afterEach(() => {
  for (const dir of dirs) rmSync(dir, { recursive: true, force: true });
  dirs.length = 0;
});

/** A module folder whose config has these autopilot lines, and an audio folder with these files. */
function moduleWith(autopilot: string, files: string[] = []): string {
  const dir = mkdtempSync(join(tmpdir(), "simmy-narrate-"));
  dirs.push(dir);
  mkdirSync(join(dir, "audio"));
  for (const file of files) writeFileSync(join(dir, "audio", file), "recorded");
  writeFileSync(join(dir, "config.yaml"), `tasks: []\n# the demo\nautopilot:\n${autopilot}poster:\n  range: B5:E9\n`);
  return join(dir, "config.yaml");
}

/** Records what it was asked to speak and writes the text as the "clip". */
function fakeSpeech() {
  const spoken: string[] = [];
  const synthesize: Synthesize = async (text, path) => {
    spoken.push(text);
    writeFileSync(path, text);
  };
  return { spoken, synthesize };
}

const options = { voice: DEFAULT_VOICE };

describe("narrateModule", () => {
  it("names new clips after the ones in use and writes them into the config", async () => {
    const config = moduleWith(
      [
        "  - { audio: 1.mp3, transcript: Recorded. }",
        "  - { transcript: \"Flow style.\" }",
        "  - task: A",
        "    # kept comment",
        "    transcript: Block style.",
        "",
      ].join("\n"),
      ["1.mp3", "4.mp3"],
    );
    const { spoken, synthesize } = fakeSpeech();
    const results = await narrateModule(config, options, synthesize);

    expect(results.map((r) => [r.audio, r.status])).toEqual([
      ["1.mp3", "kept"],
      ["5.mp3", "made"],
      ["6.mp3", "made"],
    ]);
    expect(spoken).toEqual(["Flow style.", "Block style."]);
    const text = readFileSync(config, "utf8");
    expect(text).toContain('  - { audio: 5.mp3, transcript: "Flow style." }');
    expect(text).toContain("    # kept comment\n    audio: 6.mp3\n    transcript: Block style.");
    expect(text).toContain("# the demo");
    expect(parse(text).autopilot[2]).toEqual({ task: "A", audio: "6.mp3", transcript: "Block style." });
    expect(readFileSync(join(config, "..", "audio", "6.mp3"), "utf8")).toBe("Block style.");
  });

  it("remakes a clip when its transcript or voice changes, and only then", async () => {
    const config = moduleWith("  - { transcript: First. }\n");
    const { spoken, synthesize } = fakeSpeech();
    await narrateModule(config, options, synthesize);
    expect((await narrateModule(config, options, synthesize))[0].status).toBe("current");

    writeFileSync(config, readFileSync(config, "utf8").replace("First.", "Second."));
    expect((await narrateModule(config, options, synthesize))[0].status).toBe("remade");
    const slower = { voice: { ...DEFAULT_VOICE, rate: "-10%" } };
    expect((await narrateModule(config, slower, synthesize))[0].status).toBe("remade");
    expect(spoken).toEqual(["First.", "Second.", "Second."]);
  });

  it("replaces clips it didn't make only with force", async () => {
    const config = moduleWith("  - { audio: 1.mp3, transcript: Hi. }\n", ["1.mp3"]);
    const { spoken, synthesize } = fakeSpeech();
    expect((await narrateModule(config, options, synthesize))[0].status).toBe("kept");
    expect((await narrateModule(config, { ...options, force: true }, synthesize))[0].status).toBe("remade");
    expect(spoken).toEqual(["Hi."]);
  });

  it("changes nothing on a dry run", async () => {
    const config = moduleWith("  - { transcript: Hi. }\n");
    const before = readFileSync(config, "utf8");
    const { spoken, synthesize } = fakeSpeech();
    const results = await narrateModule(config, { ...options, dryRun: true }, synthesize);
    expect(results[0]).toMatchObject({ audio: "1.mp3", status: "made" });
    expect(spoken).toEqual([]);
    expect(readFileSync(config, "utf8")).toBe(before);
    expect(existsSync(join(config, "..", "narration.json"))).toBe(false);
  });

  it("leaves remote clips alone and reports failures without stopping", async () => {
    const config = moduleWith(
      "  - { audio: https://cdn.example.com/1.mp3, transcript: Remote. }\n  - { transcript: Breaks. }\n  - { transcript: Works. }\n",
    );
    const synthesize: Synthesize = async (text, path) => {
      if (text === "Breaks.") throw new Error("service unavailable");
      writeFileSync(path, text);
    };
    const results = await narrateModule(config, options, synthesize);
    expect(results.map((r) => r.status)).toEqual(["remote", "failed", "made"]);
    expect(results[1].error).toBe("service unavailable");
    // The failed clip has no record, so the next run tries it again.
    const manifest = JSON.parse(readFileSync(join(config, "..", "narration.json"), "utf8"));
    expect(Object.keys(manifest)).toEqual(["2.mp3"]);
  });
});
