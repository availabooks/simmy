/**
 * The `simmy` command. `simmy publish` hosts a folder of modules as a Cloudflare Workers
 * static-assets site, so a page on any origin can load a module and play its clips.
 * `simmy validate` checks modules without publishing them. `simmy narrate` and `simmy tts`
 * make clips from text with Edge's read-aloud voices.
 */
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { publishHostedModules } from "./host.ts";
import { type ClipResult, narrateModule } from "./narrate.ts";
import { DEFAULT_VOICE, listVoices, synthesize, type Voice } from "./tts.ts";
import { type ModuleReport, moduleConfigs, validateModuleFile, validatePath } from "./validate.ts";

const USAGE = `Usage: simmy <command> [options]

  simmy publish [--origin <url>] [--name <worker>] [--source <dir>] [--dry-run]

Host the modules in a folder (<name>/config.yaml and <name>/audio/) as a Cloudflare
Workers static site. Needs wrangler to be logged in (npx wrangler login).
  --origin <url>    Public URL of the Worker. It is written into the clip URLs, so it
                    must match where the Worker is served.
  --name <worker>   Worker name (default: the origin's first hostname label).
  --source <dir>    Folder of modules (default: hosted-modules).
  --dry-run         Build the site into a folder and print its path; don't deploy.
Every module is validated first, and nothing deploys if one has a problem.

  simmy validate [--strict] [path...]

Check modules the way the simulator loads them, and that their clips are in the audio
folder. A path is a config.yaml, a module folder, or a folder of modules (default: the
source folder, as for publish). Exits with 1 if any module has a problem. Keys simmy
doesn't know (often misspellings) are warnings; --strict makes them problems too.

  simmy narrate [--force] [--dry-run] [voice options] [path...]

Speak each autopilot transcript into the module's audio folder. An entry with a
transcript and no audio gets the next numbered clip (written into config.yaml).
narration.json, beside config.yaml, records what each clip was made from, so a clip
is made again when its transcript or voice changes. Clips it has no record of (made
some other way) are kept unless --force. --dry-run shows what would happen. Paths as
for validate.

  simmy tts --out <file.mp3> [voice options] <text...>

Speak one piece of text into an MP3.

  simmy voices [filter]

List the voices, optionally only those whose name or locale contains the filter
(e.g. simmy voices en-US).

Voice options, for narrate and tts:
  --voice <name>    Default ${DEFAULT_VOICE.voice}.
  --rate <percent>  Speaking rate, e.g. +10% or -5%.
  --pitch <hz>      Pitch, e.g. +2Hz or -5Hz.
Voices come from Microsoft Edge's read-aloud service: no account needed, but it is
unofficial and can change.

  simmy config [set <key> <value> | unset <key> | path]

Defaults can live in simmy.json next to where you run the command:
  { "origin": "https://my-modules.example.workers.dev", "name": "my-modules", "source": "modules" }
Keys: origin, name, source, voice, rate, pitch.

To use the same project from any folder, save them globally instead:
  simmy config set origin https://my-modules.example.workers.dev
  simmy config set source ~/work/my-modules     (made absolute when saved)
Precedence: flags, then ./simmy.json, then the global config.
`;

const KEYS = ["origin", "name", "source", "voice", "rate", "pitch"] as const;
type Key = (typeof KEYS)[number];
type Options = Partial<Record<Key, string>>;

/**
 * Split a command's arguments into `--flag value` pairs, `--switch`es that are on, and the
 * rest, in order. Any other `-` argument is an error.
 */
function parseFlags(
  args: string[],
  valued: string[],
  switches: string[] = [],
): { values: Record<string, string>; on: Set<string>; rest: string[] } {
  const values: Record<string, string> = {};
  const on = new Set<string>();
  const rest: string[] = [];
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    const name = arg.replace(/^--/, "");
    if (!arg.startsWith("-") || arg === "-") {
      rest.push(arg);
    } else if (arg.startsWith("--") && switches.includes(name)) {
      on.add(name);
    } else if (arg.startsWith("--") && valued.includes(name)) {
      const value = args[++i];
      if (value === undefined) throw new Error(`${arg} needs a value`);
      values[name] = value;
    } else {
      throw new Error(`Unknown option ${arg}`);
    }
  }
  return { values, on, rest };
}

/** Where `simmy config` keeps the defaults that apply from any folder. */
function globalConfigPath(): string {
  const base = process.env.XDG_CONFIG_HOME || join(homedir(), ".config");
  return join(base, "simmy", "config.json");
}

function readJson(path: string): Options {
  if (!existsSync(path)) return {};
  try {
    return JSON.parse(readFileSync(path, "utf8"));
  } catch (error) {
    throw new Error(`${path}: ${(error as Error).message}`);
  }
}

function readConfig(): Options {
  return { ...readJson(globalConfigPath()), ...readJson("simmy.json") };
}

function config(args: string[]): number {
  const path = globalConfigPath();
  const [action, key, value] = args;
  if (action === "path") {
    console.log(path);
    return 0;
  }
  const saved = readJson(path);
  if (!action) {
    for (const k of KEYS) if (saved[k]) console.log(`${k}=${saved[k]}`);
    return 0;
  }
  if (action !== "set" && action !== "unset") throw new Error(`Unknown config action ${action}`);
  if (!KEYS.includes(key as Key)) throw new Error(`Config key must be one of ${KEYS.join(", ")}`);
  if (action === "set") {
    if (!value) throw new Error(`config set ${key} needs a value`);
    // A relative or ~ source would mean something else from the next folder.
    saved[key as Key] = key === "source" ? resolve(value.replace(/^~(?=\/|$)/, homedir())) : value;
  } else {
    delete saved[key as Key];
  }
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, `${JSON.stringify(saved, null, 2)}\n`);
  console.log(`Saved to ${path}`);
  return 0;
}

/** The module paths a command was given, or the configured source folder. */
function modulePaths(rest: string[]): string[] {
  return rest.length > 0 ? rest : [readConfig().source ?? "hosted-modules"];
}

/** Print each module's problems and warnings. Returns how many modules failed. */
function printReports(reports: ModuleReport[], strict: boolean): number {
  let failed = 0;
  for (const report of reports) {
    const fails = report.problems.length > 0 || (strict && report.warnings.length > 0);
    if (fails) failed++;
    const status = fails ? "fail" : report.warnings.length > 0 ? "warn" : "ok";
    console.log(`${status.padEnd(6)}${report.name}${status === "ok" ? "" : ` (${report.configPath})`}`);
    for (const problem of report.problems) console.log(`        - ${problem}`);
    for (const warning of report.warnings) console.log(`        ? ${warning}`);
  }
  return failed;
}

function validate(args: string[]): number {
  const { on, rest } = parseFlags(args, [], ["strict"]);
  const reports = modulePaths(rest).flatMap((path) => validatePath(resolve(path)));
  const failed = printReports(reports, on.has("strict"));
  const total = `${reports.length} module${reports.length === 1 ? "" : "s"}`;
  console.log(failed === 0 ? `\n${total} valid` : `\n${failed} of ${total} invalid`);
  return failed === 0 ? 0 : 1;
}

/** The voice from flags, then config, then the default. Bare numbers get a sign. */
function voiceOf(values: Record<string, string>): Voice {
  const saved = readConfig();
  const signed = (value: string) => (/^\d/.test(value) ? `+${value}` : value);
  const voice: Voice = {
    voice: values.voice ?? saved.voice ?? DEFAULT_VOICE.voice,
    rate: signed(values.rate ?? saved.rate ?? DEFAULT_VOICE.rate),
    pitch: signed(values.pitch ?? saved.pitch ?? DEFAULT_VOICE.pitch),
  };
  if (!/^[+-]\d+%$/.test(voice.rate)) throw new Error(`Rate "${voice.rate}" should look like +10% or -5%`);
  if (!/^[+-]\d+Hz$/.test(voice.pitch)) throw new Error(`Pitch "${voice.pitch}" should look like +2Hz or -5Hz`);
  return voice;
}

const VOICE_FLAGS = ["voice", "rate", "pitch"];

async function tts(args: string[]): Promise<number> {
  const { values, rest } = parseFlags(args, ["out", ...VOICE_FLAGS]);
  const text = rest.join(" ").trim();
  if (!text) throw new Error("Give the text to speak, e.g. simmy tts --out hello.mp3 Hello there");
  if (!values.out) throw new Error("Missing --out <file.mp3>");
  await synthesize(text, resolve(values.out), voiceOf(values));
  console.log(`Wrote ${values.out}`);
  return 0;
}

async function voices(args: string[]): Promise<number> {
  const { rest } = parseFlags(args, []);
  const filter = rest.join(" ").toLowerCase();
  const found = (await listVoices()).filter(
    (v) => !filter || v.ShortName.toLowerCase().includes(filter) || v.FriendlyName.toLowerCase().includes(filter),
  );
  for (const v of found) console.log(`${v.ShortName.padEnd(40)}${v.Gender}`);
  if (found.length === 0) console.log(`No voices match "${filter}"`);
  return 0;
}

const CLIP_STATUS: Record<ClipResult["status"], [string, string]> = {
  // [what happened, what a dry run would do]
  made: ["made", "will make"],
  remade: ["remade", "will remake"],
  current: ["current", "current"],
  kept: ["kept", "kept"],
  remote: ["remote", "remote"],
  failed: ["FAILED", "FAILED"],
};

async function narrate(args: string[]): Promise<number> {
  const { values, on, rest } = parseFlags(args, VOICE_FLAGS, ["force", "dry-run"]);
  const voice = voiceOf(values);
  const dryRun = on.has("dry-run");
  const configs = modulePaths(rest).flatMap((path) => moduleConfigs(resolve(path)));
  let failed = 0;
  let kept = 0;
  for (const configPath of configs) {
    console.log(`${dirname(configPath)} (${voice.voice}, rate ${voice.rate}, pitch ${voice.pitch})`);
    const results = await narrateModule(configPath, { voice, force: on.has("force"), dryRun }, synthesize, (clip) => {
      const status = CLIP_STATUS[clip.status][dryRun ? 1 : 0];
      const line = clip.transcript.length > 60 ? `${clip.transcript.slice(0, 57)}...` : clip.transcript;
      console.log(`  ${clip.audio.padEnd(10)}${status.padEnd(12)}${clip.error ?? line}`);
    });
    if (results.length === 0) console.log("  No autopilot entries have a transcript.");
    failed += results.filter((clip) => clip.status === "failed").length;
    kept += results.filter((clip) => clip.status === "kept").length;
    if (!dryRun) {
      const report = validateModuleFile(configPath);
      if (report.problems.length > 0) printReports([report], false);
    }
  }
  if (kept > 0) console.log(`\n"kept" clips weren't made by simmy narrate. --force replaces them.`);
  return failed === 0 ? 0 : 1;
}

function publish(args: string[]): number {
  const { values, on } = parseFlags(args, ["origin", "name", "source"], ["dry-run"]);
  const options = { ...readConfig(), ...values };
  const origin = options.origin;
  if (!origin) throw new Error("Missing --origin (or \"origin\" in simmy.json)");
  const source = resolve(options.source ?? "hosted-modules");
  if (!existsSync(source)) throw new Error(`No modules folder at ${source}`);
  const name = options.name ?? new URL(origin).hostname.split(".")[0];

  const warned = validatePath(source).filter((report) => report.warnings.length > 0 && report.problems.length === 0);
  if (warned.length > 0) {
    printReports(warned, false);
    console.log("These keys are ignored. Publishing anyway.\n");
  }

  const out = mkdtempSync(join(tmpdir(), "simmy-publish-"));
  const names = publishHostedModules(source, out, origin);
  console.log(`Built ${names.length} module${names.length === 1 ? "" : "s"} for ${origin}`);
  if (on.has("dry-run")) {
    console.log(out);
    return 0;
  }

  const deploy = spawnSync(
    "npx",
    ["--yes", "wrangler", "deploy", "--assets", out, "--name", name, "--compatibility-date", "2026-10-21"],
    { stdio: "inherit" },
  );
  rmSync(out, { recursive: true, force: true });
  return deploy.status ?? 1;
}

const COMMANDS: Record<string, (args: string[]) => number | Promise<number>> = {
  publish,
  validate,
  narrate,
  tts,
  voices,
  config,
};

const [command, ...rest] = process.argv.slice(2);
const run = COMMANDS[command];
if (!run) {
  console.log(USAGE);
  process.exitCode = command && command !== "--help" && command !== "-h" ? 1 : 0;
} else {
  try {
    process.exitCode = await run(rest);
  } catch (error) {
    console.error((error as Error).message);
    process.exitCode = 1;
  }
}
