/**
 * `simmy narrate`: speak a module's autopilot transcripts into its audio folder. An entry with
 * a transcript and no clip gets the next numbered clip name, written into config.yaml as a text
 * edit so the author's comments and layout stay as they were.
 *
 * narration.json, beside config.yaml, records what each clip was made from. A clip is made
 * again when its transcript or voice changes. A clip it has no record of was made some other
 * way (recorded by hand, say), so it is kept unless `force` is set.
 */
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, relative, resolve, sep } from "node:path";
import { isMap, isScalar, isSeq, parseDocument, type Pair, type YAMLMap } from "yaml";
import { isRelativeAudio } from "../module/audio.ts";
import type { Synthesize, Voice } from "./tts.ts";

export interface NarrateOptions {
  voice: Voice;
  /** Make every clip again, including ones narration.json has no record of. */
  force?: boolean;
  /** Report what would happen without writing anything. */
  dryRun?: boolean;
}

/** What happened to one autopilot entry. */
export interface ClipResult {
  entry: number;
  audio: string;
  transcript: string;
  status: "made" | "remade" | "current" | "kept" | "remote" | "failed";
  error?: string;
}

type Made = Voice & { transcript: string };

const scalarValue = (map: YAMLMap, key: string): unknown => {
  const node = map.get(key, true);
  return isScalar(node) ? node.value : undefined;
};

const same = (a: Made | undefined, b: Made) =>
  a !== undefined && a.transcript === b.transcript && a.voice === b.voice && a.rate === b.rate && a.pitch === b.pitch;

export async function narrateModule(
  configPath: string,
  options: NarrateOptions,
  synthesize: Synthesize,
  /** Called as each entry is handled, for progress output. */
  onClip: (result: ClipResult) => void = () => {},
): Promise<ClipResult[]> {
  const dir = dirname(resolve(configPath));
  const audioDir = join(dir, "audio");
  const manifestPath = join(dir, "narration.json");
  const text = readFileSync(configPath, "utf8");
  const doc = parseDocument(text);
  if (doc.errors.length > 0) throw new Error(`${configPath}: ${doc.errors[0].message}`);
  const autopilot = doc.get("autopilot", true);
  if (!isSeq(autopilot)) throw new Error(`${configPath} has no autopilot list to narrate`);

  const manifest: { [name: string]: Made } = existsSync(manifestPath)
    ? JSON.parse(readFileSync(manifestPath, "utf8"))
    : {};

  // New clips are numbered after every name already in use, so none is reused.
  const used = new Set<string>([
    ...Object.keys(manifest),
    ...(existsSync(audioDir) ? readdirSync(audioDir) : []),
    ...autopilot.items.flatMap((item) => (isMap(item) && typeof scalarValue(item, "audio") === "string" ? [scalarValue(item, "audio") as string] : [])),
  ]);
  let next = Math.max(0, ...[...used].map((name) => Number(/^(\d+)\.mp3$/.exec(name)?.[1] ?? 0))) + 1;

  /** Text inserted into config.yaml, applied last-first so the offsets stay right. */
  const edits: { at: number; insert: string }[] = [];
  const jobs: { result: ClipResult; file: string }[] = [];
  const results: ClipResult[] = [];

  autopilot.items.forEach((item, entry) => {
    if (!isMap(item)) return;
    const transcript = scalarValue(item, "transcript");
    if (typeof transcript !== "string" || transcript.trim() === "") return;
    let audio = scalarValue(item, "audio");
    if (typeof audio !== "string") {
      audio = `${next++}.mp3`;
      const pair = item.items.find((p) => isScalar(p.key) && p.key.value === "transcript") as Pair<{ range?: number[] }>;
      const at = pair.key!.range![0];
      const column = at - (text.lastIndexOf("\n", at - 1) + 1);
      edits.push({ at, insert: item.flow ? `audio: ${audio}, ` : `audio: ${audio}\n${" ".repeat(column)}` });
    }
    const result: ClipResult = { entry, audio: audio as string, transcript, status: "current" };
    results.push(result);
    if (!isRelativeAudio(result.audio)) {
      result.status = "remote";
      return;
    }
    const file = resolve(audioDir, result.audio);
    const fromAudio = relative(audioDir, file);
    if (fromAudio === ".." || fromAudio.startsWith(`..${sep}`)) {
      result.status = "failed";
      result.error = "leaves the audio folder";
      return;
    }
    const wanted: Made = { transcript, ...options.voice };
    if (!existsSync(file)) result.status = "made";
    else if (options.force || (manifest[result.audio] && !same(manifest[result.audio], wanted))) result.status = "remade";
    else if (!manifest[result.audio]) result.status = "kept";
    if (result.status === "made" || result.status === "remade") jobs.push({ result, file });
  });

  if (!options.dryRun && edits.length > 0) {
    let updated = text;
    for (const { at, insert } of edits.sort((a, b) => b.at - a.at)) {
      updated = updated.slice(0, at) + insert + updated.slice(at);
    }
    writeFileSync(configPath, updated);
  }

  for (const result of results) {
    const job = jobs.find((j) => j.result === result);
    if (job && !options.dryRun) {
      try {
        await synthesize(result.transcript, job.file, options.voice);
        manifest[result.audio] = { transcript: result.transcript, ...options.voice };
        // Saved after every clip, so an interrupted run picks up where it stopped.
        writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
      } catch (error) {
        result.status = "failed";
        result.error = (error as Error).message;
      }
    }
    onClip(result);
  }
  return results;
}
