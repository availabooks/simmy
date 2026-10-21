/**
 * Narration clips live in an `audio` folder next to the module file.
 * A relative name such as `1.mp3` means that file. A URL or a path that
 * starts with `/` is already a location, and is left alone.
 */
import type { SimModule } from "./types.ts";

const SCHEME = /^[a-z][a-z\d+.-]*:/i;

/** A clip name that is resolved from the module's audio folder. */
export function isRelativeAudio(audio: string): boolean {
  return audio !== "" && !audio.startsWith("/") && !SCHEME.test(audio);
}

/**
 * Point a clip at the audio folder beside the module file.
 * `moduleUrl` is the URL of the config itself, not of the folder.
 */
export function resolveAudioUrl(audio: string, moduleUrl: string | URL): string {
  if (!isRelativeAudio(audio)) return audio;
  const folder = new URL("audio/", moduleUrl);
  const url = new URL(audio, folder);
  if (url.origin !== folder.origin || !url.pathname.startsWith(folder.pathname)) {
    throw new Error(`Audio path "${audio}" must stay inside the module audio folder`);
  }
  return url.href;
}

/** Resolve every relative clip on a module loaded from `moduleUrl`. */
export function resolveModuleAudio(mod: SimModule, moduleUrl: string | URL): SimModule {
  return {
    ...mod,
    autopilot: mod.autopilot.map((step) =>
      step.audio === undefined ? step : { ...step, audio: resolveAudioUrl(step.audio, moduleUrl) },
    ),
  };
}
