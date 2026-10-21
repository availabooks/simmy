/**
 * The copy published for static hosting carries absolute clip URLs, so a page
 * on another origin can load the module and play its audio without knowing
 * about the audio folder.
 */
import { isScalar, parseDocument, visit } from "yaml";
import type { Scalar } from "yaml";
import { isRelativeAudio, resolveAudioUrl } from "./audio.ts";

function audioScalar(node: unknown): Scalar | undefined {
  return isScalar(node) && typeof node.value === "string" ? node : undefined;
}

/** Relative clip names in a module file, in source order. */
export function relativeAudioRefs(text: string): string[] {
  const refs: string[] = [];
  visit(parseDocument(text), {
    Pair(_key, pair) {
      const name = audioScalar(pair.key);
      const clip = audioScalar(pair.value);
      if (name?.value === "audio" && clip && isRelativeAudio(clip.value as string)) {
        refs.push(clip.value as string);
      }
    },
  });
  return refs;
}

/** Rewrite relative clip names to absolute URLs under the published module URL. */
export function rewriteHostedModule(text: string, moduleUrl: string | URL): string {
  const doc = parseDocument(text);
  visit(doc, {
    Pair(_key, pair) {
      const name = audioScalar(pair.key);
      const clip = audioScalar(pair.value);
      if (!name || name.value !== "audio" || !clip || !isRelativeAudio(clip.value as string)) return;
      clip.value = resolveAudioUrl(clip.value as string, moduleUrl);
    },
  });
  return doc.toString();
}
