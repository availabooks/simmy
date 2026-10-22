/**
 * Speech from Microsoft Edge's read-aloud voices (the service behind Edge's "Read aloud"), the
 * same voices and format (24 kHz, 48 kbps mono MP3) as the hosted modules' clips. It needs no
 * account or key, but it is an unofficial service, so a failure here may mean it changed.
 */
import { mkdirSync } from "node:fs";
import { dirname } from "node:path";
import { CHROMIUM_FULL_VERSION, generateSecMsGecToken, TRUSTED_CLIENT_TOKEN } from "node-edge-tts/dist/drm.js";
import { EdgeTTS } from "node-edge-tts";

export interface Voice {
  /** e.g. en-US-AvaMultilingualNeural */
  voice: string;
  /** Speaking rate relative to normal, e.g. +10% or -5%. */
  rate: string;
  /** Pitch relative to normal, e.g. +2Hz or -10Hz. */
  pitch: string;
}

export const DEFAULT_VOICE: Voice = { voice: "en-US-AvaMultilingualNeural", rate: "+0%", pitch: "+0Hz" };

/** Writes `text` spoken by `voice` to an MP3 at `path`. Replaceable in tests. */
export type Synthesize = (text: string, path: string, voice: Voice) => Promise<void>;

export const synthesize: Synthesize = async (text, path, { voice, rate, pitch }) => {
  mkdirSync(dirname(path), { recursive: true });
  const lang = voice.split("-").slice(0, 2).join("-");
  const tts = new EdgeTTS({ voice, lang, rate, pitch, outputFormat: "audio-24khz-48kbitrate-mono-mp3", timeout: 30000 });
  try {
    await tts.ttsPromise(text, path);
  } catch {
    // The service drops a connection now and then; one retry clears most of those.
    await tts.ttsPromise(text, path);
  }
};

export interface VoiceInfo {
  ShortName: string;
  Gender: string;
  Locale: string;
  FriendlyName: string;
}

export async function listVoices(): Promise<VoiceInfo[]> {
  const major = CHROMIUM_FULL_VERSION.split(".")[0];
  const url =
    "https://speech.platform.bing.com/consumer/speech/synthesize/readaloud/voices/list" +
    `?trustedclienttoken=${TRUSTED_CLIENT_TOKEN}&Sec-MS-GEC=${generateSecMsGecToken()}` +
    `&Sec-MS-GEC-Version=1-${CHROMIUM_FULL_VERSION}`;
  const response = await fetch(url, {
    headers: {
      "User-Agent": `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${major}.0.0.0 Safari/537.36 Edg/${major}.0.0.0`,
    },
  });
  if (!response.ok) throw new Error(`Listing voices failed: HTTP ${response.status}`);
  return (await response.json()) as VoiceInfo[];
}
