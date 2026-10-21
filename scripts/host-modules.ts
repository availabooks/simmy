/**
 * Publish hosted-modules for static hosting.
 *
 *   npm run host-modules -- --base https://modules.example.com
 *
 * Writes host-dist/<module>/config.yaml with absolute clip URLs, copies each
 * module's audio folder, and adds a CORS header so other sites can fetch them.
 */
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { join, relative, resolve, sep } from "node:path";
import { pathToFileURL } from "node:url";
import { relativeAudioRefs, rewriteHostedModule } from "../src/module/host.ts";

const HEADERS = `/*
  Access-Control-Allow-Origin: *
`;

function escapeHtml(value: string): string {
  return value.replace(/[&<>"]/g, (char) => {
    switch (char) {
      case "&":
        return "&amp;";
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      default:
        return "&quot;";
    }
  });
}

/** A plain index so the site root lists the published modules. */
function indexPage(names: string[]): string {
  const items = names
    .map((name) => {
      const href = `/${encodeURIComponent(name)}/config.yaml`;
      return `    <li><a href="${escapeHtml(href)}">${escapeHtml(name)}</a></li>`;
    })
    .join("\n");
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>simmy modules</title>
  </head>
  <body>
    <h1>simmy modules</h1>
    <ul>
${items}
    </ul>
  </body>
</html>
`;
}

export function publishHostedModules(sourceDir: string, outDir: string, base: string): string[] {
  const root = new URL(base.endsWith("/") ? base : `${base}/`);
  const names = readdirSync(sourceDir).filter((name) => statSync(join(sourceDir, name)).isDirectory());
  if (names.length === 0) {
    throw new Error(`No modules in ${sourceDir}`);
  }

  rmSync(outDir, { recursive: true, force: true });
  mkdirSync(outDir, { recursive: true });
  writeFileSync(join(outDir, "_headers"), HEADERS);
  writeFileSync(join(outDir, "index.html"), indexPage(names));

  for (const name of names) {
    const dir = join(sourceDir, name);
    const configPath = join(dir, "config.yaml");
    if (!existsSync(configPath)) {
      throw new Error(`${name} is missing config.yaml`);
    }
    const text = readFileSync(configPath, "utf8");
    const audioDir = resolve(dir, "audio");
    for (const ref of relativeAudioRefs(text)) {
      const file = resolve(audioDir, ref);
      const fromAudio = relative(audioDir, file);
      if (fromAudio === ".." || fromAudio.startsWith(`..${sep}`)) {
        throw new Error(`${name}: audio path "${ref}" leaves the audio folder`);
      }
      if (!existsSync(file)) {
        throw new Error(`${name}: missing audio/${ref}`);
      }
    }
    const dest = join(outDir, name);
    mkdirSync(dest, { recursive: true });
    const moduleUrl = new URL(`${encodeURIComponent(name)}/config.yaml`, root);
    writeFileSync(join(dest, "config.yaml"), rewriteHostedModule(text, moduleUrl));
    if (existsSync(audioDir)) {
      cpSync(audioDir, join(dest, "audio"), { recursive: true });
    }
  }
  return names;
}

function argument(flag: string): string | undefined {
  const index = process.argv.indexOf(flag);
  return index === -1 ? undefined : process.argv[index + 1];
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const base = argument("--base");
  if (!base) {
    console.error("Usage: npm run host-modules -- --base https://modules.example.com [--out host-dist]");
    process.exit(1);
  }
  const out = argument("--out") ?? "host-dist";
  const names = publishHostedModules(argument("--source") ?? "hosted-modules", out, base);
  console.log(`Wrote ${names.length} module${names.length === 1 ? "" : "s"} to ${out}`);
}
