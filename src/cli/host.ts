/**
 * Write a folder of modules for static hosting: <out>/<module>/config.yaml with absolute
 * clip URLs, each module's audio folder beside it, and a CORS header so other sites can
 * fetch them. Shared by `simmy publish` and the repo's own host-modules script.
 */
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { rewriteHostedModule } from "../module/host.ts";
import { validateModuleFile } from "./validate.ts";

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

/** A plain list of the published modules, linked from the docs. */
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

export function publishHostedModules(
  sourceDir: string,
  outDir: string,
  base: string,
  /** A folder of static files to serve at the root, such as the docs site. */
  siteDir?: string,
): string[] {
  const root = new URL(base.endsWith("/") ? base : `${base}/`);
  const names = readdirSync(sourceDir).filter((name) => statSync(join(sourceDir, name)).isDirectory());
  if (names.length === 0) {
    throw new Error(`No modules in ${sourceDir}`);
  }
  // Check every module before writing anything, and report all of them at once.
  const problems = names.flatMap((name) =>
    validateModuleFile(join(sourceDir, name, "config.yaml")).problems.map((problem) => `${name}: ${problem}`),
  );
  if (problems.length > 0) {
    throw new Error(`Invalid modules:\n${problems.map((p) => `  - ${p}`).join("\n")}`);
  }

  rmSync(outDir, { recursive: true, force: true });
  mkdirSync(outDir, { recursive: true });
  if (siteDir) {
    for (const entry of readdirSync(siteDir)) {
      if (names.includes(entry)) {
        throw new Error(`${siteDir}/${entry} has the same name as a module`);
      }
      cpSync(join(siteDir, entry), join(outDir, entry), { recursive: true });
    }
  }
  writeFileSync(join(outDir, "_headers"), HEADERS);
  writeFileSync(join(outDir, siteDir ? "modules.html" : "index.html"), indexPage(names));

  for (const name of names) {
    const dir = join(sourceDir, name);
    const text = readFileSync(join(dir, "config.yaml"), "utf8");
    const audioDir = resolve(dir, "audio");
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
