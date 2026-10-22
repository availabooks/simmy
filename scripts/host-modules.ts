/**
 * Publish hosted-modules for static hosting.
 *
 *   npm run host-modules -- --base https://modules.example.com
 *
 * Writes host-dist/<module>/config.yaml with absolute clip URLs, copies each
 * module's audio folder, and adds a CORS header so other sites can fetch them.
 * The documentation site (docs/) goes at the root, modules.html lists the modules, and
 * skill/ has the Claude skill.
 */
import { pathToFileURL } from "node:url";
import { publishHostedModules } from "../src/cli/host.ts";
import { writeSkill } from "./skill.ts";

export { publishHostedModules };

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
  const names = publishHostedModules(argument("--source") ?? "hosted-modules", out, base, "docs");
  writeSkill(out);
  console.log(`Wrote ${names.length} module${names.length === 1 ? "" : "s"} to ${out}`);
}
