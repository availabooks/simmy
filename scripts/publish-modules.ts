/**
 * Build host-dist (the docs site at the root, the modules with absolute clip URLs and the
 * Claude skill beside it),
 * then deploy it as Workers static assets, and refresh jsDelivr's copy of @latest for the docs.
 *
 * The origin is modules-site/package.json `simmyOrigin`. It has to be the public
 * URL of the simmy-modules Worker, because those URLs are written into the YAML.
 */
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { publishHostedModules } from "./host-modules.ts";
import { writeSkill } from "./skill.ts";

const site = JSON.parse(readFileSync("modules-site/package.json", "utf8")) as { simmyOrigin?: string };
const origin = site.simmyOrigin;
if (!origin) {
  throw new Error("modules-site/package.json is missing simmyOrigin");
}

const names = publishHostedModules("hosted-modules", "host-dist", origin, "docs");
writeSkill("host-dist");
console.log(`Wrote ${names.length} module${names.length === 1 ? "" : "s"} for ${origin}`);

const deploy = spawnSync("npm", ["exec", "--", "wrangler", "deploy"], {
  cwd: "modules-site",
  stdio: "inherit",
});
if (deploy.status !== 0) {
  process.exit(deploy.status ?? 1);
}

// The docs load @latest from jsDelivr, which caches what @latest means. Purge it so the site
// shows the release just published rather than the one jsDelivr last saw.
const purge = await fetch("https://purge.jsdelivr.net/npm/@availabooks/simmy@latest/dist/simmy.js");
console.log(purge.ok ? "Refreshed jsDelivr's @availabooks/simmy@latest" : `jsDelivr purge failed: HTTP ${purge.status}`);
