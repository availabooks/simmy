// The iframe page ships next to the standalone build, so it loads ./simmy.js.
import { copyFileSync, mkdirSync } from "node:fs";

mkdirSync("dist", { recursive: true });
copyFileSync("embed/index.html", "dist/embed.html");
