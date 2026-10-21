import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { parse } from "yaml";
import { parseModule } from "../src/module/parse.ts";
import type { SimModule } from "../src/module/types.ts";

export function moduleText(name = "references-and-calculations"): string {
  return readFileSync(resolve("modules", `${name}.yaml`), "utf8");
}

/** A module from modules/, parsed the way the element parses it. */
export function loadFixture(name?: string): SimModule {
  return parseModule(parse(moduleText(name)));
}
