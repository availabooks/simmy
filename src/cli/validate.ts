/**
 * Checks module folders (<dir>/config.yaml and <dir>/audio/) the way the simulator would load
 * them, plus what only shows up once hosted: clip files that are missing or outside the audio
 * folder, and solutions that fail their own rubric. Used by `simmy validate`, and by publishing so a broken module never goes live.
 */
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { basename, dirname, join, relative, resolve, sep } from "node:path";
import { parse } from "yaml";
import { gradeAssignment, solvedInputs } from "../module/grade.ts";
import { relativeAudioRefs } from "../module/host.ts";
import { unknownKeys } from "../module/keys.ts";
import { ModuleError, parseModule } from "../module/parse.ts";

export interface ModuleReport {
  /** The module's folder name. */
  name: string;
  configPath: string;
  /** Every problem found, each prefixed with the path to the field. Empty when the module is fine. */
  problems: string[];
  /** Keys simmy ignores, most likely misspelled. The module still loads. */
  warnings: string[];
}

/** Check one module from its config file. Its clips are looked up in the audio folder beside it. */
export function validateModuleFile(configPath: string): ModuleReport {
  const name = basename(dirname(resolve(configPath)));
  const report: ModuleReport = { name, configPath, problems: [], warnings: [] };
  if (!existsSync(configPath)) {
    report.problems.push("missing config.yaml");
    return report;
  }
  const text = readFileSync(configPath, "utf8");
  let data: unknown;
  try {
    data = parse(text);
  } catch (error) {
    // Nothing past this point can be trusted without a parse.
    report.problems.push(`YAML: ${(error as Error).message}`);
    return report;
  }
  report.warnings.push(...unknownKeys(data));
  try {
    const mod = parseModule(data);
    // Do it writes each solution, and a module whose own answer loses points can't be passed.
    gradeAssignment(mod, solvedInputs(mod)).tasks.forEach((task, i) => {
      for (const { points, message } of task.deductions) {
        report.problems.push(`tasks[${i}].solution: Do it loses ${points} point${points === 1 ? "" : "s"} (${message})`);
      }
    });
  } catch (error) {
    if (!(error instanceof ModuleError)) throw error;
    report.problems.push(...error.problems);
  }

  const audioDir = resolve(dirname(configPath), "audio");
  for (const ref of relativeAudioRefs(text)) {
    const fromAudio = relative(audioDir, resolve(audioDir, ref));
    if (fromAudio === ".." || fromAudio.startsWith(`..${sep}`)) {
      report.problems.push(`audio "${ref}": leaves the audio folder`);
    } else if (!existsSync(resolve(audioDir, ref))) {
      report.problems.push(`audio "${ref}": missing audio/${ref}`);
    }
  }
  return report;
}

/**
 * The config files a path names: a config file, a module folder (it has a config.yaml), or a
 * folder of module folders.
 */
export function moduleConfigs(path: string): string[] {
  if (!existsSync(path)) throw new Error(`Nothing at ${path}`);
  if (!statSync(path).isDirectory()) return [path];
  if (existsSync(join(path, "config.yaml"))) return [join(path, "config.yaml")];
  const names = readdirSync(path).filter((name) => statSync(join(path, name)).isDirectory());
  if (names.length === 0) throw new Error(`No modules in ${path}`);
  return names.map((name) => join(path, name, "config.yaml"));
}

/** Check every module a path names (see `moduleConfigs`). */
export function validatePath(path: string): ModuleReport[] {
  return moduleConfigs(path).map(validateModuleFile);
}
