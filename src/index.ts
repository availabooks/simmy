/**
 * @availabooks/simmy: an embeddable Excel simulator with a graded task pane.
 * Importing this module registers <simmy-app>.
 */
import { defineSimmy } from "./element.ts";

export { defineSimmy, SimmyElement } from "./element.ts";
export * from "./engine/sheet.ts";
export {
  type Deduction,
  doItWrites,
  type GradeReport,
  gradeAssignment,
  gradeTask,
  initialInputs,
  solvedInputs,
  type TaskGrade,
  undoItWrites,
} from "./module/grade.ts";
export { DEFAULT_COPY, defaultAutopilot, loadModule, ModuleError, parseModule } from "./module/parse.ts";
export type * from "./module/types.ts";
export { SimContext, type SimContextValue } from "./ui/context.ts";
export { detectPlatform, Simmy, type SimmyMode, type SimmyProps } from "./ui/Simmy.tsx";

defineSimmy();
