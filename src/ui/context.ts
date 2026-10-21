import { createContext } from "preact";
import { useContext } from "preact/hooks";
import type { GradeReport } from "../module/grade.ts";
import type { SimModule } from "../module/types.ts";

export type SimContextValue = {
  module: SimModule;
  /** Called when the student submits and the grader finishes. */
  onGraded?: (report: GradeReport) => void;
};

export const SimContext = createContext<SimContextValue | null>(null);

export function useSim(): SimContextValue {
  const value = useContext(SimContext);
  if (!value) {
    throw new Error("Simulator components must render inside <SimContext.Provider>.");
  }
  return value;
}
