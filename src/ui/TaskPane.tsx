import { useState } from "preact/hooks";
import { addrName, type Inputs, parseRange, rangeAddrs } from "../engine/sheet.ts";
import { doItWrites, type GradeReport, gradeAssignment, type TaskGrade, undoItWrites } from "../module/grade.ts";
import type { SimTask, UnsupportedTool } from "../module/types.ts";
import { useSim } from "./context.ts";
import { type DemoIconName, ICONS } from "./icons.ts";

export function FaIcon({ name, class: className = "" }: { name: DemoIconName; class?: string }) {
  const icon = ICONS[name];
  return (
    <svg
      class={`ee-icon ${className}`}
      viewBox={`0 0 ${icon.width} 512`}
      aria-hidden="true"
      focusable="false"
    >
      <path fill="currentColor" transform="translate(0 448) scale(1 -1)" d={icon.d} />
    </svg>
  );
}

const CELL_REF = /(\b[A-Z]{1,2}\$?\d{1,3}(?::[A-Z]{1,2}\$?\d{1,3})?\b)/;

export type TaskPaneProps = {
  inputs: Inputs;
  onSelect: (ref: string) => void;
  onWrite: (writes: Inputs) => void;
  onClose: () => void;
  /** Tools the demo can't do (video, textbook) report here instead of changing the pane. */
  onUnsupported: (kind: UnsupportedTool, task: SimTask) => void;
};

type View = "tasks" | "submit" | "results";

type Phase = { kind: "working" } | { kind: "grading"; step: number; rule: number } | { kind: "graded" };

/** The Excel Educator task pane: the add-in's task guide, with local grading. */
export function TaskPane({ inputs, onSelect, onWrite, onClose, onUnsupported }: TaskPaneProps) {
  const { module: mod, onGraded } = useSim();
  const tasks = mod.tasks;
  const submitSteps = mod.copy.submitSteps;
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});
  const [finished, setFinished] = useState<Record<string, boolean>>({});
  const [doneIt, setDoneIt] = useState<Record<string, boolean>>({});
  const [busy, setBusy] = useState<string | null>(null);
  const [phase, setPhase] = useState<Phase>({ kind: "working" });
  const [report, setReport] = useState<GradeReport | null>(null);
  const [view, setView] = useState<View>("tasks");

  const toggleFinished = (task: SimTask, index: number) => {
    const nowFinished = !finished[task.id];
    setFinished({ ...finished, [task.id]: nowFinished });
    if (!nowFinished) {
      setCollapsed({ ...collapsed, [task.id]: false });
      return;
    }
    // Finishing a task folds it and opens the next one that is still open, like the add-in.
    const next = tasks.slice(index + 1).find((t) => !finished[t.id]);
    setCollapsed({ ...collapsed, [task.id]: true, ...(next ? { [next.id]: false } : {}) });
    if (next) {
      onSelect(next.hint);
    }
  };

  const runDoIt = (task: SimTask) => {
    if (busy) {
      return;
    }
    setBusy(task.id);
    onSelect(task.range);
    window.setTimeout(() => {
      const undoing = doneIt[task.id];
      onWrite(undoing ? undoItWrites(task) : doItWrites(task));
      setDoneIt((prev) => ({ ...prev, [task.id]: !undoing }));
      setBusy(null);
    }, 450);
  };

  const submit = () => {
    if (phase.kind === "grading") {
      return;
    }
    const result = gradeAssignment(mod, inputs);
    let step = 0;
    let rule = 0;
    setPhase({ kind: "grading", step, rule });
    const tick = () => {
      if (step === 3 && rule < tasks.length) {
        rule++;
        setPhase({ kind: "grading", step, rule });
        window.setTimeout(tick, 140);
        return;
      }
      step++;
      if (step >= submitSteps.length) {
        setReport(result);
        setPhase({ kind: "graded" });
        onGraded?.(result);
        setView("results");
        return;
      }
      setPhase({ kind: "grading", step, rule });
      window.setTimeout(tick, step === 3 ? 140 : 380);
    };
    window.setTimeout(tick, 380);
  };

  // The add-in's precheck: any answer cell still empty counts as an incomplete task.
  const incomplete = tasks.some((task) =>
    rangeAddrs(parseRange(task.range) ?? { c1: 0, r1: 0, c2: 0, r2: 0 }).some(
      (addr) => !inputs[addrName(addr)],
    ),
  );

  const grades = new Map(report?.tasks.map((g) => [g.taskId, g]));

  return (
    <aside class="xd-pane" aria-label="Excel Educator task pane">
      <div class="xd-pane-head">
        <span class="xd-pane-title">{mod.copy.paneTitle}</span>
        <span class="xd-pane-actions">
          <button type="button" class="xd-pane-btn" title="Task pane options" tabIndex={-1}>
            <svg viewBox="0 0 16 16" aria-hidden="true">
              <path d="M4 6l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.2" />
            </svg>
          </button>
          <button type="button" class="xd-pane-btn" title="Close" onClick={onClose}>
            <svg viewBox="0 0 16 16" aria-hidden="true">
              <path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" stroke-width="1.2" />
            </svg>
          </button>
        </span>
      </div>
      <div class="ee-root">
        <div class="ee-top" aria-hidden="true">
          <span class="ee-dots">•••</span>
          <span class="ee-info">i</span>
        </div>
        {view === "results" && report ? (
          <>
            <div class="ee-scroll">
              <h2 class="ee-section">{mod.assignment.section}</h2>
              <div class="ee-header ee-header-flat">
                <div class="ee-qnum">1</div>
                <div class="ee-title">{mod.assignment.title}</div>
                <ScoreBadge earned={report.earned} points={report.points} />
              </div>
              <ResultsPage report={report} onSelect={onSelect} />
            </div>
            <div class="ee-nav ee-nav-end">
              <button
                type="button"
                class="ee-btn ee-btn-back"
                aria-label="Back"
                onClick={() => setView("tasks")}
              >
                <Arrow dir="left" />
              </button>
            </div>
          </>
        ) : view === "tasks" ? (
          <>
            <div class="ee-scroll">
              <h2 class="ee-section">{mod.assignment.section}</h2>
              <div class="ee-box">
                <div class="ee-header">
                  <div class="ee-qnum">1</div>
                  <div class="ee-title">{mod.assignment.title}</div>
                  {report && <ScoreBadge earned={report.earned} points={report.points} />}
                </div>
                <div class="ee-intro">{mod.assignment.text}</div>
                <div class="ee-tasks">
                  {tasks.map((task, index) => {
                    const isCollapsed = collapsed[task.id] ?? true;
                    return (
                      <TaskItem
                        key={task.id}
                        task={task}
                        index={index}
                        collapsed={isCollapsed}
                        finished={finished[task.id] ?? false}
                        doneIt={doneIt[task.id] ?? false}
                        busy={busy === task.id}
                        grade={grades.get(task.id)}
                        onToggle={() => {
                          setCollapsed({ ...collapsed, [task.id]: !isCollapsed });
                          if (isCollapsed) {
                            onSelect(task.hint);
                          }
                        }}
                        onFinish={() => toggleFinished(task, index)}
                        onDoIt={() => runDoIt(task)}
                        onUnsupported={(kind) => onUnsupported(kind, task)}
                        onSelect={onSelect}
                      />
                    );
                  })}
                </div>
              </div>
            </div>
            <div class="ee-nav ee-nav-end">
              <button type="button" class="ee-btn ee-btn-primary" onClick={() => setView("submit")}>
                Next <Arrow dir="right" />
              </button>
            </div>
          </>
        ) : (
          <>
            <div class="ee-scroll">
              <div class="ee-box ee-submit">
                <div class="ee-card-head">Submit</div>
                <div class="ee-submit-body">
                  {incomplete && (
                    <>
                      <p>Please note the following:</p>
                      <div class="ee-warn">
                        <span class="ee-warn-num">1</span>
                        <ul>
                          <li>At least one task on the "{mod.workbook.sheet}" worksheet is incomplete.</li>
                        </ul>
                      </div>
                    </>
                  )}
                  <p class="ee-ready">Ready to Submit?</p>
                  <div class="ee-center">
                    <button
                      type="button"
                      class="ee-btn ee-btn-primary"
                      disabled={phase.kind === "grading"}
                      onClick={submit}
                    >
                      Submit Now
                    </button>
                  </div>
                </div>
              </div>
            </div>
            <div class="ee-nav ee-nav-end">
              <button
                type="button"
                class="ee-btn ee-btn-back"
                aria-label="Back"
                onClick={() => setView("tasks")}
              >
                <Arrow dir="left" />
              </button>
            </div>
          </>
        )}
        {phase.kind === "grading" && (
          <SubmitProgress steps={submitSteps} step={phase.step} rule={phase.rule} total={tasks.length} />
        )}
      </div>
    </aside>
  );
}

function Arrow({ dir }: { dir: "left" | "right" }) {
  return (
    <svg class="ee-arrow" viewBox="0 0 20 14" width="18" height="12" aria-hidden="true">
      <path
        d={dir === "left" ? "M19 7H2M7 1.5L1.5 7 7 12.5" : "M1 7h17M13 1.5L18.5 7 13 12.5"}
        fill="none"
        stroke="currentColor"
        stroke-width="1.5"
      />
    </svg>
  );
}

function ScoreBadge({ earned, points }: { earned: number; points: number }) {
  const variant = earned >= points ? "success" : earned > 0 ? "warning" : "danger";
  return (
    <span class={`ee-badge ee-badge-${variant}`}>
      {earned} / {points}
    </span>
  );
}

type TaskItemProps = {
  task: SimTask;
  index: number;
  collapsed: boolean;
  finished: boolean;
  doneIt: boolean;
  busy: boolean;
  grade: TaskGrade | undefined;
  onToggle: () => void;
  onFinish: () => void;
  onDoIt: () => void;
  onUnsupported: (kind: UnsupportedTool) => void;
  onSelect: (ref: string) => void;
};

function TaskItem(props: TaskItemProps) {
  const { task, index, collapsed, finished, doneIt, busy, grade } = props;
  return (
    <div class={`ee-task${collapsed ? " is-collapsed" : ""}`} data-task={task.id}>
      <div
        class="ee-task-inner"
        onClick={(event) => {
          const target = event.target as HTMLElement;
          if (target.closest(".no-trigger-collapse")) {
            return;
          }
          props.onToggle();
        }}
      >
        <div class="ee-task-text">
          <FaIcon name={collapsed ? "angleDown" : "angleUp"} class="ee-caret" />
          <button
            type="button"
            class={`ee-pill no-trigger-collapse ${finished ? "is-finished" : ""}`}
            title="Mark as finished"
            onClick={props.onFinish}
          >
            1.{index + 1}
          </button>{" "}
          <CellText text={task.text} onSelect={props.onSelect} />{" "}
          <ol type="a" class="ee-details">
            {task.details.map((line) => (
              <li key={line}>
                <CellText text={line} onSelect={props.onSelect} />
              </li>
            ))}
          </ol>
        </div>
        <div class="ee-tools no-trigger-collapse">
          <button
            type="button"
            class="ee-tool"
            title={doneIt ? "Clear this task so I can do it." : "Complete this task for me"}
            aria-label={doneIt ? "Undo it" : "Do it"}
            data-pilot="doit"
            onClick={props.onDoIt}
          >
            <FaIcon
              name={busy ? "spinnerThird" : doneIt ? "trashUndo" : "magic"}
              class={busy ? "ee-spin" : ""}
            />
          </button>
          <button
            type="button"
            class="ee-tool"
            title="Short video solution"
            onClick={() => props.onUnsupported("video")}
          >
            <FaIcon name="tvRetro" />
          </button>
          <button
            type="button"
            class="ee-tool"
            title="Open the textbook"
            onClick={() => props.onUnsupported("book")}
          >
            <FaIcon name="book" />
          </button>
          <Points task={task} grade={grade} />
        </div>
      </div>
    </div>
  );
}

function Points({ task, grade }: { task: SimTask; grade: TaskGrade | undefined }) {
  const title = grade
    ? grade.earned === grade.points
      ? `You earned ${grade.points === 1 ? "the point" : `all ${grade.points} points`} for this task`
      : `You earned ${grade.earned} of ${grade.points} ${grade.points === 1 ? "point" : "points"} for this task`
    : `This task is worth ${task.points} ${task.points === 1 ? "point" : "points"}`;
  return (
    <span class="ee-points" title={title}>
      {grade ? `${grade.earned}/${grade.points}` : task.points}
    </span>
  );
}

const PRAISE = ["Outstanding", "Excellent", "Fabulous"];

/** The add-in's results page: one colored card per task, with the grader's deductions. */
function ResultsPage({ report, onSelect }: { report: GradeReport; onSelect: (ref: string) => void }) {
  const tasks = useSim().module.tasks;
  return (
    <div class="ee-results">
      {tasks.map((task, index) => {
        const grade = report.tasks.find((g) => g.taskId === task.id);
        if (!grade) {
          return null;
        }
        const tone = grade.earned === grade.points ? "correct" : grade.earned === 0 ? "incorrect" : "partial";
        const lost = grade.points - grade.earned;
        return (
          <div class="ee-task-card" key={task.id}>
            <div class={`ee-task-card-header ${tone}`}>
              <div class="ee-task-card-title">Task {index + 1}</div>
              <div>
                {grade.earned}/{grade.points} {grade.points === 1 ? "point" : "points"}
              </div>
            </div>
            <div class="ee-task-card-text">
              <div>
                <CellText text={task.text} onSelect={onSelect} />
                {lost > 0 && (
                  <ol type="a" class="ee-details">
                    {task.details.map((line) => (
                      <li key={line}>
                        <CellText text={line} onSelect={onSelect} />
                      </li>
                    ))}
                  </ol>
                )}
              </div>
              {lost === 0 ? (
                <div class="ee-praise">
                  {PRAISE[index % PRAISE.length]}! You earned full credit for this task.
                </div>
              ) : (
                <ul class="ee-deductions">
                  {grade.deductions.map((deduction) => (
                    <li key={deduction.message}>
                      <div class="ee-deduction-points">[-{deduction.points}]</div>
                      <div>
                        <CellText text={deduction.message} onSelect={onSelect} />
                      </div>
                    </li>
                  ))}
                </ul>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/** Task text with cell references turned into the add-in's select-this-range buttons. */
function CellText({ text, onSelect }: { text: string; onSelect: (ref: string) => void }) {
  const parts = text.split(CELL_REF);
  return (
    <>
      {parts.map((part, i) =>
        i % 2 === 1 ? (
          <button
            key={i}
            type="button"
            class="ee-range-link no-trigger-collapse"
            title={`Select ${part} on this sheet`}
            onClick={() => onSelect(part)}
          >
            {part}
          </button>
        ) : (
          part
        ),
      )}
    </>
  );
}

function SubmitProgress({
  steps,
  step,
  rule,
  total,
}: {
  steps: string[];
  step: number;
  rule: number;
  total: number;
}) {
  return (
    <div class="ee-progress-veil">
      <div class="ee-card">
        <div class="ee-card-head">Submit Assignment</div>
        <ul class="ee-steps">
          {steps.map((label, i) => (
            <li key={label} class={i < step ? "is-done" : i === step ? "is-active" : ""}>
              <FaIcon
                name={i < step ? "checkCircle" : i === step ? "spinnerThird" : "circle"}
                class={i === step ? "ee-spin" : ""}
              />
              <span>{label}</span>
              {i === 3 && i === step && (
                <div class="ee-bar">
                  <div class="ee-bar-fill" style={{ width: `${(rule / total) * 100}%` }}>
                    {rule} / {total}
                  </div>
                </div>
              )}
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}
