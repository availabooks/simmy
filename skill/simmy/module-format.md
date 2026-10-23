# Module format

A module is one YAML file (JSON works too). Only `workbook.name`, `assignment.title`,
`assignment.text` and at least one task are required. Everything else has a default.

```yaml
version: 1                       # optional; only 1 exists
workbook:
  name: Book Order               # required: the workbook title in Excel's title bar
  sheet: Order                   # the sheet tab; default Sheet1
assignment:                      # the task pane's assignment heading
  title: Book Order Totals       # required
  text: What the student is asked to do overall.   # required
  section: Assessment Section 1  # default Assessment Section 1
sheet:
  rows: 40                       # default 40
  columns: 20                    # default 20
  columnWidth: 64                # px, default 64
  widths: { A: 20, B: 140 }      # px per column
  start: B2                      # the cell selected at first; default: the first task's hint
  cells: { B2: Shipping, C2: 4.5, E9: "=SUM(E5:E7)" }
  styles: [...]                  # see below
  outlines: [...]                # see below
tasks: [...]                     # required; see below
note: { title: Try it, steps: [First step., Second step.] }   # the "try it" card
autopilot: [...]                 # the demo; default: type each solution and fill its range
poster: { range: B4:E9 }         # the still image phones show; default: the top-left of the content
copy:                            # wording overrides, rarely needed
  paneTitle: Excel Educator
  unsupported:                   # what the pane's video and textbook buttons say
    video: { title: …, text: … }
    book: { title: …, text: … }
  submitSteps: [Validate Key, Prepare Rubric, Check for Completeness, Evaluate Tasks, Finalize Score]  # exactly 5
```

Every cell address and range must fall inside `rows` × `columns`.

## Cells

`sheet.cells` maps an address to text, a number, `true`/`false`, or a formula starting with `=`.
Leave out cells the student fills. Numbers are stored as numbers: write `0.3`, not `30%`, and
use a `format` style to show a percentage.

## Styles and outlines

`styles` is a list, applied in order, so a later entry wins. Each has a `range` (one range, several
separated by spaces or commas, or a list) and any of:

| Key | Value |
|---|---|
| `fill` | Background color, `"#rrggbb"` (quoted) |
| `color` | Text color |
| `bold`, `italic` | `true` or `false` |
| `lined` | `true`: thin gridlines inside a filled range stay visible |
| `align` | `center` (the only option) |
| `format` | `General`, `"0"`, `"0.00"`, `"$#,##0.00"`, `"0.00%"` or `accounting` |

Any other style key is an error.

`outlines` draws a border around ranges: `{ range: B4:E7, weight: 1, color: "#a6a6a6" }`, where
`weight` is 1 or 2, and 1 and `#a6a6a6` are the defaults.

## Tasks

```yaml
- id: A                  # default A, B, C… by position; ids must be unique
  name: Line cost        # the pane's tab label; default: the id
  text: In E5, calculate the cost, then copy it down through E7.   # required
  details: [Cost is price times quantity.]                         # extra bullet points
  range: E5:E7           # where Do it writes the solution; given together with solution
  hint: E5               # the cell selected when the task opens; default: the range's top-left
  solution: =C5*D5       # the formula for the range's top-left cell
  names: { Rate: G2 }    # names Do it defines (and Undo it deletes)
  points: 10             # default: the rules' points added up
  rules: [...]
```

A task needs `range` and `solution`, or `names`, or both. Do it writes `solution` into every
cell of `range`, shifting relative references as a fill would, so write it for the top-left
cell, and defines the task's `names`. `simmy validate` grades that result and reports any rule
it fails.

A formula that spills (SORT, UNIQUE, SEQUENCE, or a range such as `=A2:A9*2`) goes in one cell:
make `range` that one cell, and grade the cells it fills with a `value` rule's `range`.

## Names

`workbook.names` defines names from the start, and a task's `names` are what its Do it defines:
`{ Rate: G2, Sales: D2:D9 }`. A name follows Excel's rules: a letter, `_` or `\` first, then
letters, numbers, `.` and `_`; no spaces; nothing that reads as a cell reference (`Q1`, `R1C1`);
not `C` or `R`; at most 255 characters. Names match without regard to case. Every name a
formula uses must be defined in one of those places, and `simmy validate` says when one isn't.

## Rules

Every rule has `kind`, `points` and `message`. The message is shown to the student when the rule
fails, so write it as advice: "E5 should use the price in C5."

| Kind | Keys | Passes when |
|---|---|---|
| `value` | `cell` or `range`, `formula` | The cell's value equals what `formula` gives on the student's own sheet. Write `formula` with `$` on every reference (or plain cells), since it is evaluated away from the cell. Numbers match within rounding error. With `range`, `formula`'s array must fill the range exactly: use it for a result that spills. |
| `reference` | `cell`, `ref`, optional `forms` | `cell` holds a formula that refers to `ref`, alone or inside a range (`SUM(C2:C5)` refers to C3), or through a name for it. With `forms`, a reference must be spelled exactly as one of them: `forms: [$C$2, C$2]` requires the row locked. A typed number never passes. |
| `copied` | `range` | Every cell in `range` holds the first cell's formula, shifted as a fill would shift it. |
| `equals` | `cell`, `value`, optional `tolerance`, `trim` | The cell holds `value` exactly (text, a number, true/false, or an error such as `"#N/A"`), or any of a list of them. A number may be off by up to `tolerance`; with `trim: true`, text ignores every space. |
| `contains` | `cell` or `range`, `text` | Every cell holds a formula that includes `text`, ignoring case: `text: "TEXTJOIN("` checks the function was used. A typed value never passes. |
| `name` | `name`, optional `range` | The name is defined (any case). With `range`, it is spelled exactly so and refers to that range. |
| `between` | `cell` or `range`, `lower`, `upper`, optional `endpoints` | Every cell is a number from `lower` to `upper` (each a number or a formula), both included unless `endpoints: exclude`. Use it with `contains` to grade RANDBETWEEN, whose numbers change. |
| `selected` | `range` | The student ended on this range: by Go To, the Name Box, or a click. At most one per module. |

Each kind mirrors a grading step of the Excel Educator add-in: `value` is values_match,
`reference` find_references, `copied` cells_copied, `equals` value_correct (and
value_correct_tolerance), `contains` formula_contains, `name` check_name_exists and
check_named_range, `between` value_between, and `selected` activecell.

A task scores its `points` minus the points of the rules it fails, never below 0. To give full
credit only for a right formula, put most points on `value`. To require that inputs are
referenced and not typed, add a `reference` rule per input cell.

## Autopilot

The autopilot plays the assignment as a demo. Without `autopilot`, it types each task's solution
at the top-left of its range and fills the rest. A script is a list of steps:

| Step | Does |
|---|---|
| `{ task: A }` | Opens task A in the pane. |
| `{ task: A, at: D6 }` | Types A's solution in D6. |
| `{ task: A, type: =C6*C$2, at: D6 }` | Types that formula in D6 (`at` defaults to the task's hint). |
| `{ task: A, type: …, at: D6, fill: D6:D12 }` | Types, then drags the fill handle over D6:D12. |
| `{ task: A, at: D6, fill: D6:D12 }` | Fills from D6 without typing. |
| `{ task: B, doit: true }` | Clicks the wand (Do it). Add `undo: true` to then show Undo it. |
| `{ audio: 1.mp3, transcript: … }` | Only narrates. Add `at: C2` to select a cell meanwhile. |
| `{ submit: true, audio: …, transcript: … }` | Presses Next, then Submit Now. |

`done: false` on a typing, fill or wand step leaves the task unfinished, so a later step can
finish it.

**Narration.** Each step can carry `audio` (a clip) and `transcript` (its words). If any step
has audio, every step needs both. A relative name like `1.mp3` is a file in the module's `audio/`
folder. Write only the transcripts and let `simmy narrate` add the names and make the clips. The
cursor's movement is fitted into each clip's length, so a step lasts as long as its sentence.

## Formulas

Formulas support `+ - * / ^ & %`, comparisons (`= <> < > <= >=`), ranges, `$` references and
these functions: ABS, AVERAGE, CONCAT, CORREL, COUNT, COUNTA, COUNTBLANK, COUNTIF, IF, LEFT, LEN,
LOWER, MAX, MEDIAN, MID, MIN, MODE, MODE.SNGL, RANDBETWEEN, RIGHT, ROUND, SEQUENCE, SORT, SQRT,
STDEV, STDEV.P, STDEV.S, SUM, TEXTJOIN, UNIQUE, UPPER, VAR, VAR.P, VAR.S. Defined names work in
formulas, and so do Excel 365's dynamic arrays: a formula that returns several values (SORT,
UNIQUE, SEQUENCE, `=A2:A9*2`) spills into the cells below and to the right, or shows `#SPILL!`
when something is in the way, and `A2#` refers to everything A2 spills. A function not on this
list is reported by `simmy validate`, so design tasks around these.
