# simmy

An embeddable Excel simulator: a model macOS or Windows desktop, Excel with a working
formula engine, and the Excel Educator task pane that grades the student's work. One
YAML **module** describes the workbook, the tasks and the rubric.

It ships as a custom element, `<simmy-app>`, that renders into a shadow root, so it
looks the same on any page.

## Use it

**Script tag** (one self-contained file; Preact, the styles and the YAML parser inside):

```html
<script type="module" src="https://cdn.jsdelivr.net/npm/@availabooks/simmy@0.1.0/dist/simmy.js"></script>

<simmy-app src="/modules/my-assignment.yaml"></simmy-app>
```

**npm** (shares Preact with your app; the YAML parser loads only when needed):

```sh
npm install @availabooks/simmy
```

```js
import "@availabooks/simmy"; // registers <simmy-app>
```

**iframe**, for hosts that can't run scripts (an LMS page, a CMS block). `dist/embed.html`
sits next to the standalone build:

```html
<iframe src="https://…/embed.html?module=https://…/my-assignment.yaml"
        style="width:100%;height:820px;border:0" title="Excel simulator"></iframe>
```

The embed page posts `{ type: "simmy:graded", report }` to its parent when a student submits.

## The element

| Attribute | |
|---|---|
| `src` | URL of a YAML (or JSON) module. |
| `platform` | `mac`, `win` or `auto` (default: match the visitor's computer). |
| `autoplay` | `false` starts with the visitor in control. The auto demo also stays off for `prefers-reduced-motion`. |
| `fit-screen` | Size the desktop to the viewport. |
| `chrome-offset` | With `fit-screen`, space taken by a sticky header: a length (`72px`) or a selector to measure (`.site-header`). |
| `history` | On phones, push a history entry so Back closes the full-screen view. |
| `fonts` | `none` skips loading Carlito and Inter from Google Fonts. |

Put a module inline instead of `src`:

```html
<simmy-app>
  <script type="text/yaml">
    workbook: { name: Totals }
    …
  </script>
</simmy-app>
```

or set it from script: `el.module = { … }`.

A child with `slot="head"` shows above the desktop, beside the toolbar. It stays in your page's
DOM, so your styles (and search engines) see it.

**Events** (they bubble): `simmy:ready` (detail: the module), `simmy:graded` (detail:
`{ tasks, earned, points }`), `simmy:error` (detail: the error; the element also shows the
problems in place).

**Theming:** set `--excel-green`, `--excel-green-dark`, `--excel-green-fill`, `--excel-grid`,
`--excel-ink` or `--font-sans` on the element.

## Modules

[`modules/references-and-calculations.yaml`](modules/references-and-calculations.yaml) is a
complete example. The shape:

```yaml
version: 1
workbook: { name: Workbook title, sheet: Sheet tab name }
assignment: { title: …, text: …, section: Assessment Section 1 }
sheet:
  rows: 40            # default 40
  columns: 20         # default 20
  columnWidth: 64     # px, default 64
  widths: { A: 12, B: 74 }
  start: D6           # the cell selected at the start; default: the first task's hint
  cells: { B2: Discount, C2: 0.3, E14: "=SUM(E6:E12)" }
  styles:             # applied in order; later entries win
    - { range: C2:C3, format: 0.00% }        # General, 0, 0.00, $#,##0.00, 0.00%, accounting
    - { range: B5:E5 D14:D16, fill: "#800000", color: "#ffffff", bold: true }
    - { range: D6:E12, fill: "#f0e68c", lined: true, align: center, italic: false }
  outlines:
    - { range: D6:E12, weight: 1, color: "#a6a6a6" }
tasks:
  - id: A             # default A, B, C…
    name: Discount    # default: the id
    text: Calculate the "Discount" in cell D6.
    details: [ … ]
    hint: D6          # default: the range's top-left cell
    range: D6:D12     # where Do it writes
    solution: =C6*$C$2
    points: 25        # default: the rules' points added up
    rules:
      - { kind: value, cell: D6, formula: =C6*$C$2, points: 5, message: … }
      - { kind: reference, cell: D6, ref: C2, forms: [C$2, $C$2], points: 5, message: … }
      - { kind: copied, range: D6:D12, points: 5, message: … }
note: { title: Try it, steps: [ … ] }
autopilot:            # optional; default: type each solution and fill its range
  - { task: A, type: =C6*C$2, at: D6, fill: D6:D12 }
  - { task: B, doit: true, undo: true }
poster: { range: B5:E9 }   # the still phones see; default: the top-left of the content
copy:                 # optional wording overrides
  paneTitle: Excel Educator
  unsupported: { video: { title: …, text: … }, book: { title: …, text: … } }
  submitSteps: [Validate Key, Prepare Rubric, Check for Completeness, Evaluate Tasks, Finalize Score]
```

Rules:

- `value`: the cell equals what `formula` gives on the student's own sheet.
- `reference`: the cell's formula refers to `ref`; with `forms`, only those `$` spellings count.
- `copied`: every cell in `range` holds the first cell's formula, copied down or across.

`parseModule()` checks a module and reports every problem with its path, for example
`tasks[1].rules[0].cell: "D66" is outside the sheet (20 columns, 40 rows)`.

Formulas support `+ - * / ^ & %`, comparisons, ranges, `$` references and ABS, AVERAGE,
CORREL, COUNT, COUNTA, COUNTIF, IF, MAX, MEDIAN, MIN, MODE, MODE.SNGL, ROUND, SQRT, STDEV,
STDEV.P, STDEV.S, SUM, VAR, VAR.P, VAR.S.

## Develop

```sh
npm install
npm run dev     # http://localhost:5173, a page with deliberately hostile styles
npm test
npm run check
npm run build   # dist/index.js (npm), dist/simmy.js (standalone), dist/embed.html
```

Publishing: push a `v*` tag; `.github/workflows/publish.yml` publishes to npm.
