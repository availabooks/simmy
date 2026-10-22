# simmy

An embeddable Excel simulator: a model macOS or Windows desktop, Excel with a working
formula engine, and the Excel Educator task pane that grades the student's work. One
YAML **module** describes the workbook, the tasks and the rubric.

It ships as a custom element, `<simmy-app>`, that renders into a shadow root, so it
looks the same on any page.

## Use it

**Script tag** (one self-contained file; Preact, the styles and the YAML parser inside):

```html
<script type="module" src="https://cdn.jsdelivr.net/npm/@availabooks/simmy@0.2.3/dist/simmy.js"></script>

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
| `platform` | `mac`, `win`, `web` or `auto` (default: match the visitor's computer, macOS or Windows). `web` is Excel on its own, filling the frame like a browser tab, with no OS chrome. |
| `mode` | `movie` (default) or `kiosk`. See below. |
| `ribbon` | `open` (default) or `collapsed`: start with the full ribbon, or with only its tabs. Either way the visitor can collapse or expand it (the corner button, a double-click on a tab, or Ctrl+F1), and a collapsed ribbon shows over the sheet when a tab is clicked. |
| `autoplay` | `false` starts with the visitor in control, with no Play screen. A kiosk also starts in control for `prefers-reduced-motion`. |
| `fit-screen` | Fit the whole simulator, heading and playback bar included, to the viewport. It never gets shorter than `--simmy-min-height` (600px); on a shorter screen the page scrolls. |
| `chrome-offset` | With `fit-screen`, space taken by a sticky header: a length (`72px`) or a selector to measure (`.site-header`). |
| `history` | On phones, push a history entry so Back closes the full-screen view. |
| `fonts` | `none` skips loading Carlito and Inter from Google Fonts. |

**Modes.** In `movie` mode (the default) nothing plays until the visitor asks: the desktop opens
darkened behind a **Play** button and a **Try it yourself** button. The demo plays once, with sound,
and stops on the finished sheet with the visitor in control; Play starts it over. The desktop
switcher and Start over sit in a settings panel behind the gear on the playback bar, so the page
around the movie stays clear. In `kiosk` mode the demo starts by itself, muted (the visitor can
unmute from the playback bar), and loops, and the desktop switcher and Start over sit above the
desktop.

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
`--excel-ink` or `--font-sans` on the element. With `fit-screen`, `--simmy-min-height` sets the
shortest the desktop gets (default `600px`).

## Modules

[`hosted-modules/references-and-calculations/config.yaml`](hosted-modules/references-and-calculations/config.yaml)
is a complete example. Its clips are files in the `audio` folder beside that config. The shape:

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
  - { audio: 1.mp3, transcript: Peter's Pool Palace is having a sale. }
  - { task: A, type: =C6*C$2, at: D6, done: false, audio: 2.mp3, transcript: Type the discount in D6. }
  - { task: A, at: D6, fill: D6:D12, audio: 3.mp3, transcript: Fill it down through D12. }
  - { submit: true, audio: 4.mp3, transcript: Press Next, then Submit Now. }
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

Each narrated entry has its own `audio` file and `transcript`. The clip loads when that line starts, plays at its own speed, and the cursor fits the entry into the clip. An entry with only those two fields waits on the line. `done: false` types or uses the wand without marking the task finished. `submit: true` presses Next, then Submit Now. Leave `audio` and `transcript` off and the demo keeps its own pace.

A relative `audio` name, such as `1.mp3`, is a file in the `audio` folder next to the module. When the module is loaded from `src`, that name is resolved from the module's URL, so `hosted-modules/pool/config.yaml` plays `hosted-modules/pool/audio/1.mp3`. A full URL or a path that starts with `/` is used as written. Inline YAML has no folder, so a relative name there is left unchanged.

A bar along the bottom of the desktop pauses, resumes, restarts, rewinds, fast-forwards, mutes, and changes the playback speed. It shows the elapsed time and the total. Dragging the time bar moves the cursor and the sheet with it. Clicking the sheet or the task pane pauses the narration where it is and hides the cursor; resume continues from that moment.

`parseModule()` checks a module and reports every problem with its path, for example
`tasks[1].rules[0].cell: "D66" is outside the sheet (20 columns, 40 rows)`.

Formulas support `+ - * / ^ & %`, comparisons, ranges, `$` references and ABS, AVERAGE,
CORREL, COUNT, COUNTA, COUNTIF, IF, MAX, MEDIAN, MIN, MODE, MODE.SNGL, ROUND, SQRT, STDEV,
STDEV.P, STDEV.S, SUM, VAR, VAR.P, VAR.S.

## Develop

```sh
npm install
npm run dev     # http://localhost:5173, a page with deliberately hostile styles
npm run docs    # the documentation site in docs/, with a live demo
npm test
npm run check
npm run build   # dist/index.js (npm), dist/simmy.js (standalone), dist/embed.html
```

Publishing the package: push a `v*` tag; `.github/workflows/publish.yml` publishes to npm.

Publishing your own modules from any project: install the CLI once with `npm i -g @availabooks/simmy` (or `npm link` in this checkout), then run `simmy publish --origin https://my-modules.<account>.workers.dev` in a folder that has `<module>/config.yaml` and `<module>/audio/` (default `./hosted-modules`, or `--source`). It rewrites clip names to absolute URLs and deploys the result as a Workers static-assets site named after the origin's first label (`--name` to override), through `npx wrangler` (log in once with `npx wrangler login`). `--dry-run` builds the site into a temp folder and prints its path. Publishing validates every module first and deploys nothing if one has a problem. To check modules without publishing, run `simmy validate` (the source folder by default, or any `config.yaml`, module folder or folder of modules). It lists every problem with the path to the field, flags clips missing from `audio/`, and exits with 1 on failure, so it can run in CI. Keys simmy doesn't know, usually misspellings such as `colunmWidth`, are warnings with a suggestion; the module still loads with the default. `--strict` fails on them too.

Making the narration: write each autopilot entry's `transcript` and run `simmy narrate` (same paths as `validate`). It speaks each transcript with Microsoft Edge's read-aloud voices into the module's `audio/` folder, in the same 24 kHz mono MP3 format as the hosted clips. An entry with a transcript and no `audio` gets the next numbered clip, written into `config.yaml` without disturbing its comments or layout. `narration.json`, beside `config.yaml`, records what each clip was made from, so editing a transcript or changing the voice and running it again remakes only those clips. A clip it has no record of (recorded by hand, say) is kept unless `--force`. `--dry-run` shows what would happen. Pick a voice with `--voice` (`simmy voices en-US` lists them), and adjust it with `--rate +10%` or `--pitch -2Hz`. All three can be saved with `simmy config set voice …` or put in `simmy.json`. `simmy tts --out clip.mp3 Some text` speaks one line. The service needs no account, but it is unofficial and could change or stop working.

A Claude skill for writing modules (the format, the grading rules and these commands) is in [`skill/simmy`](skill/simmy) and downloadable from the docs site at `/skill/simmy.zip`. Defaults can go in a `simmy.json` (`origin`, `name`, `source`). To publish the same project from any folder, save them globally: `simmy config set origin https://…`, `simmy config set source ~/work/my-modules` (stored in `~/.config/simmy/config.json`; `simmy config` lists, `unset <key>` removes). Flags beat `./simmy.json`, which beats the global config.

Publishing the module library: `hosted-modules/<name>/config.yaml` keeps clip names relative (`1.mp3`). `npm run publish-modules` rewrites them to absolute URLs, puts the docs site (`docs/`) at the root with the module list at `/modules`, then deploys `host-dist/` as Workers static assets. The public origin is `simmyOrigin` in `modules-site/package.json` (`https://simmy-modules.cotc.workers.dev`). A page on another origin can load `https://simmy-modules.cotc.workers.dev/references-and-calculations/config.yaml` and the clips play. `npm run host-modules -- --base https://…` writes that directory without deploying.
