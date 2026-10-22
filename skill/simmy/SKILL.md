---
name: simmy
description: Write, check, narrate, preview and publish simmy modules, the YAML files behind the simmy Excel simulator (<simmy-app>, @availabooks/simmy). Use when someone wants to create or edit a simmy module or an Excel practice assignment for simmy, write its tasks and rubric, add a narrated demo or generate its audio from transcripts, fix a module that won't load, embed one in a page, or publish modules with the simmy CLI.
---

# simmy modules

simmy is an embeddable Excel simulator. One YAML **module** describes the workbook, the tasks a
student does, the rubric that grades them, and an optional narrated demo. The `simmy` command
checks modules, makes their narration audio, and publishes them.

- `module-format.md` (next to this file): every key, what it does, and how grading works. Read
  it before writing or changing a module.
- `example-module.yaml`: a complete, valid module to start from.

## The command

Use `simmy` if it is installed (`simmy --help`). Otherwise run it through npx:
`npx -y @availabooks/simmy <command>`. It needs Node 20 or later.

| Command | What it does |
|---|---|
| `simmy validate [path...]` | Checks modules. A path is a `config.yaml`, a module folder or a folder of modules. |
| `simmy narrate [path...]` | Speaks each autopilot transcript into the module's `audio/` folder. |
| `simmy voices [filter]` | Lists narration voices, e.g. `simmy voices en-US`. |
| `simmy tts --out f.mp3 <text>` | Speaks one line into an MP3. |
| `simmy publish --origin <url>` | Deploys a folder of modules to Cloudflare Workers. |
| `simmy config set <key> <value>` | Saves a default (`origin`, `name`, `source`, `voice`, `rate`, `pitch`). |

## Layout

Each module is a folder: `<modules>/<name>/config.yaml`, with its clips in `<modules>/<name>/audio/`.
The folder name is the module's name in URLs, so use lowercase words and hyphens.

## Writing a module

1. Start from `example-module.yaml`, or from the user's existing module. Read `module-format.md`.
2. Lay out the sheet first: labels, given numbers, and the empty cells the student fills. Only
   cells the student doesn't write go in `sheet.cells`.
3. Write each task: what to do (`text`), where Do it writes (`range`), the answer (`solution`),
   and `rules` that award points. Prefer a `value` rule for the answer, a `reference` rule for each
   input cell that must be referenced rather than typed, and a `copied` rule when the formula is
   filled across a range.
4. Run `simmy validate <folder>` and fix everything it reports. Repeat until it says valid.
   - Lines starting with `-` are errors: the module won't load, or its own solution fails its
     rubric (`Do it loses N points`). Fix the solution or the rule, whichever is wrong.
   - Lines starting with `?` are keys simmy ignores. They are almost always typos. Fix them.
     `--strict` makes them errors.

Don't report a module as finished until `simmy validate` passes. Validation can't tell whether
the task wording matches the rules, so read each task's `text` against its rules yourself.

### YAML pitfalls

- Inside `{ … }`, a comma ends the value. Quote any text with a comma:
  `message: "Use C2, not a typed number."`
- `#` after a space starts a comment. Quote number formats: `format: "$#,##0.00"`.
- Quote colors: `fill: "#1f4e78"`.
- Text that starts with `=`, `[`, `{`, `*`, `&` or `!`, or that holds `: `, needs quotes.
- Formulas are written as the student would type them: `=SUM(E5:E7)+$C$2`.

## Narrating the demo

The autopilot is a scripted demo that plays the assignment. With narration, each step has a clip
and its transcript.

1. Write the `autopilot` entries with a `transcript` and **no `audio`**. Keep each to one or two
   sentences that match what the step shows. Say cell names the way they should sound: "cell D6".
2. Run `simmy narrate --dry-run <folder>` to see the plan, then `simmy narrate <folder>`. It
   names each new clip (`1.mp3`, `2.mp3`, …), writes the name into `config.yaml`, and makes the
   MP3. Don't pick clip names or edit `narration.json` by hand.
3. After editing a transcript or changing the voice, run `simmy narrate` again. Only those
   clips are remade. Clips that `simmy narrate` didn't make, such as recorded ones, are kept
   unless `--force`. Ask before using `--force`, because it replaces recordings.
4. Before or between steps, transcript-only entries make `simmy validate` fail, because each
   entry needs its clip. That clears once `simmy narrate` has run.

Voices: the default is `en-US-AvaMultilingualNeural`. Change it with `--voice en-US-AndrewNeural`, `--rate +10%`, `--pitch -2Hz`. Save a choice with
`simmy config set voice …`. Use one voice for every clip in a module. The voices come from
Microsoft Edge's read-aloud service. It needs no account but is unofficial, so if `narrate`
fails for every clip, the service may be down or changed. Say so rather than retrying in a loop.

## Previewing

Fetching a module needs a web server. Put a page beside the modules folder:

```html
<!doctype html>
<meta charset="utf-8" />
<title>Preview</title>
<script type="module" src="https://cdn.jsdelivr.net/npm/@availabooks/simmy@0.3.1/dist/simmy.js"></script>
<simmy-app src="modules/my-module/config.yaml" fit-screen></simmy-app>
```

Serve the folder (`npx -y serve .` or `python3 -m http.server`) and open the page. The Play
button runs the demo, Do it (the wand) writes each solution, and Submit grades the sheet.

## Publishing

`simmy publish --origin https://<worker>.<account>.workers.dev --source <modules folder>`
validates every module, writes absolute clip URLs, and deploys the folder as a Cloudflare Workers
site through wrangler. Any page can then load
`https://<worker>.<account>.workers.dev/<name>/config.yaml`.

- It needs `npx wrangler login` once. That opens a browser, so the user has to run it.
- Publishing is public. Confirm with the user before running it without `--dry-run`.
- `--dry-run` builds the site into a temporary folder and prints its path, without deploying.
- The origin is written into the clip URLs, so it must be the URL the site is served from.

## Embedding

```html
<script type="module" src="https://cdn.jsdelivr.net/npm/@availabooks/simmy@0.3.1/dist/simmy.js"></script>
<simmy-app src="https://…/my-module/config.yaml"></simmy-app>
```

Attributes include `platform` (`mac`, `win`, `web`, `auto`), `mode` (`movie` or `kiosk`) and
`fit-screen`. The element fires `simmy:graded` with `{ tasks, earned, points }` when a student
submits. For pages that can't run scripts, use the iframe at `dist/embed.html?module=<url>`.
The full reference is at https://simmy-modules.cotc.workers.dev/.

simmy grades locally to approximate the Excel Educator add-in. Its scores are for practice, not
a grade of record.
