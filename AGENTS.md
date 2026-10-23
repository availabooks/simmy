# simmy

An embeddable Excel simulator, published as `@availabooks/simmy`. It renders a model macOS or
Windows desktop with a working formula engine and a task pane that grades the student's work.
One YAML **module** describes the workbook, the tasks, the rubric and an optional narrated demo.
The README documents the public API (the `<simmy-app>` element, its attributes and events) and
the module format.

## Relationship to ee-atlas

The `ee-atlas` project (a sibling checkout at `../ee-atlas`) is the Excel Educator add-in that
simmy is modeled after. simmy simulates that add-in's task pane for people who don't have it
installed, so keep the two consistent:

- The task pane's behavior, wording and look, and the grading rule kinds (`value`, `reference`,
  `copied`), follow the add-in. When simmy and ee-atlas disagree, ee-atlas is the source of truth.
- Check ee-atlas before changing how a rule is graded or what the pane shows.
- simmy is a local approximation, not the real grader. Don't treat its scores as authoritative.

## Layout

- `src/engine/`: the spreadsheet engine, no UI. `sheet.ts` parses and evaluates formulas (dynamic
  arrays that spill, defined names), shifts references and formats values. `names.ts` has Excel's
  naming rules and what the name dialogs need. `workbook.ts` is the state the grid, Do it and
  grading share: cells, names and the seed that picks RANDBETWEEN's numbers.
- `src/module/`: `types.ts` (the module shape), `parse.ts` (validation, with path-specific errors),
  `grade.ts` (rubric rules, Do it and Undo it).
- `src/ui/`: Preact UI. `Simmy.tsx` (desktop and windows), `Workbook.tsx` (grid), `Ribbon.tsx`,
  `TaskPane.tsx`, `autopilot.ts` (the auto demo), `styles.css`. `Dialog.tsx` is the shell for
  Excel's dialogs (`NameDialogs.tsx`: Go To and the name dialogs); `Menu.tsx` is the dropdown.
- `src/element.ts`: the `<simmy-app>` custom element. It renders into a shadow root.
- `src/index.ts`: the public exports. It also registers the element.
- `hosted-modules/`: module sources (`<name>/config.yaml` and `<name>/audio/`). `modules-site/`: the Workers static-assets project that publishes them. `embed/` and `examples/`: iframe wrapper and a plain-HTML example.
- `test/`: Vitest tests, one file per area.
- `skill/simmy/`: the Claude skill for writing modules. `npm run publish-modules` puts it on the
  site as `/skill/simmy.zip`. Keep `module-format.md` in step with `parse.ts`; `test/skill.test.ts`
  checks its example and that it names every key.
- `docs/`: the documentation site, plain static files (`npm run docs` serves it). Its live demo
  uses the latest published build from jsDelivr and the hosted module, so it shows what users get.

## Commands

```sh
npm install
npm run dev      # http://localhost:5173; index.html uses deliberately hostile page styles
npm run check    # tsc --noEmit
npm test         # vitest run (happy-dom)
npm run build    # dist/index.js (npm), dist/simmy.js (standalone), dist/types, dist/embed.html
```

Run `npm run check` and `npm test` before committing. CI runs both, plus the build. UI test
coverage is thin, so check UI changes in the browser with `npm run dev`.

## Conventions

- Match the surrounding code: double quotes, 2-space indents, doc comments that explain why.
- Everything runs client-side. There is no backend.
- The host page's CSS must never reach the simulator, and the simulator's must never leak out.
  Keep styles inside the shadow root (`src/ui/styles.css`).
- New formula functions go in `sheet.ts` (with their arguments and Excel category in `FUNCTIONS`)
  and get a test in `test/functions.test.ts`; the docs list every function, which
  `test/docs.test.ts` checks. A new rubric rule kind touches `types.ts`, `parse.ts`, `grade.ts`,
  `keys.ts` and the skill's `module-format.md`, and should mirror an ee-atlas grading step.
- Excel for Windows (Microsoft 365) is the reference for how a feature looks and behaves.
  `plans/MO-210_verification.md` lists Microsoft's pages to check each feature against.

## Releasing

Bump `version` in `package.json`, update the version in the jsDelivr URLs in the README and in
`skill/simmy/SKILL.md` (`test/skill.test.ts` checks), and the badge beside the docs logo in
`docs/index.html` (`test/docs.test.ts` checks; the site itself always loads `@latest`). Commit,
then `git tag vX.Y.Z && git push --tags`. The Publish workflow checks that the tag matches
`package.json`, then publishes to npm through trusted publishing (no token). Don't tag a version
that is already on npm. Once npm has it, run `npm run publish-modules`: it deploys the docs site
and modules, and purges jsDelivr's cached `@latest` so the site shows the new release.
