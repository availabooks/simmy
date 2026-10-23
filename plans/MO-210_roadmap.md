# MO-210 roadmap

How simmy will cover every skill in [MO-210_Teaching_Checklist.md](MO-210_Teaching_Checklist.md), one
phase at a time. [MO-210_verification.md](MO-210_verification.md) lists Microsoft's pages to check
each feature against. Update this file as a phase moves along, in the same commits as the work.

## Decisions

- **Real features:** every feature really works and can be graded by a rubric rule; none are just
  for show.
- **Windows is the reference:** Excel for Microsoft 365 on Windows, the version the exam uses, sets
  how features look and behave. They work on all three desktops (Windows, Mac and web) using the
  Windows layout.
- **File and system features are simulated:** import, export, print and inspect show Excel's
  dialogs and record the student's choices so they can be graded. No real files or network.
- **Grading follows ee-atlas:** new rules mirror the add-in's grading steps. Where the add-in can't
  grade something (freeze panes, filters, views, Show Formulas, the Quick Access Toolbar,
  hyperlinks, comments), simmy gets its own rule.
- **Releases:** one per phase, and Eli approves each one before it's tagged. excel-ed.tech loads
  simmy `@latest`, so a release goes live within minutes.

## Status

| Phase | Scope | Size | Status |
|---|---|---|---|
| 1 | Formulas and functions | L | Built on the `phase-1-formulas` branch; not merged or released |
| 2 | Workbook model and frameworks | L | Planned |
| 3 | Cell formatting | L | Planned |
| 4 | Data and navigation | L | Planned |
| 5 | Tables, sort and filter | L | Planned |
| 6 | Page layout, printing, views and windows | XL | Planned |
| 7 | Conditional formatting and sparklines | M | Planned |
| 8 | Charts | XL | Planned |
| 9 | File, import, inspect and options | L | Planned |

Each phase gets its own detailed plan before work starts.

## What "done" means for a phase

- It works, with Excel for Windows' wording and placement.
- It has unit tests.
- Its new rule kinds and Do it support are wired through `types.ts`, `parse.ts`, `keys.ts` and
  `grade.ts`.
- The skill's `module-format.md` is updated (`test/skill.test.ts` checks it).
- A hosted demo module passes `simmy validate`.
- `npm run check`, `npm test` and the build pass.
- It has been checked in the browser against the Microsoft pages for that phase.
- It has been released, after Eli approves.

## Phases

### 1. Formulas and functions

Built: dynamic arrays that spill (the blue border, `#SPILL!`, `A1#`), and these functions:
COUNTBLANK, LEFT, RIGHT, MID, UPPER, LOWER, LEN, CONCAT, TEXTJOIN, SEQUENCE, SORT, UNIQUE and
RANDBETWEEN. Also:

- **Calculation and references:** Calculate Now (F9) and F4 reference cycling.
- **Names:** defined names, with the Name Box, Go To (F5 or Ctrl+G), New Name, Create Names from
  Selection, the Name Manager, Use in Formula and Paste Name (F3).
- **Function Library menus:** filled in on the Formulas tab.
- **Rules:**
  - `equals`, `contains`, `name`, `between` and `selected`;
  - `value` over a spilled range;
  - module and task `names`.
- **Practice module:** `hosted-modules/formulas-and-names`.

Left for later:

- **New Name's Refers to:** clicking cells to fill it in.
- **Disabled controls:** Go To Special (phase 2), and the Name Manager's Filter.
- **Autopilot:** a step that types into the Name Box.
- **Checking against real Excel:** the Go To, New Name and Create Names from Selection dialogs,
  which Microsoft's pages don't picture.

### 2. Workbook model and frameworks

- **State:** one workbook state for everything that gets graded, with undo across all of it.
- **Multiple sheets:** sheet tabs and their right-click menu.
- **Shared pieces:**
  - a command registry for the ribbon, menus, shortcuts and the autopilot;
  - the shells for dialogs, the right-click menu and the File view (backstage);
  - a `sim_property` rule for what ee-atlas can't grade.
- **Features:** Show Formulas (Ctrl+`) and Go To Special.

### 3. Cell formatting

- **Format Cells:** every tab, including Protection.
- **Number formats:** currency, accounting, percentage, dates and decimal places.
- **Alignment:** orientation and indent; wrap text; Merge & Center and Merge Across.
- **Styles and painting:** cell styles, the Format Painter, and Clear Formats.
- **Sizes:** row height and column width, including AutoFit.
- **Group worksheets.**

### 4. Data and navigation

- **Paste Special:** values, formats, transpose and operations.
- **Auto Fill:** series and patterns, and the Fill Series dialog.
- **Insert and delete:** rows, columns and cells, with references adjusted.
- **Find & Replace,** including Within: Workbook.
- **Hyperlinks,** and comments versus notes.

### 5. Tables, sort and filter

- **Tables:** create one, and convert it back to a range; styles and style options; the Total
  Row; adding rows and columns.
- **Structured references.**
- **Sorting and filtering:** the Sort dialog with several levels, and AutoFilter.

### 6. Page layout, printing, views and windows

- **Page Setup:** orientation, margins, scaling and size; headers and footers; the print area,
  print titles and print settings.
- **Views:** Normal, Page Layout and Page Break Preview.
- **Windows:** Freeze Panes, Split, New Window and Arrange.

### 7. Conditional formatting and sparklines

- **Conditional formatting:** highlight rules, top/bottom, data bars, color scales and icon sets,
  and clearing rules (from the selection or the whole sheet).
- **Sparklines:** line, column and win/loss, with markers and styles.

### 8. Charts

Charts are drawn by simmy itself, as SVG.

- **Making and placing:** create a chart, and move it to a chart sheet.
- **Data:** Select Data and adding a series; Switch Row/Column.
- **Look:** chart elements, Quick Layout and chart styles.
- **Alt text.**

### 9. File, import, inspect and options

- **Workbook properties.**
- **Checkers:** the Document Inspector, the Accessibility Checker and the Compatibility Checker.
- **Saving:** Save As and Export, with the choices recorded.
- **Options** and the Quick Access Toolbar.
- **Get Data:** from text/CSV and from the web, using data that comes with the module.

## Where each checklist item lands

| Checklist section | Phase |
|---|---|
| Import data (text/CSV, web) | 9 |
| Find across a workbook | 4 |
| Name Box and Go To | 1 (Go To Special in 2) |
| Hyperlinks | 4 |
| Page setup; headers and footers; views | 6 |
| Row height and column width | 3 |
| Options and the Quick Access Toolbar | 9 |
| Freeze panes; New Window, Arrange, Split | 6 |
| Workbook properties | 9 |
| Show formulas | 2 |
| Print area; print settings | 6 |
| Save and export in other formats; inspect workbooks | 9 |
| Comments and notes | 4 |
| Paste Special; Auto Fill; insert and delete rows, columns and cells | 4 |
| RANDBETWEEN and SEQUENCE | 1 |
| Merge; alignment; Format Painter; wrap; number formats; Format Cells; cell styles; clear formats; group worksheets | 3 |
| Named ranges (define, use in formulas) | 1 |
| Sparklines; conditional formatting | 7 |
| Tables: create, style, convert, rows and columns, style options, Total Row, filter, sort | 5 |
| Relative, absolute and mixed references (F4) | 1 |
| Structured references | 5 |
| SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, IF | 1 (most already existed) |
| SORT and UNIQUE; LEFT, RIGHT, MID; UPPER, LOWER, LEN; CONCAT and TEXTJOIN | 1 |
| Charts: create, chart sheets, Select Data, switch rows and columns, elements, layouts and styles, alt text | 8 |
