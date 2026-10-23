# Checking simmy against real Excel (MO-210)

Microsoft's own documentation for every item in
[MO-210_Teaching_Checklist.md](MO-210_Teaching_Checklist.md), for checking by eye that simmy
looks and behaves like **Excel for Microsoft 365 on Windows**, the version the exam uses. Each
phase's browser check works through its section.

"Screenshots" says whether the page has images, and of what. Look at the images themselves:
they were found by their descriptions.

## What the pages can't show

- **Missing dialogs:** no page has a screenshot of Go To, Go To Special, New Name or Create
  Names from Selection. Check those against real Excel for Windows (or Microsoft's training
  videos), in a short session at the end of each phase.
- **Inconsistent examples:** a few pages contradict themselves. Check the live page before
  copying their numbers into tests:
  - the COUNT examples' ranges (A2:A6 vs A2:A7);
  - SORT's source range (A2:A17 vs A5:D20);
  - UNIQUE's caption, which says "occurs_once" for `exactly_once`.
- **Wrong platform:** two pages that look relevant are for the Mac. Use the Windows pages
  listed below instead:
  - the CSV import page `/excel/import-data-from-a-csv-html-or-text-file`;
  - "View multiple panes, sheets, or workbooks".

## Phase 1: formulas and functions

### Dynamic arrays and spill
- https://support.microsoft.com/en-us/excel/dynamic-array-formulas-and-spilled-array-behavior
- https://support.microsoft.com/en-us/excel/how-to-correct-a-spill-error
- https://support.microsoft.com/en-US/excel/spilled-range-operator

**Screenshots:** `=SORT(D2:D11,1,-1)` spilling from F2, the blue border, a ghosted formula, and
`#SPILL!` with one image per cause. `=SUM(A2#)`.

**Check:**
- Selecting any cell of a spill draws a border around the whole range. Selecting a cell outside
  it removes the border.
- A non-anchor cell shows the formula greyed ("ghosted") in the formula bar. Only the first cell
  is editable.
- A blocked spill shows `#SPILL!`, with a dashed border around the range it would fill.
- Causes: "Spill range isn't blank", "Extends beyond the worksheet's edge", "Spill into merged
  cells", "Table formula".
- A2 `=SEQUENCE(10)`, then `=SUM(A2#)` gives 55, and updates if A2 becomes `=SEQUENCE(20)`.

### SORT, UNIQUE, SEQUENCE
- https://support.microsoft.com/en-us/excel/sort-function
- https://support.microsoft.com/en-us/excel/unique-function
- https://support.microsoft.com/en-us/excel/sequence-function

**Check:**
- **SORT:**
  - `=SORT(array,[sort_index],[sort_order],[by_col])`, with defaults 1, 1 (ascending; -1 is
    descending) and FALSE.
  - The result has the same shape as the array.
- **UNIQUE:**
  - `=UNIQUE(array,[by_col],[exactly_once])`, both defaulting to FALSE.
  - `exactly_once` TRUE returns only values that appear once.
  - Tip example: `=SORT(UNIQUE(B2:B12&" "&A2:A12))`.
- **SEQUENCE:**
  - `=SEQUENCE(rows,[columns],[start],[step])`, all defaulting to 1.
  - `=SEQUENCE(4,5)` is a 4×5 grid of 1–20, filled across each row first.
  - `=SEQUENCE(5,1,1001,1000)` gives 1001, 2001, 3001, 4001, 5001.

### RANDBETWEEN and recalculation
- https://support.microsoft.com/en-us/office/randbetween-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685
- https://support.microsoft.com/en-us/Excel/change-formula-recalculation-iteration-or-precision-in-excel

**Check:**
- `RANDBETWEEN(bottom, top)`: both arguments required, and it returns an integer.
- It gives a new number every time the sheet calculates: after any entry, or on F9.
- F9 recalculates; Shift+F9 recalculates the active sheet; Ctrl+Alt+F9 recalculates every
  formula.

### Text functions
- **TEXTJOIN:** https://support.microsoft.com/en-us/office/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c
- **CONCAT:** https://support.microsoft.com/en-us/office/concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2
- **LEFT:** https://support.microsoft.com/en-us/office/left-leftb-functions-9203d2d2-7960-479b-84c6-1ea52b99640c
- **RIGHT:** https://support.microsoft.com/en-us/office/right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f
- **MID:** https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028
- **UPPER:** https://support.microsoft.com/en-us/office/upper-function-c11f29b3-d1a3-4537-8df6-04d0049963d6
- **LOWER:** https://support.microsoft.com/en-us/office/lower-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4
- **LEN:** https://support.microsoft.com/en-us/office/len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb

**Check:**
- **TEXTJOIN:**
  - `TEXTJOIN(delimiter, ignore_empty, text1, ...)`.
  - With `ignore_empty` FALSE, empty cells leave ", , ".
  - A range can be the delimiter.
- **CONCAT:**
  - has no delimiter, and joins a range row by row;
  - `=CONCAT(B2:C8)` gives "a1b1a2b2a4b4…";
  - `=CONCAT(B2," ",C2)` gives "Andreas Hauser".
- **LEFT and RIGHT:**
  - `num_chars` defaults to 1 and can't be negative; more than the text length returns all of
    it;
  - `=LEFT("Sale Price",4)` gives "Sale"; `=RIGHT("Sale Price",5)` gives "Price".
- **MID** (with A2 "Fluid Flow"):
  - `=MID(A2,1,5)` gives "Fluid", `=MID(A2,7,20)` gives "Flow", `=MID(A2,20,5)` gives "";
  - a start below 1, or a negative count, gives `#VALUE!`.
- **UPPER, LOWER, LEN:**
  - `UPPER("total")` gives "TOTAL"; `LOWER("Apt. 2B")` gives "apt. 2b";
  - `LEN("Phoenix, AZ")` gives 11, and LEN of a blank cell gives 0.

### COUNTBLANK, COUNT, COUNTA
- **COUNTBLANK:** https://support.microsoft.com/en-us/office/countblank-function-6a92d772-675c-4bee-b346-24af6bd3ac22
- **COUNT:** https://support.microsoft.com/en-us/office/count-function-a59cd7fc-b623-4d93-87a4-d23bf411294c
- **COUNTA:** https://support.microsoft.com/en-us/office/counta-function-7dc98875-d5c1-46f1-9a82-53f3219e2509

**Check:**
- **COUNTBLANK** counts a formula's `""` result as blank, but not zeros.
- **COUNTA** counts everything but truly empty cells, errors and `""` included. 39790, 19,
  22.24, TRUE and #DIV/0! give 5.
- **COUNT** counts only the numbers in a range.

### Relative, absolute and mixed references (F4)
https://support.microsoft.com/en-us/Excel/switch-between-relative-absolute-and-mixed-references

**Screenshots:** a formula copied two cells down and two right.

**Check:**
- Select a reference in the formula and press F4 to cycle A1 → $A$1 → A$1 → $A1 → A1.
- Copied two down and two right: $A$1 stays $A$1, A$1 becomes C$1, $A1 becomes $A3, and A1
  becomes C3.

### Defined names
- https://support.microsoft.com/en-US/Excel/get-started/define-and-use-names-in-formulas
- https://support.microsoft.com/en-us/excel/names-in-formulas
- https://support.microsoft.com/en-us/Excel/use-the-name-manager-in-excel

**Screenshots:** the Name Box and the Name Manager.

**Check:**
- **Create Names from Selection:** check boxes Top row, Left column, Bottom row, Right column.
- **New Name:** Name, Scope (default Workbook), Comment (up to 255 characters) and Refers to.
  Clicking a cell inserts an absolute reference such as `Sheet1!$B$1`.
- **Name Manager:**
  - columns Name, Value, Refers To, Scope, Comment;
  - Edit, Delete (asks to confirm) and Filter;
  - can't be used while editing a cell.
- **Naming rules:**
  - the first character is a letter, `_` or `\`; the rest are letters, numbers, `.` or `_`;
  - no spaces, and nothing that looks like a cell reference (Z$100, R1C1);
  - C, c, R and r alone aren't allowed;
  - at most 255 characters, and case-insensitive.
- **Using names:** type the name, pick it from Formula AutoComplete, or use Formulas › Use in
  Formula. Deleting a name that a formula uses gives `#NAME?`.

### Go To and the Name Box
- https://support.microsoft.com/en-us/Excel/select-specific-cells-or-ranges-in-excel
- https://support.microsoft.com/en-us/excel/find-and-select-cells-that-meet-specific-conditions-in-excel
- https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f

**Check:**
- **Go To** (F5 or Ctrl+G) has a "Go to" list and a "Reference" box. `A1#` selects a spill.
  Special… opens Go To Special.
- **The Name Box** takes B3 or B1:B3 and Enter. Its arrow lists the names.
- **F3** opens Paste Name.

### Structured references (tables phase)
https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables

**Check:**
- `=[@[Sales Amount]]*[@[% Commission]]`.
- Special items #All, #Data, #Headers, #Totals and @.

## Later phases

| Checklist item | Microsoft page | Screenshots |
|---|---|---|
| Import text/CSV | https://support.microsoft.com/en-us/excel/get-started/import-or-export-text-txt-or-csv-files | no |
| Import from web | https://support.microsoft.com/en-us/excel/get-started/import-data-from-the-web | no |
| Find & Replace, Within: Workbook | https://support.microsoft.com/en-us/excel/get-started/find-or-replace-text-and-numbers-on-a-worksheet | yes: the dialog and its options |
| Hyperlinks | https://support.microsoft.com/en-us/office/work-with-links-in-excel-7fc80d8d-68f9-482f-ab01-584c44d72b3e | yes (Excel 2016) |
| Page setup | https://support.microsoft.com/en-us/excel/page-setup | yes: every tab |
| Row height, column width, AutoFit | https://support.microsoft.com/en-us/office/change-the-column-width-and-row-height-72f5e3cc-994d-43e8-ae58-9774a0905f46 | yes |
| Headers and footers | https://support.microsoft.com/en-us/excel/headers-and-footers-in-a-worksheet | yes |
| Quick Access Toolbar | https://support.microsoft.com/en-us/office/customize-the-quick-access-toolbar | yes |
| Normal / Page Layout / Page Break Preview | https://support.microsoft.com/en-us/office/insert-move-or-delete-page-breaks-in-a-worksheet-ad3dc726-beec-4a4c-861f-ed640612bdc2 | yes (no single page covers all three views) |
| Freeze panes | https://support.microsoft.com/en-us/excel/get-started/freeze-panes-to-lock-rows-and-columns | one image |
| New Window, Arrange | https://support.microsoft.com/en-us/excel/compare-two-or-more-worksheets-at-the-same-time | yes |
| Split | https://support.microsoft.com/en-us/excel/split-panes-to-lock-rows-or-columns-in-separate-worksheet-areas | yes |
| Workbook properties | https://support.microsoft.com/en-us/office/collab-files/view-or-change-the-properties-for-an-office-file | yes |
| Show formulas | https://support.microsoft.com/en-us/excel/display-or-hide-formulas | yes |
| Print area | https://support.microsoft.com/en-us/excel/set-or-clear-a-print-area-on-a-worksheet | yes (Excel 2013) |
| Print titles and settings | https://support.microsoft.com/en-us/office/repeat-specific-rows-or-columns-on-every-printed-page-0d6dac43-7ee7-4f34-8b08-ffcc8b022409 | yes |
| Save as PDF | https://support.microsoft.com/en-us/office/save-or-convert-to-pdf-or-xps-in-office-desktop-apps-d85416c5-7d77-4fd6-a216-6f4bf7c7c110 | few, none of Excel |
| Save in other formats | https://support.microsoft.com/en-us/office/save-a-workbook-in-another-file-format-6a16c862-4a36-48f9-a300-c2ca0065286e | not checked |
| Document Inspector | https://support.microsoft.com/en-us/office/collab-files/remove-hidden-data-and-personal-information-by-inspecting-documents-presentations-or-workbooks | none of Excel |
| Accessibility Checker | https://support.microsoft.com/en-us/accessibility/office-accessibility/improve-accessibility-with-the-accessibility-checker | yes |
| Compatibility Checker | https://support.microsoft.com/en-us/office/check-file-compatibility-with-earlier-versions-d9856881-5875-4c58-915f-06859b2943a7 | no |
| Comments vs notes | https://support.microsoft.com/en-us/excel/the-difference-between-threaded-comments-and-notes | yes |
| Paste Special | https://support.microsoft.com/en-us/excel/paste-options | yes |
| Auto Fill, Fill Series | https://support.microsoft.com/en-us/excel/get-started/fill-data-automatically-in-worksheet-cells | icons only |
| Insert/delete rows, columns, cells | https://support.microsoft.com/en-us/excel/get-started/insert-or-delete-rows-and-columns-in-excel | one image |
| Merge & Center vs Merge Across | https://support.microsoft.com/en-us/excel/get-started/merge-and-unmerge-cells-in-excel | yes (doesn't explain Merge Across) |
| Alignment, orientation, indent | https://support.microsoft.com/en-us/excel/get-started/align-or-rotate-text-in-a-cell | yes |
| Format Painter | https://support.microsoft.com/en-us/office/copy-cell-formatting-b9fe82ea-c0a0-41de-837b-d2f15dd41ea9 | no |
| Wrap text | https://support.microsoft.com/en-us/excel/wrap-text-in-a-cell-in-excel | one image |
| Number formats | https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2 | yes |
| Format Cells, Protection tab | https://support.microsoft.com/en-us/excel/get-started/lock-or-unlock-specific-areas-of-a-protected-worksheet | yes |
| Cell styles | https://support.microsoft.com/en-us/office/apply-create-or-remove-a-cell-style-472213bf-66bd-40c8-815c-594f0f90cd22 | one image |
| Clear formats | https://support.microsoft.com/en-us/excel/clear-cells-of-contents-or-formats | no |
| Group worksheets | https://support.microsoft.com/en-us/excel/group-worksheets | yes |
| Sparklines | https://support.microsoft.com/en-us/excel/analyze-trends-in-data-using-sparklines | yes |
| Conditional formatting, clear rules | https://support.microsoft.com/en-us/excel/use-conditional-formatting-to-highlight-information-in-excel | yes |
| Table styles and style options | https://support.microsoft.com/en-us/excel/format-an-excel-table | yes |
| Convert a table to a range | https://support.microsoft.com/en-us/excel/convert-an-excel-table-to-a-range-of-data | not checked |
| Total Row | https://support.microsoft.com/en-us/office/total-the-data-in-an-excel-table-6944378f-a222-4449-93d8-474386b11f20 | yes |
| Resize a table | https://support.microsoft.com/en-us/excel/resize-a-table-by-adding-or-removing-rows-and-columns-in-excel | not checked |
| Filter | https://support.microsoft.com/en-us/excel/get-started/filter-data-in-a-range-or-table-in-excel | yes |
| Sort by several columns | https://support.microsoft.com/en-us/excel/sort-data-in-a-range-or-table-in-excel | yes |
| Create charts | https://support.microsoft.com/en-us/excel/get-started/create-a-chart-from-start-to-finish | yes (styles and Switch Row/Column only in its macOS section) |
| Move chart to a chart sheet | https://support.microsoft.com/en-us/office/move-or-resize-a-chart-f9d6087b-1d7b-4a72-8d81-1d0788ea783f | yes |
| Select Data, add series | https://support.microsoft.com/en-us/office/add-a-data-series-to-your-chart | not checked |
| Switch Row/Column | https://support.microsoft.com/en-us/excel/change-how-rows-and-columns-of-data-are-plotted-in-a-chart | yes |
| Chart titles, labels, legend | https://support.microsoft.com/en-us/office/excelexp/add-or-remove-titles-in-a-chart | not checked |
| Quick Layout, chart styles | https://support.microsoft.com/en-us/office/change-the-layout-or-style-of-a-chart-a346e438-d22a-4540-aa87-bce9feb719cf | not checked |
| Alt text | https://support.microsoft.com/en-us/office/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669 | yes |
