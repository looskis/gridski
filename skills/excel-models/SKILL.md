---
name: excel-models
description: Read, audit, and build financial models in Microsoft Excel on macOS through the gridski MCP server. Use when reviewing a workbook's formulas, checking a model for consistency errors, or building period/time-series sheets in Excel.
---

# Working with Excel models

gridski drives the copy of Excel running on this Mac. It edits the analyst's live
workbook — what you write appears on their screen immediately, and **Excel's Undo will
not reverse it**. Keep the `previous_formulas` that `write_range` and `fill_right`
return until the user has confirmed the change.

Excel must already be running; gridski never launches it. Every tool defaults
`workbook` and `sheet` to whatever is active, so name them only when acting on
something else.

## Start by looking

Before any other call:

- `get_selection` — the cells the user has selected. This is the fastest way to learn
  what they mean by "this row" or "the model".
- `list_workbooks` — what is open, each sheet's used range, which is active.

Then `read_range` for the content. It returns values with empty cells as `null`, dates
as `YYYY-MM-DD`, errors as their codes (`#DIV/0!`), and formulas sparsely, keyed by cell:
`{"C7": "=SUM($E$7:$P$7)"}`. A large range is cut off and reports a `next_range` to
continue from; read that rather than guessing what is below.

## Reviewing a model

Work in this order, because each step explains the next:

1. `read_range` on the used range — the shape of the sheet and its formulas.
2. `list_names` — models refer to `Purchase_Price`, not `Inputs!$C$5`. Read the names
   before trying to interpret formulas that use them.
3. `audit_formulas` — the mechanical checks, described below.
4. `read_formats` — color coding carries meaning (see conventions). Ask the user to save
   first: on a saved file gridski reads styles straight from the `.xlsx`, which is fast,
   covers whole sheets, and includes borders. With unsaved changes it has to ask Excel
   property by property, which is slow, may cover only part of the range (check
   `unread`), and omits borders.

### Reading audit_formulas

`inconsistent_rows` lists rows of 3+ formulas whose formula changes along the row.
Comparison is in R1C1, so a formula copied correctly across every period reads
identically; anything that differs broke the pattern. Each row comes back as `runs` —
consecutive cells sharing a pattern, with the first cell's formula — plus `constants`,
typed values sitting between formulas. Rows broken in exactly the same way (rows copied
down from one another) are folded into one finding, with the others listed in
`same_rows`.

Each list is capped by `max_findings` (default 50). `inconsistent_rows_total` and
`embedded_numbers_total` give the full counts, and `truncated` is set when something was
left out: audit a narrower `range` (one block of a stacked sheet at a time) rather than
raising the cap by much.

**These are leads, not defects.** A first period that differs from the rest is usually
legitimate (an opening balance, a stub period). A run that changes two thirds of the way
across a row usually is not. Read the actual formulas before reporting anything, and say
which cells you checked.

`embedded_numbers` lists formulas with numbers typed into them, grouped by pattern.
0, 1 and 12 are treated as structural and ignored. A `*1.03` buried in a formula is the
finding worth surfacing: it is an assumption the user cannot see or change from the
inputs sheet.

## Building a period sheet

`build_rows` writes a whole block of time-series rows in one call: labels, units, a
total column, one first-period formula per row copied across every period, plus number
formats and styles. Prefer it over `write_range` for anything periodic — it is one
Excel round trip instead of dozens, and it keeps every row's formula consistent by
construction, so it cannot produce the errors `audit_formulas` looks for.

### Layout

The default layout, which the defaults assume:

| A | B | C | D | E … |
|---|---|---|---|---|
| label | units | total | *(blank spacer)* | periods |

**Leave the spacer column blank.** `{prev:key}` in the first period points at the column
before it, so a blank spacer makes an opening balance read 0 instead of `#REF!` or, worse,
the total. The spacer also stops `audit_formulas` from flagging the total column as a
break in the row.

### Referring to other rows

Give a row a `key` and other rows refer to it, wherever they sit in the block:

| Token | Resolves to | Use for |
|---|---|---|
| `{key}` | that row, same period | `={open}+{draw}` |
| `{prev:key}` | that row, previous period | `={prev:close}` for an opening balance |
| `{row:key}` | the row's whole period range, absolute | `=SUM({row:rent})` in the total column |
| `{total:key}` | the row's total cell, absolute | `=MAX({row:close})/{total:draw}` |

Other sheets are referenced directly in the first-period formula, with the row locked:
`=Amount*Timing!E$9`. Array constants like `{1,2}` pass through untouched. An unknown key
is an error, so the block either resolves completely or is not written.

### Styles

`style` applies the conventions below: `header` bolds the label, `total` bolds the row and
adds a thin top border, `input` colors the total and period cells blue, `link` colors the
period cells green. Set `number_format` per row (e.g. `"#,##0;(#,##0);-"`, `"0.0%"`);
it covers the total and period cells.

```json
{
  "start_row": 5, "first_period": "E", "last_period": "P",
  "rows": [
    {"label": "Loan", "style": "header"},
    {"key": "open",  "label": "Opening balance", "first": "={prev:close}", "number_format": "#,##0"},
    {"key": "draw",  "label": "Draw", "units": "$", "first": "=Amount*Timing!E$9",
     "total": "=SUM({row:draw})", "number_format": "#,##0"},
    {"key": "close", "label": "Closing balance", "first": "={open}+{draw}",
     "style": "total", "number_format": "#,##0"}
  ]
}
```

`build_rows` refuses to write over non-empty cells unless `overwrite` is true — when it
refuses, read the block before overriding. Limits: 500 rows and 60,000 cells per call;
split larger blocks by section.

### Everything else

- `write_range` for one-off blocks, then `fill_right` to copy a column across periods
  (the same relative-reference behavior as Ctrl+R, e.g. write `E9:E60`, then fill
  `E9:EG60`).
- `manage_sheet` to add, rename, move or delete sheets; `define_name` for inputs the
  model should refer to by name.
- `format_range` for formatting outside `build_rows`. Set only what changes.

## Conventions

The color conventions gridski encodes, and the ones to follow when reviewing or writing:

- **Blue** — a hardcoded input, a number someone is meant to change.
- **Black** — a formula.
- **Green** — a link to another sheet.

So blue in the middle of a calculation block is a plug worth asking about, and a black
cell that should be an input is an assumption nobody can find. When a workbook clearly
uses different conventions, follow the workbook's and say so.

Keep inputs on their own sheet, give them defined names, and refer to the name in
formulas rather than repeating the cell address.

## Files and Excel's sandbox

Files are opened and saved only inside the workspace folder (`GRIDSKI_ROOT`, default
`~/gridski`); pass paths relative to it. Anything outside is refused before Excel sees it.

Excel asks the user for access to any file it did not create itself, even inside the
workspace, with a modal dialog that **blocks every gridski call until someone answers
it**. So for a workbook the user downloaded or copied in: `open_workbook` it once, then
`save_workbook` it to a new path in the workspace and work from that copy, which never
prompts again.

If a call reports that Excel is busy or did not respond, a dialog is probably open on
screen. Ask the user to look at Excel rather than retrying.

A save-as into a new folder can be slow. `save_workbook` waits for the file and reports
`saved late` when it lands; if it still errors, Excel may already show the workbook under
its new name with the file yet to appear, so check for the file before saving again.
Excel can't hold two workbooks with the same file name (`source.xlsx` from two folders):
close one first.

## Not available

There are no tools for charts, pivot tables, conditional formatting, inserting or
deleting rows and columns, or running macros (workbooks open with macros disabled).
Say so rather than approximating — and never simulate a missing tool by writing values
where a formula belongs.
