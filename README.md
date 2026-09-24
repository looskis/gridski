<p align="center">
  <img src="assets/gridski-icon.png" alt="gridski logo" width="128" height="128">
</p>

# gridski

An MCP server that lets AI assistants read and edit workbooks open in Microsoft Excel on macOS.

It talks to the running Excel app through Apple Events (JavaScript for Automation via `osascript`), so it sees exactly what the analyst sees — unsaved edits, the current selection, live formulas.

## Tools

| Tool | What it does |
|---|---|
| `list_workbooks` | Open workbooks, their sheets and used ranges, and which are active |
| `get_selection` | The cells the user has selected, with values and formulas |
| `read_range` | Values and formulas from a range (defaults to the active sheet's used range) |
| `write_range` | Write a grid of values/formulas; returns the previous contents so it can be reverted |
| `list_names` | Defined names and what each refers to |
| `audit_formulas` | Rows whose formula changes across columns (compared in R1C1), and numbers typed into formulas; rows broken the same way are folded together, and each list is capped by `max_findings` (default 50) with full totals |
| `read_formats` | Font color, bold/italic, fill, number format, and borders as merged rectangles; reads the saved file when there are no unsaved changes |
| `open_workbook` | Open a file inside the workspace (macros disabled, links not updated), or create a blank workbook; refuses a file name that a different open workbook already has |
| `save_workbook` | Save in place or save as a new file inside the workspace (creating folders); verifies the file was written, waiting up to 60 s more when Excel is slow, and refuses a file name another open workbook has |
| `close_workbook` | Close a workbook, saving or discarding changes; errors if Excel leaves it open |
| `manage_sheet` | Add, rename, move, or delete a worksheet |
| `define_name` | Create, change, or remove a workbook-level name |
| `build_rows` | Build a period sheet's rows in one call: label, units, total, and a first-period formula filled across all periods; rows refer to each other by `{key}`; styles and number formats applied |
| `fill_right` | Copy each row's first cell across a range, like Excel's Fill Right; for time-series rows |
| `calculation` | Read or set calculation mode, iterative calculation for circular models, and force a recalc |
| `format_range` | Set font color, bold/italic, fill, number format, alignment, indent, column width, and borders |

## Skill

`skills/excel-models/SKILL.md` carries the judgment the tool descriptions can't: how to review a
model, how to lay out a period sheet, the color conventions, and Excel's file-access behavior.
Load it into a client alongside the server; the server's own instructions stay limited to routing.

Results are normalized for models: empty cells are `null`, dates are `YYYY-MM-DD`, errors are codes like `#DIV/0!`, formulas are listed sparsely by cell, and large reads are truncated with a `next_range` to continue from.

## Build

```bash
cargo build --release
```

## Use with an MCP client

Point your client at the binary over stdio, e.g. in Claude Desktop's config:

```json
{
  "mcpServers": {
    "gridski": { "command": "/path/to/gridski/target/release/gridski" }
  }
}
```

The first tool call triggers a macOS prompt asking whether the **client app** (Claude, Terminal, …) may control Excel. If it was denied, re-enable it under System Settings → Privacy & Security → Automation.

gridski never launches Excel; it reports an error if Excel isn't running.

### Workspace folder

Excel is sandboxed: the first time it opens or saves in a new folder it shows a modal "Grant File Access" dialog, which blocks every Apple Event until someone answers it, and outside granted folders it can report a save that never reached disk. So gridski opens and saves files only inside one workspace folder, set with `GRIDSKI_ROOT` (default `~/gridski`, created if missing). Paths in tool calls are relative to it; anything outside is refused before Excel sees it.

The first save into the workspace may prompt once: choose the workspace folder itself in the dialog, and the grant covers saving anywhere beneath it. Opening is different: Excel asks once for every file it didn't create itself (a download, a copy made by another program), even inside a granted folder. To ask only once per file, open it and `save_workbook` it to a new path — Excel's own saved copy never prompts again.

Excel can't hold two workbooks with the same file name, even from different folders, so `open_workbook` and `save_workbook` refuse one that's already taken rather than let Excel fail with a bare "Parameter error". A save-as into a folder Excel hasn't written to before can take long enough to outlast the 30 s call timeout, or land after Excel reports it done; `save_workbook` then keeps watching for the file and reports `"note": "saved late …"` once it appears.

```json
{ "gridski": { "command": "/path/to/gridski", "env": { "GRIDSKI_ROOT": "~/models" } } }
``` Set `GRIDSKI_LOG=debug` for verbose logs on stderr.

## Development

```bash
cargo test
```

`scripts/smoke.py` drives the binary end to end against a running Excel. It writes to `Z1000:AA1001` on the active sheet and restores it afterwards.

## Layout

- `src/server.rs` — MCP tools (rmcp)
- `src/bridge/` — the `ExcelBridge` trait and its JXA implementation; `scripts/*.js` run inside Excel
- `src/cells.rs` — turns raw Excel output into model-friendly results
- `src/formats.rs`, `src/xlsx.rs` — format blocks; reading styles straight from a saved `.xlsx`
- `src/audit.rs` — formula-consistency and typed-number checks

Some Excel commands (positioning sheets, creating names, save-as) only work from AppleScript; `prelude.js` runs those in-process through `NSAppleScript`. Format queries over Apple Events cost ~10 ms per property per range, which is why `read_formats` prefers the file.
