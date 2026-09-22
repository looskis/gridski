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
| `audit_formulas` | Rows whose formula changes across columns (compared in R1C1), and numbers typed into formulas |
| `read_formats` | Font color, bold/italic, fill, number format, and borders as merged rectangles; reads the saved file when there are no unsaved changes |
| `open_workbook` | Open a file inside the workspace, or create a blank workbook |
| `save_workbook` | Save in place or save as a new file inside the workspace; verifies the file was written |
| `close_workbook` | Close a workbook, saving or discarding changes |
| `manage_sheet` | Add, rename, move, or delete a worksheet |
| `define_name` | Create, change, or remove a workbook-level name |
| `fill_right` | Copy each row's first cell across a range, like Excel's Fill Right; for time-series rows |
| `format_range` | Set font color, bold/italic, fill, number format, alignment, indent, column width, and borders |

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

The first save into the workspace may prompt once: choose the workspace folder itself in the dialog, and the grant covers everything beneath it.

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
