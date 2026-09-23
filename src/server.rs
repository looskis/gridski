//! The MCP tool surface.

use std::path::{Component, PathBuf};
use std::sync::Arc;
use std::time::Duration;

use rmcp::{
    ErrorData as McpError, ServerHandler,
    handler::server::{router::tool::ToolRouter, wrapper::Parameters},
    model::{CallToolResult, ContentBlock, Implementation, ServerCapabilities, ServerConfig},
    schemars, tool, tool_handler, tool_router,
};
use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::bridge::{
    BridgeError, ExcelBridge, FormatSource, LiveFormats, RawFormulas, ReadRequest, Script, SelectionRequest, SheetRef,
    WorkbookInfo, WriteRequest,
};
use crate::formats::{self, FormatsOutput};
use crate::{audit, cells, xlsx};

const DEFAULT_MAX_CELLS: usize = 2_000;
const MAX_READ_CELLS: usize = 20_000;
const MAX_WRITE_CELLS: usize = 10_000;
/// Leaves headroom under the bridge's 30 s timeout for the live format reader.
const LIVE_FORMAT_BUDGET_MS: u64 = 20_000;
const MAX_FORMAT_BLOCKS: usize = 2_000;
const DEFAULT_MAX_FINDINGS: usize = 50;
const MAX_FINDINGS: usize = 1_000;
/// How long save_workbook keeps watching for the file after Excel stops responding.
const LATE_SAVE_WINDOW: Duration = Duration::from_secs(60);
const LATE_SAVE_POLL: Duration = Duration::from_millis(500);
const LATE_SAVE_CONFIRM: Duration = Duration::from_secs(10);

const INSTRUCTIONS: &str = "Reads, audits, and builds workbooks in Microsoft Excel on this Mac, \
which must already be running. Start with get_selection to see what the user is looking at, \
or list_workbooks to see what's open; workbook and sheet default to whatever is active. \
Files are opened and saved only inside the workspace folder; pass paths relative to it. \
Edits change the user's live workbook and cannot be undone with Excel's Undo; \
write_range and fill_right return the previous contents so you can restore them. \
The excel-models skill covers reviewing and building models with these tools.";

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ReadRangeParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    /// Worksheet name. Defaults to the workbook's active sheet.
    pub sheet: Option<String>,
    /// A1-style range such as "A1:D20", a whole column "B:B", or a defined name. Defaults to the sheet's used range.
    pub range: Option<String>,
    /// Maximum number of cells to return (default 2000, max 20000). Larger ranges are cut off by rows, and the result says where to continue.
    pub max_cells: Option<usize>,
    /// Also return formulas for cells that have them (default true).
    pub include_formulas: Option<bool>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SelectionParams {
    /// Maximum number of cells to return (default 2000, max 20000).
    pub max_cells: Option<usize>,
    /// Also return formulas for cells that have them (default true).
    pub include_formulas: Option<bool>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct WriteRangeParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    /// Worksheet name. Defaults to the workbook's active sheet.
    pub sheet: Option<String>,
    /// Top-left cell to start writing at, e.g. "B2".
    pub start: String,
    /// Rectangular grid of rows. Each cell is a number, string, boolean, or null (clears the cell). Strings starting with "=" are entered as formulas; other text is read as if typed ("1/0" becomes a date), so prefix ' to keep it as text.
    pub values: Vec<Vec<Value>>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct OpenWorkbookParams {
    /// Path of an .xlsx/.xlsm file, relative to the workspace folder (or absolute inside it). Omit to create a new blank workbook.
    pub path: Option<String>,
    /// Open the file read-only (default false).
    pub read_only: Option<bool>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SaveWorkbookParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    /// Path to save as (.xlsx, .xlsm, or .xlsb), relative to the workspace folder (or absolute inside it). Omit to save in place.
    pub path: Option<String>,
    /// Replace an existing file at path (default false).
    pub overwrite: Option<bool>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CloseWorkbookParams {
    /// Workbook name as shown by list_workbooks.
    pub workbook: String,
    /// true saves first; false discards unsaved changes.
    pub save: bool,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "lowercase")]
pub enum SheetAction {
    Add,
    Rename,
    Move,
    Delete,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
pub struct ManageSheetParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    pub action: SheetAction,
    /// The sheet to rename, move, or delete.
    pub sheet: Option<String>,
    /// New name, for add (optional) and rename.
    pub name: Option<String>,
    /// For add and move: place before this sheet.
    pub before: Option<String>,
    /// For add and move: place after this sheet. Add defaults to the end.
    pub after: Option<String>,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
pub struct WorkbookParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DefineNameParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    /// Workbook-level name, e.g. "Purchase_Price".
    pub name: String,
    /// What it refers to, e.g. "=Inputs!$C$5". Required unless delete is true.
    pub refers_to: Option<String>,
    /// Remove the name instead (default false).
    pub delete: Option<bool>,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
pub struct SheetRangeParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    /// Worksheet name. Defaults to the workbook's active sheet.
    pub sheet: Option<String>,
    /// A1-style range. Defaults to the sheet's used range.
    pub range: Option<String>,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "lowercase")]
pub enum CalcMode {
    Automatic,
    Manual,
    Semiautomatic,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
pub struct CalculationParams {
    /// Calculation mode; leave unset to keep the current one.
    pub mode: Option<CalcMode>,
    /// Resolve circular references by iterating. A model with deliberate circularity
    /// (interest on average balances) needs this on, or Excel warns instead of solving.
    pub iterative: Option<bool>,
    /// Iterations per calculation (Excel's default is 100).
    pub max_iterations: Option<i32>,
    /// Stop when every value moves less than this (Excel's default is 0.001).
    pub max_change: Option<f64>,
    /// Force a full recalculation of all open workbooks.
    pub recalculate: Option<bool>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct BuildRowsParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    /// Worksheet name. Defaults to the workbook's active sheet.
    pub sheet: Option<String>,
    /// Row where the first spec is written; specs go on consecutive rows.
    pub start_row: u32,
    /// Column of the first period, e.g. "E" (leave a blank spacer column before it so {prev:key} reads 0 in period 1).
    pub first_period: String,
    /// Column of the last period, e.g. "EG".
    pub last_period: String,
    /// Label column (default "A").
    pub label_column: Option<String>,
    /// Units column (default "B").
    pub units_column: Option<String>,
    /// Total/constant column (default "C").
    pub total_column: Option<String>,
    pub rows: Vec<crate::rows::RowSpec>,
    /// Replace non-empty cells in the block (default false: refuse).
    pub overwrite: Option<bool>,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
pub struct FillRightParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    /// Worksheet name. Defaults to the workbook's active sheet.
    pub sheet: Option<String>,
    /// A1-style range at least two columns wide, e.g. "E9:EG60"; its first column is copied across.
    pub range: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct AuditFormulasParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    /// Worksheet name. Defaults to the workbook's active sheet.
    pub sheet: Option<String>,
    /// A1-style range. Defaults to the sheet's used range.
    pub range: Option<String>,
    /// Maximum cells to examine (default and max 20000).
    pub max_cells: Option<usize>,
    /// Most inconsistent-row groups and embedded-number groups to list, each (default 50, max 1000).
    pub max_findings: Option<usize>,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "lowercase")]
pub enum Alignment {
    Left,
    Center,
    Right,
    General,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "lowercase")]
pub enum LineStyle {
    Thin,
    Medium,
    Thick,
    Double,
    Dashed,
    None,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
pub struct BorderSpec {
    pub top: Option<LineStyle>,
    pub bottom: Option<LineStyle>,
    pub left: Option<LineStyle>,
    pub right: Option<LineStyle>,
    pub inside_horizontal: Option<LineStyle>,
    pub inside_vertical: Option<LineStyle>,
}

#[derive(Debug, Serialize, Deserialize, schemars::JsonSchema)]
pub struct FormatRangeParams {
    /// Workbook name as shown by list_workbooks. Defaults to the active workbook.
    pub workbook: Option<String>,
    /// Worksheet name. Defaults to the workbook's active sheet.
    pub sheet: Option<String>,
    /// A1-style range, whole columns ("B:D") for column_width.
    pub range: String,
    /// Font color as "#RRGGBB", e.g. "#0000FF" for inputs.
    pub font_color: Option<String>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    /// Fill color as "#RRGGBB", or "none" to clear.
    pub fill: Option<String>,
    /// Excel number format code, e.g. "#,##0;(#,##0);-", "0.0%", "mmm-yy".
    pub number_format: Option<String>,
    pub horizontal_alignment: Option<Alignment>,
    /// Indent level (0–15).
    pub indent: Option<u32>,
    /// Width of the range's columns, in characters.
    pub column_width: Option<f64>,
    /// Edges of the range to set; "none" removes a line.
    pub borders: Option<BorderSpec>,
}

#[derive(Debug, Serialize)]
struct WriteOutput {
    workbook: String,
    sheet: String,
    address: String,
    /// Pass these back to write_range at the same start cell to undo this write.
    previous_formulas: Vec<Vec<Value>>,
}

#[derive(Clone)]
pub struct GridskiServer<B: ExcelBridge> {
    bridge: Arc<B>,
    /// The one folder files are opened from and saved to (see [`GridskiServer::resolve`]).
    root: Arc<PathBuf>,
    tool_router: ToolRouter<Self>,
}

#[tool_router]
impl<B: ExcelBridge> GridskiServer<B> {
    pub fn new(bridge: B, root: PathBuf) -> Self {
        Self { bridge: Arc::new(bridge), root: Arc::new(root), tool_router: Self::tool_router() }
    }

    #[tool(
        description = "List the workbooks open in Excel with their sheets and used ranges, and which ones are active.",
        annotations(read_only_hint = true)
    )]
    async fn list_workbooks(&self) -> Result<CallToolResult, McpError> {
        respond(self.bridge.list_workbooks().await)
    }

    #[tool(
        description = "Get the cells the user currently has selected in Excel: workbook, sheet, address, values, and formulas. The best way to see what the user is working on.",
        annotations(read_only_hint = true)
    )]
    async fn get_selection(&self, Parameters(p): Parameters<SelectionParams>) -> Result<CallToolResult, McpError> {
        let req = SelectionRequest {
            max_cells: clamp_cells(p.max_cells),
            include_formulas: p.include_formulas.unwrap_or(true),
        };
        respond(self.bridge.get_selection(req).await.map(cells::normalize))
    }

    #[tool(
        description = "Read values (and optionally formulas) from a range in an open workbook. Empty cells are null, dates are YYYY-MM-DD, and errors appear as codes like #DIV/0!.",
        annotations(read_only_hint = true)
    )]
    async fn read_range(&self, Parameters(p): Parameters<ReadRangeParams>) -> Result<CallToolResult, McpError> {
        let req = ReadRequest {
            target: SheetRef { workbook: p.workbook, sheet: p.sheet },
            range: p.range,
            max_cells: clamp_cells(p.max_cells),
            include_formulas: p.include_formulas.unwrap_or(true),
        };
        respond(self.bridge.read_range(req).await.map(cells::normalize))
    }

    #[tool(
        description = "Write a grid of values and formulas into an open workbook, starting at a cell. This edits the user's live workbook and cannot be undone with Excel's Undo; the result includes the previous contents so the write can be reverted.",
        annotations(read_only_hint = false, destructive_hint = true, idempotent_hint = true)
    )]
    async fn write_range(&self, Parameters(p): Parameters<WriteRangeParams>) -> Result<CallToolResult, McpError> {
        if let Err(msg) = validate_grid(&p.values) {
            return respond::<()>(Err(BridgeError::InvalidInput(msg)));
        }
        let req = WriteRequest {
            target: SheetRef { workbook: p.workbook, sheet: p.sheet },
            start: p.start,
            values: p.values,
        };
        respond(self.bridge.write_range(req).await.map(|w| WriteOutput {
            workbook: w.workbook,
            sheet: w.sheet,
            address: w.address.replace('$', ""),
            previous_formulas: w.previous_formulas,
        }))
    }

    #[tool(
        description = "Open an .xlsx/.xlsm file in Excel, or create a new blank workbook when no path is given. Returns the workbook's name (use it as `workbook` in other tools) and sheets. A file that is already open is returned as is; a different open workbook with the same file name is refused, since Excel can't hold both. Macros are disabled and external links are not updated, so neither prompts. Excel asks once for access to a file it did not create; save_workbook a copy into the workspace to avoid asking again.",
        annotations(read_only_hint = false, destructive_hint = false, idempotent_hint = true)
    )]
    async fn open_workbook(&self, Parameters(p): Parameters<OpenWorkbookParams>) -> Result<CallToolResult, McpError> {
        let path = match p.path.as_deref().map(|path| self.resolve(path)).transpose() {
            Ok(path) => path,
            Err(e) => return respond::<()>(Err(e)),
        };
        let args = serde_json::json!({ "path": path, "read_only": p.read_only.unwrap_or(false) });
        respond(self.bridge.call::<_, WorkbookInfo>(Script::OpenWorkbook, &args).await)
    }

    #[tool(
        description = "Save a workbook in place, or save it as a new file when path is given (the workbook then takes the new file's name). Creates missing folders; refuses to replace an existing file unless overwrite is true, and refuses a file name another open workbook already has. Excel can only write where it has folder access (your home folder, not /tmp); the tool checks the file was actually written, and if Excel is slow it keeps watching for up to 60 s and reports a late save in note.",
        annotations(read_only_hint = false, destructive_hint = false, idempotent_hint = true)
    )]
    async fn save_workbook(&self, Parameters(p): Parameters<SaveWorkbookParams>) -> Result<CallToolResult, McpError> {
        let path = match p.path.as_deref().map(|path| self.resolve(path)).transpose() {
            Ok(path) => path,
            Err(e) => return respond::<()>(Err(e)),
        };
        // Excel can't create folders: a save-as into a missing one fails with "Parameter error"
        // yet still renames the workbook to the unwritten path. The path is inside the workspace.
        if let Some(parent) = path.as_deref().and_then(|p| std::path::Path::new(p).parent())
            && let Err(e) = std::fs::create_dir_all(parent)
        {
            return respond::<()>(Err(BridgeError::InvalidInput(format!("Could not create {}: {e}", parent.display()))));
        }
        // Where the file will land, so a save that outlives the bridge's timeout can still be confirmed.
        let target = match &path {
            Some(path) => Some(path.clone()),
            None => self.workbook_path(p.workbook.as_deref()).await,
        };
        let before = target.as_deref().and_then(file_stamp);
        let args = serde_json::json!({
            "workbook": p.workbook,
            "path": path,
            "root": self.root.to_string_lossy(),
            "overwrite": p.overwrite.unwrap_or(false),
        });
        let save_as = path.is_some();
        match (self.bridge.call::<_, Value>(Script::SaveWorkbook, &args).await, target) {
            (Err(BridgeError::Busy), Some(target)) => respond(self.late_save(target, before).await),
            // The file of a save-as can appear after Excel reports the save done.
            (Err(BridgeError::NotWritten(_)), Some(target)) if save_as => respond(self.late_save(target, before).await),
            (result, _) => respond(result),
        }
    }

    #[tool(
        description = "Close a workbook. save: true saves it first; save: false discards unsaved changes. Errors if the workbook is still open afterwards (Excel busy or showing a dialog).",
        annotations(read_only_hint = false, destructive_hint = true, idempotent_hint = false)
    )]
    async fn close_workbook(&self, Parameters(p): Parameters<CloseWorkbookParams>) -> Result<CallToolResult, McpError> {
        let args = serde_json::json!({ "workbook": p.workbook, "save": p.save });
        respond(self.bridge.call::<_, Value>(Script::CloseWorkbook, &args).await.map_err(|e| match e {
            BridgeError::Busy => BridgeError::Unsupported(format!(
                "Excel did not respond while closing \"{}\", so it may still be open; it is probably busy or \
                 showing a dialog. Ask the user to check Excel, then confirm with list_workbooks.",
                p.workbook
            )),
            other => other,
        }))
    }

    #[tool(
        description = "Add, rename, move, or delete a worksheet. add takes an optional name and before/after (default: end); rename takes sheet and name; move takes sheet and before or after; delete takes sheet and removes it with its contents. Returns the workbook's sheets in order.",
        annotations(read_only_hint = false, destructive_hint = true, idempotent_hint = false)
    )]
    async fn manage_sheet(&self, Parameters(p): Parameters<ManageSheetParams>) -> Result<CallToolResult, McpError> {
        respond(self.bridge.call::<_, Value>(Script::ManageSheet, &p).await)
    }

    #[tool(
        description = "List a workbook's defined names and what each refers to (e.g. Purchase_Price → ='Inputs'!$C$5). Sheet-scoped names appear as 'Sheet'!Name. Read these before auditing a model that uses names in its formulas.",
        annotations(read_only_hint = true)
    )]
    async fn list_names(&self, Parameters(p): Parameters<WorkbookParams>) -> Result<CallToolResult, McpError> {
        respond(self.bridge.call::<_, Value>(Script::ListNames, &p).await)
    }

    #[tool(
        description = "Create or change a workbook-level defined name (refers_to, e.g. \"=Inputs!$C$5\"), or remove one (delete: true). Returns the previous reference.",
        annotations(read_only_hint = false, destructive_hint = true, idempotent_hint = true)
    )]
    async fn define_name(&self, Parameters(p): Parameters<DefineNameParams>) -> Result<CallToolResult, McpError> {
        let delete = p.delete.unwrap_or(false);
        if !delete && p.refers_to.is_none() {
            return respond::<()>(Err(BridgeError::InvalidInput("Pass refers_to, or delete: true to remove the name.".into())));
        }
        let args = serde_json::json!({
            "workbook": p.workbook,
            "name": p.name,
            "refers_to": if delete { None } else { p.refers_to },
        });
        respond(self.bridge.call::<_, Value>(Script::DefineName, &args).await)
    }

    #[tool(
        description = "Read cell formatting as rectangles of identically formatted cells: font_color, bold, italic, fill, number_format, and borders. Only non-default properties are listed, and cells with default formatting (black, regular, no or white fill, General, no borders) are left out. Reads the saved file when the workbook has no unsaved changes (fast, whole sheets); otherwise asks Excel directly, which is slow, may cover only part of the range (see unread), and omits borders. Save first for complete results.",
        annotations(read_only_hint = true)
    )]
    async fn read_formats(&self, Parameters(p): Parameters<SheetRangeParams>) -> Result<CallToolResult, McpError> {
        respond(self.read_formats_inner(p).await)
    }

    #[tool(
        description = "Format a range: font_color, bold, italic, fill, number_format, horizontal_alignment, indent, column_width, and borders on its edges. Set only what should change. Not undoable.",
        annotations(read_only_hint = false, destructive_hint = false, idempotent_hint = true)
    )]
    async fn format_range(&self, Parameters(p): Parameters<FormatRangeParams>) -> Result<CallToolResult, McpError> {
        respond(self.bridge.call::<_, Value>(Script::FormatRange, &p).await)
    }

    #[tool(
        description = "Read or change Excel's calculation settings, which are application-wide: mode (automatic/manual/semiautomatic), iterative calculation for circular references with its iteration and change limits, and recalculate to force a full recalculation. Pass nothing to read the current settings. Turn iterative on before building a model whose circularity is deliberate (interest on average balances); it is off by default and Excel then shows a circular-reference warning instead of solving.",
        annotations(read_only_hint = false, destructive_hint = false, idempotent_hint = true)
    )]
    async fn calculation(&self, Parameters(p): Parameters<CalculationParams>) -> Result<CallToolResult, McpError> {
        respond(self.bridge.call::<_, Value>(Script::Calculation, &p).await)
    }

    #[tool(
        description = "Build a block of time-series rows in one call: per row a label, units, a total/constant cell, and a first-period formula that is copied across every period to last_period (relative references adjust, like Fill Right), plus number_format and style (header, total, link, input). Rows refer to each other by key: {key} = that row in the same period, {prev:key} = previous period (e.g. opening = {prev:close}), {row:key} = the row's whole period range (for totals: =SUM({row:rent})), {total:key} = its total cell. Other sheets are referenced directly, e.g. Timing!E$9 in the first period. Refuses to write over non-empty cells unless overwrite is true. Returns each key's row number.",
        annotations(read_only_hint = false, destructive_hint = true, idempotent_hint = true)
    )]
    async fn build_rows(&self, Parameters(p): Parameters<BuildRowsParams>) -> Result<CallToolResult, McpError> {
        let col = |c: &Option<String>, default: &str| crate::rows::column(c.as_deref().unwrap_or(default));
        let layout = (|| {
            Ok::<_, String>(crate::rows::Layout {
                start_row: p.start_row.max(1),
                label_col: col(&p.label_column, "A")?,
                units_col: col(&p.units_column, "B")?,
                total_col: col(&p.total_column, "C")?,
                first_col: crate::rows::column(&p.first_period)?,
                last_col: crate::rows::column(&p.last_period)?,
            })
        })();
        let plan = match layout.and_then(|l| crate::rows::plan(&l, &p.rows)) {
            Ok(plan) => plan,
            Err(e) => return respond::<()>(Err(BridgeError::InvalidInput(e))),
        };
        let args = serde_json::json!({
            "workbook": p.workbook,
            "sheet": p.sheet,
            "overwrite": p.overwrite.unwrap_or(false),
            "target": plan.target,
            "head_start": plan.head_start,
            "head": plan.head,
            "fill": plan.fill,
            "first_col": plan.first_col,
            "formats": plan.formats,
            "keys": plan.keys,
        });
        respond(self.bridge.call::<_, Value>(Script::BuildRows, &args).await)
    }

    #[tool(
        description = "Fill Right, like Excel's Ctrl+R: copy each row's first cell across the rest of the range, adjusting relative references (e.g. write E9:E60 with write_range, then fill_right E9:EG60). Copies formulas and values, including blanks; formats are unchanged. Not undoable; returns previous contents of overwritten cells (up to 500).",
        annotations(read_only_hint = false, destructive_hint = true, idempotent_hint = true)
    )]
    async fn fill_right(&self, Parameters(p): Parameters<FillRightParams>) -> Result<CallToolResult, McpError> {
        respond(self.bridge.call::<_, Value>(Script::FillRight, &p).await)
    }

    #[tool(
        description = "Audit a sheet's formulas. inconsistent_rows: rows with 3+ formulas whose formula changes along the row (compared in R1C1, so a correctly copied formula matches in every column) or with typed values between formulas; a blank cell splits a row into blocks judged separately, so a totals column set off by a blank spacer is not flagged; each lists its runs with the first cell's formula. Rows with identical runs (same columns and R1C1 patterns) are folded into one finding's same_rows. embedded_numbers: formulas containing typed numbers other than 0, 1 and 12, grouped by pattern. Each list is capped at max_findings (default 50); *_total give the full counts and truncated is set when anything was left out. A different first-period formula is often legitimate; judge each finding.",
        annotations(read_only_hint = true)
    )]
    async fn audit_formulas(&self, Parameters(p): Parameters<AuditFormulasParams>) -> Result<CallToolResult, McpError> {
        let args = serde_json::json!({
            "workbook": p.workbook,
            "sheet": p.sheet,
            "range": p.range,
            "max_cells": p.max_cells.unwrap_or(MAX_READ_CELLS).clamp(1, MAX_READ_CELLS),
        });
        let max_findings = p.max_findings.unwrap_or(DEFAULT_MAX_FINDINGS).clamp(1, MAX_FINDINGS);
        let result = self.bridge.call::<_, RawFormulas>(Script::AuditFormulas, &args).await.map(|raw| {
            let mut out = audit::audit(raw.workbook, raw.sheet, &raw.address, &raw.formulas, &raw.r1c1, max_findings);
            let returned = raw.formulas.len() * raw.formulas.first().map_or(0, Vec::len);
            if returned < raw.total_rows * raw.total_cols {
                let cells = format!(
                    "Only {} of {}x{} cells were examined; audit the rest with a narrower range.",
                    out.address, raw.total_rows, raw.total_cols
                );
                out.note = Some(match out.note.take() {
                    Some(findings) => format!("{cells} {findings}"),
                    None => cells,
                });
                out.truncated = true;
            }
            out
        });
        respond(result)
    }
}

impl<B: ExcelBridge> GridskiServer<B> {
    /// Resolve a tool's path against the workspace root and refuse anything outside it.
    /// Excel is sandboxed and asks for access to each new folder with a modal dialog
    /// (which blocks every later Apple Event), so keep it inside the one granted folder.
    fn resolve(&self, path: &str) -> Result<String, BridgeError> {
        let joined = self.root.join(path);
        let mut normal = PathBuf::new();
        for part in joined.components() {
            match part {
                Component::ParentDir => {
                    normal.pop();
                }
                Component::CurDir => {}
                other => normal.push(other),
            }
        }
        if !normal.starts_with(self.root.as_path()) {
            return Err(BridgeError::InvalidInput(format!(
                "{} is outside the workspace folder {}; Excel can only open and save files inside it. \
                 Use a path relative to the workspace.",
                normal.display(),
                self.root.display()
            )));
        }
        Ok(normal.to_string_lossy().into_owned())
    }

    /// The saved file behind a workbook (by name, or the active one), if it has one.
    async fn workbook_path(&self, workbook: Option<&str>) -> Option<String> {
        let books = self.bridge.list_workbooks().await.ok()?;
        let book = books.into_iter().find(|b| workbook.map_or(b.active, |name| b.name == name))?;
        book.path.starts_with('/').then_some(book.path)
    }

    /// Excel sometimes finishes a save well after the bridge gave up on it (a first save into a
    /// new folder, a large workbook). Wait for the file to land before calling it a failure.
    async fn late_save(&self, target: String, before: Option<FileStamp>) -> Result<Value, BridgeError> {
        tracing::info!(%target, "save not confirmed; watching for the file");
        if !wait_for_write(&target, before, LATE_SAVE_WINDOW, LATE_SAVE_POLL).await {
            let waited = LATE_SAVE_WINDOW.as_secs();
            // Excel can rename the workbook to the new path (and call it saved) well before the
            // file appears; it has been seen landing minutes later.
            return Err(BridgeError::Unsupported(match self.open_at(&target).await {
                Some(book) => format!(
                    "Excel did not confirm the save. It already shows the workbook as \"{}\" at {target}, but the \
                     file had not appeared after a further {waited} s; Excel may still be writing it. Check for the \
                     file before saving again, and don't save it under another name meanwhile.",
                    book.name
                ),
                None => format!(
                    "Excel did not confirm the save, and {target} was not written within a further {waited} s. \
                     Excel may be showing a dialog (such as a file-access prompt), be stuck, or lack access to \
                     the folder; ask the user to check Excel."
                ),
            }));
        }
        let name = target.rsplit('/').next().unwrap_or(&target).to_string();
        // Best effort: Excel is usually responsive again once the file is down.
        Ok(match self.open_at(&target).await {
            Some(book) => serde_json::json!({
                "workbook": book.name,
                "path": book.path,
                "saved": book.saved,
                "note": "saved late (Excel was slow to respond)",
            }),
            None => serde_json::json!({
                "workbook": name,
                "path": target,
                "note": "saved late (Excel was slow to respond); the file is written, but Excel has not yet \
                         confirmed the workbook under its new name — check list_workbooks",
            }),
        })
    }

    /// The open workbook whose file is `path`, if Excel answers promptly.
    async fn open_at(&self, path: &str) -> Option<WorkbookInfo> {
        let books = tokio::time::timeout(LATE_SAVE_CONFIRM, self.bridge.list_workbooks()).await.ok()?.ok()?;
        books.into_iter().find(|b| b.path == path)
    }

    async fn read_formats_inner(&self, p: SheetRangeParams) -> Result<FormatsOutput, BridgeError> {
        let src: FormatSource = self.bridge.call(Script::FormatSource, &p).await?;
        let address = cells::strip_dollars(&src.address);
        let from_file = src.saved && [".xlsx", ".xlsm"].iter().any(|ext| src.path.to_lowercase().ends_with(ext));
        if let (true, Some(bounds)) = (from_file, cells::bounds(&address)) {
            let (path, sheet) = (src.path.clone(), src.sheet.clone());
            let read = tokio::task::spawn_blocking(move || xlsx::read_cell_formats(&path, &sheet, bounds)).await;
            match read {
                Ok(Ok(cells)) => {
                    return Ok(capped(FormatsOutput {
                        workbook: src.workbook,
                        sheet: src.sheet,
                        address,
                        source: "file",
                        blocks: formats::blocks_from_cells(cells),
                        unread: Vec::new(),
                    }));
                }
                Ok(Err(e)) => tracing::warn!(error = %e, "falling back to live format reads"),
                Err(e) => tracing::warn!(error = %e, "format reader panicked; falling back to live reads"),
            }
        }
        let args = serde_json::json!({
            "workbook": src.workbook,
            "sheet": src.sheet,
            "range": address,
            "budget_ms": LIVE_FORMAT_BUDGET_MS,
        });
        let live: LiveFormats = self.bridge.call(Script::ReadFormatsLive, &args).await?;
        Ok(capped(FormatsOutput {
            workbook: live.workbook,
            sheet: live.sheet,
            address: cells::strip_dollars(&live.address),
            source: "live",
            blocks: formats::blocks_from_live(live.blocks),
            unread: live.unread.iter().map(|a| cells::strip_dollars(a)).collect(),
        }))
    }
}

/// A file's modification time and size, to tell whether a save rewrote it.
type FileStamp = (std::time::SystemTime, u64);

fn file_stamp(path: &str) -> Option<FileStamp> {
    let meta = std::fs::metadata(path).ok()?;
    Some((meta.modified().ok()?, meta.len()))
}

/// Poll until `path` holds a non-empty file that differs from `before` and has stopped
/// changing between two polls, or until `window` runs out.
async fn wait_for_write(path: &str, before: Option<FileStamp>, window: Duration, poll: Duration) -> bool {
    let deadline = tokio::time::Instant::now() + window;
    let mut last: Option<FileStamp> = None;
    loop {
        let now = file_stamp(path).filter(|stamp| stamp.1 > 0 && Some(*stamp) != before);
        if now.is_some() && now == last {
            return true;
        }
        last = now;
        if tokio::time::Instant::now() >= deadline {
            return false;
        }
        tokio::time::sleep(poll).await;
    }
}

/// Keep huge, heterogeneous sheets from flooding the result.
fn capped(mut out: FormatsOutput) -> FormatsOutput {
    if out.blocks.len() > MAX_FORMAT_BLOCKS {
        out.blocks.truncate(MAX_FORMAT_BLOCKS);
        out.unread.push(format!("more than {MAX_FORMAT_BLOCKS} blocks; read a narrower range"));
    }
    out
}

#[tool_handler(router = self.tool_router)]
impl<B: ExcelBridge> ServerHandler for GridskiServer<B> {
    fn get_info(&self) -> ServerConfig {
        ServerConfig::new(ServerCapabilities::builder().enable_tools().build())
            .with_server_info(Implementation::new("gridski", env!("CARGO_PKG_VERSION")))
            .with_instructions(format!("{INSTRUCTIONS} Workspace folder: {}.", self.root.display()))
    }
}

/// Excel problems (not running, bad range, ...) are tool errors the model can react to,
/// not protocol errors.
fn respond<T: Serialize>(result: Result<T, BridgeError>) -> Result<CallToolResult, McpError> {
    match result {
        Ok(value) => Ok(CallToolResult::success(vec![ContentBlock::json(value)?])),
        Err(err) => {
            tracing::warn!(error = %err, "tool call failed");
            Ok(CallToolResult::error(vec![ContentBlock::text(err.to_string())]))
        }
    }
}

fn clamp_cells(requested: Option<usize>) -> usize {
    requested.unwrap_or(DEFAULT_MAX_CELLS).clamp(1, MAX_READ_CELLS)
}

fn validate_grid(values: &[Vec<Value>]) -> Result<(), String> {
    let cols = values.first().map_or(0, Vec::len);
    if cols == 0 {
        return Err("values must contain at least one row with at least one cell.".into());
    }
    if let Some(i) = values.iter().position(|row| row.len() != cols) {
        return Err(format!("values must be rectangular: row {} has {} cells, expected {cols}.", i + 1, values[i].len()));
    }
    if values.len() * cols > MAX_WRITE_CELLS {
        return Err(format!("Writes are limited to {MAX_WRITE_CELLS} cells; split this into smaller blocks."));
    }
    if values.iter().flatten().any(|v| v.is_array() || v.is_object()) {
        return Err("Each cell must be a number, string, boolean, or null.".into());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn grid_validation() {
        assert!(validate_grid(&[vec![json!(1), json!("=A1")]]).is_ok());
        assert!(validate_grid(&[]).is_err());
        assert!(validate_grid(&[vec![json!(1)], vec![json!(1), json!(2)]]).is_err());
        assert!(validate_grid(&[vec![json!([1])]]).is_err());
    }

    #[test]
    fn paths_stay_in_workspace() {
        struct NoBridge;
        impl ExcelBridge for NoBridge {
            async fn list_workbooks(&self) -> Result<Vec<WorkbookInfo>, BridgeError> {
                unreachable!()
            }
            async fn read_range(&self, _: ReadRequest) -> Result<crate::bridge::RawRange, BridgeError> {
                unreachable!()
            }
            async fn get_selection(&self, _: SelectionRequest) -> Result<crate::bridge::RawRange, BridgeError> {
                unreachable!()
            }
            async fn write_range(&self, _: WriteRequest) -> Result<crate::bridge::WriteResult, BridgeError> {
                unreachable!()
            }
            async fn call<A: Serialize + Sync, T: serde::de::DeserializeOwned>(
                &self,
                _: Script,
                _: &A,
            ) -> Result<T, BridgeError> {
                unreachable!()
            }
        }
        let server = GridskiServer::new(NoBridge, PathBuf::from("/Users/me/work"));
        assert_eq!(server.resolve("a/b.xlsx").unwrap(), "/Users/me/work/a/b.xlsx");
        assert_eq!(server.resolve("/Users/me/work/c.xlsx").unwrap(), "/Users/me/work/c.xlsx");
        assert!(server.resolve("../elsewhere.xlsx").is_err());
        assert!(server.resolve("/private/tmp/x.xlsx").is_err());
        assert!(server.resolve("/Users/me/workshop/x.xlsx").is_err());
    }

    #[tokio::test]
    async fn late_save_waits_for_the_file() {
        let dir = std::env::temp_dir().join(format!("gridski-test-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("late.xlsx");
        let path_str = path.to_string_lossy().into_owned();
        let poll = Duration::from_millis(20);

        // Never written: gives up when the window closes.
        assert!(!wait_for_write(&path_str, None, Duration::from_millis(100), poll).await);

        // Written partway through the window.
        let writer = {
            let path = path.clone();
            tokio::spawn(async move {
                tokio::time::sleep(Duration::from_millis(80)).await;
                std::fs::write(&path, b"PK").unwrap();
            })
        };
        assert!(wait_for_write(&path_str, None, Duration::from_secs(2), poll).await);
        writer.await.unwrap();

        // An existing file that is not rewritten does not count as saved.
        let before = file_stamp(&path_str);
        assert!(!wait_for_write(&path_str, before, Duration::from_millis(100), poll).await);
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn cell_limits_clamp() {
        assert_eq!(clamp_cells(None), DEFAULT_MAX_CELLS);
        assert_eq!(clamp_cells(Some(0)), 1);
        assert_eq!(clamp_cells(Some(1_000_000)), MAX_READ_CELLS);
    }
}
