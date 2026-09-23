//! Talking to a running copy of Excel.
//!
//! [`ExcelBridge`] is the seam between the MCP tools and however we reach Excel.
//! Today that is [`jxa::JxaBridge`] (spawning `osascript`); an in-process OSAKit
//! implementation can slot in behind the same trait later.

pub mod jxa;

use std::future::Future;

use serde::de::DeserializeOwned;
use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Debug, thiserror::Error)]
pub enum BridgeError {
    #[error("Microsoft Excel is not running. Open Excel and try again.")]
    NotRunning,
    #[error(
        "Not allowed to control Excel. Grant access under System Settings → Privacy & Security → \
         Automation for the app running this MCP server."
    )]
    PermissionDenied,
    #[error(
        "Excel did not respond in time. It may be showing a dialog or busy recalculating; \
         close any open dialog and try again."
    )]
    Busy,
    #[error("{0}")]
    NotFound(String),
    #[error("{0}")]
    Unsupported(String),
    #[error("{0}")]
    InvalidInput(String),
    #[error("Excel scripting error: {0}")]
    Script(String),
    #[error("failed to run osascript: {0}")]
    Io(#[from] std::io::Error),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookInfo {
    pub name: String,
    pub path: String,
    pub saved: bool,
    pub active: bool,
    pub sheets: Vec<SheetInfo>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetInfo {
    pub name: String,
    pub used_range: String,
    pub active: bool,
}

/// Which sheet to act on; `None` means whatever is active in Excel.
#[derive(Debug, Clone, Default, Serialize)]
pub struct SheetRef {
    pub workbook: Option<String>,
    pub sheet: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct ReadRequest {
    #[serde(flatten)]
    pub target: SheetRef,
    /// A1-style address; `None` reads the sheet's used range.
    pub range: Option<String>,
    pub max_cells: usize,
    pub include_formulas: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct SelectionRequest {
    pub max_cells: usize,
    pub include_formulas: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct WriteRequest {
    #[serde(flatten)]
    pub target: SheetRef,
    /// Top-left cell of the block to write.
    pub start: String,
    /// Rectangular grid; strings starting with `=` are formulas, `null` clears a cell.
    pub values: Vec<Vec<Value>>,
}

/// A range as Excel reports it, before normalization (see [`crate::cells`]).
#[derive(Debug, Clone, Deserialize)]
pub struct RawRange {
    pub workbook: String,
    pub sheet: String,
    /// Address of the cells actually returned (may be smaller than requested).
    pub address: String,
    pub total_rows: usize,
    pub total_cols: usize,
    /// `value()`: typed values, but errors come back as `""`.
    pub values: Vec<Vec<Value>>,
    /// `stringValue()`: displayed text, including error codes like `#DIV/0!`.
    pub text: Vec<Vec<Value>>,
    pub formulas: Option<Vec<Vec<Value>>>,
    /// Only set for selections: every area's address.
    #[serde(default)]
    pub areas: Option<Vec<String>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WriteResult {
    pub workbook: String,
    pub sheet: String,
    pub address: String,
    /// What the cells held before the write; writing these back undoes it.
    pub previous_formulas: Vec<Vec<Value>>,
}

/// Scripts reached through [`ExcelBridge::call`], for tools whose arguments and results
/// pass through the bridge without Rust-side shaping.
#[derive(Debug, Clone, Copy)]
pub enum Script {
    OpenWorkbook,
    SaveWorkbook,
    CloseWorkbook,
    ManageSheet,
    ListNames,
    DefineName,
    FormatSource,
    ReadFormatsLive,
    FormatRange,
    AuditFormulas,
    FillRight,
    BuildRows,
    Calculation,
}

/// Where a sheet's formatting can be read from (see `format_source.js`).
#[derive(Debug, Clone, Deserialize)]
pub struct FormatSource {
    pub workbook: String,
    pub sheet: String,
    pub address: String,
    pub path: String,
    pub saved: bool,
}

#[derive(Debug, Clone, Deserialize)]
pub struct LiveFormats {
    pub workbook: String,
    pub sheet: String,
    pub address: String,
    pub blocks: Vec<crate::formats::LiveBlock>,
    pub unread: Vec<String>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct RawFormulas {
    pub workbook: String,
    pub sheet: String,
    pub address: String,
    pub total_rows: usize,
    pub total_cols: usize,
    pub formulas: Vec<Vec<Value>>,
    pub r1c1: Vec<Vec<Value>>,
}

pub trait ExcelBridge: Send + Sync + 'static {
    fn list_workbooks(&self) -> impl Future<Output = Result<Vec<WorkbookInfo>, BridgeError>> + Send;
    fn read_range(&self, req: ReadRequest) -> impl Future<Output = Result<RawRange, BridgeError>> + Send;
    fn get_selection(
        &self,
        req: SelectionRequest,
    ) -> impl Future<Output = Result<RawRange, BridgeError>> + Send;
    fn write_range(&self, req: WriteRequest) -> impl Future<Output = Result<WriteResult, BridgeError>> + Send;
    fn call<A: Serialize + Sync, T: DeserializeOwned>(
        &self,
        script: Script,
        args: &A,
    ) -> impl Future<Output = Result<T, BridgeError>> + Send;
}
