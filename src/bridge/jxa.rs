//! [`ExcelBridge`] implemented by piping JavaScript for Automation into `osascript`.

use std::process::Stdio;
use std::time::Duration;

use serde::Serialize;
use serde::de::DeserializeOwned;
use tokio::io::AsyncWriteExt;
use tokio::process::Command;
use tokio::sync::Mutex;

use super::{
    BridgeError, ExcelBridge, RawRange, ReadRequest, Script, SelectionRequest, WorkbookInfo, WriteRequest,
    WriteResult,
};

const PRELUDE: &str = include_str!("scripts/prelude.js");
const LIST_WORKBOOKS: &str = include_str!("scripts/list_workbooks.js");
const READ_RANGE: &str = include_str!("scripts/read_range.js");
const GET_SELECTION: &str = include_str!("scripts/get_selection.js");
const WRITE_RANGE: &str = include_str!("scripts/write_range.js");

impl Script {
    fn source(self) -> &'static str {
        match self {
            Script::OpenWorkbook => include_str!("scripts/open_workbook.js"),
            Script::SaveWorkbook => include_str!("scripts/save_workbook.js"),
            Script::CloseWorkbook => include_str!("scripts/close_workbook.js"),
            Script::ManageSheet => include_str!("scripts/manage_sheet.js"),
            Script::ListNames => include_str!("scripts/list_names.js"),
            Script::DefineName => include_str!("scripts/define_name.js"),
            Script::FormatSource => include_str!("scripts/format_source.js"),
            Script::ReadFormatsLive => include_str!("scripts/read_formats.js"),
            Script::FormatRange => include_str!("scripts/format_range.js"),
            Script::AuditFormulas => include_str!("scripts/audit_formulas.js"),
            Script::FillRight => include_str!("scripts/fill_right.js"),
            Script::BuildRows => include_str!("scripts/build_rows.js"),
        }
    }
}

// Apple Event error numbers we translate into friendlier errors.
const ERR_NOT_FOUND: i64 = -1728;
const ERR_APP_NOT_RUNNING: i64 = -600;
const ERR_TIMEOUT: i64 = -1712;
const ERR_NOT_AUTHORIZED: i64 = -1743;

pub struct JxaBridge {
    timeout: Duration,
    /// Excel handles one script at a time; serializing also keeps writes from interleaving.
    lock: Mutex<()>,
}

impl JxaBridge {
    pub fn new(timeout: Duration) -> Self {
        Self { timeout, lock: Mutex::new(()) }
    }

    async fn run<A: Serialize, T: DeserializeOwned>(&self, body: &str, args: &A) -> Result<T, BridgeError> {
        // Arguments are inlined as a JSON literal rather than passed on argv, so large
        // writes aren't limited by ARG_MAX and nothing needs shell escaping.
        let args = serde_json::to_string(args).expect("bridge args serialize");
        let source = format!("const ARGS = {args};\n{PRELUDE}\n{body}");

        let _guard = self.lock.lock().await;
        let mut child = Command::new("/usr/bin/osascript")
            .args(["-l", "JavaScript", "-"])
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .kill_on_drop(true)
            .spawn()?;
        let mut stdin = child.stdin.take().expect("stdin is piped");
        stdin.write_all(source.as_bytes()).await?;
        drop(stdin);

        let output = tokio::time::timeout(self.timeout, child.wait_with_output())
            .await
            .map_err(|_| BridgeError::Busy)??;

        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr);
            return Err(classify(stderr_error_number(&stderr), None, stderr.trim().to_string()));
        }
        parse_envelope(&String::from_utf8_lossy(&output.stdout))
    }
}

impl ExcelBridge for JxaBridge {
    async fn list_workbooks(&self) -> Result<Vec<WorkbookInfo>, BridgeError> {
        self.run(LIST_WORKBOOKS, &serde_json::json!({})).await
    }

    async fn read_range(&self, req: ReadRequest) -> Result<RawRange, BridgeError> {
        self.run(READ_RANGE, &req).await
    }

    async fn get_selection(&self, req: SelectionRequest) -> Result<RawRange, BridgeError> {
        self.run(GET_SELECTION, &req).await
    }

    async fn write_range(&self, req: WriteRequest) -> Result<WriteResult, BridgeError> {
        self.run(WRITE_RANGE, &req).await
    }

    async fn call<A: Serialize + Sync, T: DeserializeOwned>(&self, script: Script, args: &A) -> Result<T, BridgeError> {
        self.run(script.source(), args).await
    }
}

#[derive(serde::Deserialize)]
#[serde(untagged)]
enum Envelope<T> {
    Ok { ok: T },
    Err { error: ScriptError },
}

#[derive(serde::Deserialize)]
struct ScriptError {
    code: Option<String>,
    number: Option<i64>,
    message: String,
}

fn parse_envelope<T: DeserializeOwned>(stdout: &str) -> Result<T, BridgeError> {
    let envelope: Envelope<T> = serde_json::from_str(stdout.trim())
        .map_err(|e| BridgeError::Script(format!("unexpected output from osascript ({e}): {}", stdout.trim())))?;
    match envelope {
        Envelope::Ok { ok } => Ok(ok),
        Envelope::Err { error } => Err(classify(error.number, error.code.as_deref(), error.message)),
    }
}

fn classify(number: Option<i64>, code: Option<&str>, message: String) -> BridgeError {
    match (code, number) {
        (Some("NOT_RUNNING"), _) | (_, Some(ERR_APP_NOT_RUNNING)) => BridgeError::NotRunning,
        (Some("NOT_FOUND"), _) => BridgeError::NotFound(message),
        (Some("UNSUPPORTED"), _) => BridgeError::Unsupported(message),
        (Some("INVALID"), _) => BridgeError::InvalidInput(message),
        (_, Some(ERR_NOT_AUTHORIZED)) => BridgeError::PermissionDenied,
        (_, Some(ERR_TIMEOUT)) => BridgeError::Busy,
        (_, Some(ERR_NOT_FOUND)) => BridgeError::NotFound(clean_message(&message)),
        _ => BridgeError::Script(clean_message(&message)),
    }
}

/// JXA doubles up prefixes ("Error: Error: ..."); strip them for readability.
fn clean_message(message: &str) -> String {
    let mut m = message.trim();
    while let Some(rest) = m.strip_prefix("Error:") {
        m = rest.trim_start();
    }
    m.to_string()
}

/// osascript failures end with the error number in parentheses: "... (-1743)".
fn stderr_error_number(stderr: &str) -> Option<i64> {
    let trimmed = stderr.trim_end();
    let inner = trimmed.strip_suffix(')')?;
    let start = inner.rfind('(')?;
    inner[start + 1..].parse().ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_ok_envelope() {
        let v: Vec<u32> = parse_envelope(r#"{"ok":[1,2]}"#).unwrap();
        assert_eq!(v, vec![1, 2]);
    }

    #[test]
    fn maps_script_error_codes() {
        let err = parse_envelope::<()>(r#"{"error":{"code":"NOT_RUNNING","number":null,"message":"x"}}"#);
        assert!(matches!(err, Err(BridgeError::NotRunning)));
        let err = parse_envelope::<()>(
            r#"{"error":{"code":null,"number":-1728,"message":"Error: Error: The object does not exist"}}"#,
        );
        assert!(matches!(err, Err(BridgeError::NotFound(m)) if m == "The object does not exist"));
    }

    #[test]
    fn reads_error_number_from_stderr() {
        let stderr = "execution error: Not authorized to send Apple events to Microsoft Excel. (-1743)\n";
        assert_eq!(stderr_error_number(stderr), Some(-1743));
        assert!(matches!(classify(stderr_error_number(stderr), None, stderr.into()), BridgeError::PermissionDenied));
        assert_eq!(stderr_error_number("no number here"), None);
    }
}
