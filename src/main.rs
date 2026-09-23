mod audit;
mod bridge;
mod cells;
mod formats;
mod rows;
mod server;
mod xlsx;

use std::path::PathBuf;
use std::time::Duration;

use rmcp::ServiceExt;
use tracing_subscriber::EnvFilter;

use crate::bridge::jxa::JxaBridge;
use crate::server::GridskiServer;

/// Long enough for big reads, short enough to report a stuck modal dialog promptly.
const EXCEL_TIMEOUT: Duration = Duration::from_secs(30);

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    // stdout carries the MCP protocol, so logs must go to stderr.
    tracing_subscriber::fmt()
        .with_env_filter(EnvFilter::try_from_env("GRIDSKI_LOG").unwrap_or_else(|_| EnvFilter::new("info")))
        .with_writer(std::io::stderr)
        .with_ansi(false)
        .init();

    let root = workspace_root()?;
    tracing::info!(root = %root.display(), "gridski starting on stdio");
    let service = GridskiServer::new(JxaBridge::new(EXCEL_TIMEOUT), root)
        .serve(rmcp::transport::stdio())
        .await?;
    service.waiting().await?;
    Ok(())
}

/// The one folder Excel is granted access to: `GRIDSKI_ROOT`, or `~/gridski` by default.
/// Created if missing so the first save can prompt for it.
fn workspace_root() -> anyhow::Result<PathBuf> {
    let home = std::env::var_os("HOME").map(PathBuf::from).ok_or_else(|| anyhow::anyhow!("HOME is not set"))?;
    let root = match std::env::var("GRIDSKI_ROOT") {
        Ok(r) if r == "~" => home,
        Ok(r) => match r.strip_prefix("~/") {
            Some(rest) => home.join(rest),
            None => PathBuf::from(r),
        },
        Err(_) => home.join("gridski"),
    };
    anyhow::ensure!(root.is_absolute(), "GRIDSKI_ROOT must be an absolute path, got {}", root.display());
    std::fs::create_dir_all(&root)?;
    Ok(root.canonicalize()?)
}
