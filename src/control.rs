//! Menu-blocking state and the public control API.
//!
//! The state itself is a process-global atomic; all transport goes through
//! [`crate::pipe`], which hosts the single duplex named pipe shared with the
//! `rcm` CLI — there is no longer a separate control pipe.
//!
//! The CBT hook and `QueryContextMenu` consult [`is_enabled`] before
//! intercepting the native menu.
//!
//! Public API: [`enable`], [`disable`], [`query`], [`start`],
//! [`get_log_level`], [`set_log_level`], [`try_set_remote_log_level`],
//! [`get_client`], and [`set_client`].

use std::sync::atomic::{AtomicBool, Ordering};
use std::time::Duration;

use crate::error::{RcmError, Result};
use crate::logging::LogLevel;
use crate::pipe::{self, Request, Response};

// =============================================================================
// Global state
// =============================================================================

/// `true` = block the native context menu (default).
/// `false` = let the system menu appear normally.
static MENU_BLOCKING_ENABLED: AtomicBool = AtomicBool::new(true);

/// Timeout for one-shot control commands (`enable` / `disable` / `query`).
const CLIENT_TIMEOUT: Duration = Duration::from_secs(3);
/// Short timeout for best-effort notifications such as pushing a log level.
const NOTIFY_TIMEOUT: Duration = Duration::from_millis(200);

// =============================================================================
// DLL-internal state
// =============================================================================

/// Ensure the pipe server is running.
///
/// Called from `cf_create_instance` (a normal COM activation thread), never
/// from `DllMain`, which runs under the loader lock. Idempotent.
pub fn start() {
    pipe::start_server();
}

/// Check whether menu blocking is currently enabled.
///
/// Called from the CBT hook and `QueryContextMenu` on every right-click.
pub fn is_enabled() -> bool {
    MENU_BLOCKING_ENABLED.load(Ordering::Relaxed)
}

/// Update the blocking state (called by the pipe server for `enable`/`disable`).
pub(crate) fn set_enabled(enabled: bool) {
    MENU_BLOCKING_ENABLED.store(enabled, Ordering::Relaxed);
}

// =============================================================================
// Public API
// =============================================================================

/// Enable context-menu blocking (the default).
pub async fn enable() -> Result<()> {
    expect_ok(Request::Enable).await
}

/// Disable context-menu blocking.
pub async fn disable() -> Result<()> {
    expect_ok(Request::Disable).await
}

/// Query whether context-menu blocking is currently enabled.
///
/// Reads the state back from the DLL, so callers always get the *real* state
/// (unlike [`is_enabled`], which only reads this process's local copy).
pub async fn query() -> Result<bool> {
    match pipe::request(&Request::Query, CLIENT_TIMEOUT).await? {
        Response::State { enabled } => Ok(enabled),
        other => Err(unexpected(other)),
    }
}

/// Register `path` as the program currently using the pipe.
pub async fn set_client(path: String) -> Result<()> {
    expect_ok(Request::SetClient { path }).await
}

/// Query the absolute path of the program registered as using the pipe.
///
/// Returns `None` when nothing has been registered yet.
pub async fn get_client() -> Result<Option<String>> {
    match pipe::request(&Request::GetClient, CLIENT_TIMEOUT).await? {
        Response::Client { path } => Ok(path),
        other => Err(unexpected(other)),
    }
}

/// Query the log level of the running DLL.
pub async fn get_log_level() -> Result<LogLevel> {
    match pipe::request(&Request::GetLog, CLIENT_TIMEOUT).await? {
        Response::LogLevel { level } => Ok(level),
        other => Err(unexpected(other)),
    }
}

/// Change the log level of a running DLL.
///
/// Returns an error when the shell extension is not currently loaded; use
/// [`try_set_remote_log_level`] for a best-effort variant.
pub async fn set_log_level(level: LogLevel) -> Result<()> {
    expect_ok(Request::SetLog { level }).await
}

/// Best-effort request to change the log level of a running DLL.
///
/// Returns `false` when the shell extension is not currently loaded, so the
/// caller can report that the new level applies only after it reloads.
pub async fn try_set_remote_log_level(level: LogLevel) -> bool {
    matches!(
        pipe::request(&Request::SetLog { level }, NOTIFY_TIMEOUT).await,
        Ok(Response::Ok)
    )
}

// =============================================================================
// Helpers
// =============================================================================

/// Send a request that answers with a plain acknowledgement.
async fn expect_ok(request: Request) -> Result<()> {
    match pipe::request(&request, CLIENT_TIMEOUT).await? {
        Response::Ok => Ok(()),
        Response::Error { message } => Err(RcmError::Environment(message)),
        other => Err(unexpected(other)),
    }
}

fn unexpected(response: Response) -> RcmError {
    RcmError::Environment(format!(
        "unexpected response from the shell extension: {response:?}"
    ))
}
