//! Wire protocol and the **control** channel's client side.
//!
//! Two named pipes carry all traffic, with opposite ownership (see
//! [`crate::consts`]):
//!
//! * **Control** (`\\.\pipe\rcm_com_control`) — the shell-extension DLL hosts
//!   it and the `rcm` CLI connects, so `enable` / `query` / `log` work whenever
//!   the extension is loaded. This module holds the request/response types and
//!   the blocking client; the server lives in [`crate::control`].
//! * **Events** (`\\.\pipe\rcm_com`) — the listener hosts it and the DLL
//!   connects. See [`crate::events`].
//!
//! Messages are newline-delimited JSON. JSON escapes control characters, so
//! newline framing stays unambiguous for path data.

use std::time::{Duration, Instant};

use serde::{Deserialize, Serialize};
use tokio::io::{AsyncWrite, AsyncWriteExt};
use tokio::net::windows::named_pipe::{NamedPipeServer, PipeMode, ServerOptions};

use crate::consts::CONTROL_PIPE_NAME;
use crate::error::{RcmError, Result};
use crate::helpers::PipeSecurity;
use crate::logging::LogLevel;

/// Delay between connection attempts by a client.
pub(crate) const CONNECT_RETRY: Duration = Duration::from_millis(100);

// =============================================================================
// Protocol
// =============================================================================

/// A control request sent from the CLI to the shell extension.
#[derive(Debug, Serialize, Deserialize)]
#[serde(tag = "cmd", rename_all = "snake_case")]
pub(crate) enum Request {
    /// Block the native context menu.
    Enable,
    /// Stop blocking the native context menu.
    Disable,
    /// Ask whether menu blocking is currently enabled.
    Query,
    /// Change the log level of the running DLL.
    SetLog { level: LogLevel },
    /// Ask the running DLL for its current log level.
    GetLog,
    /// Record the absolute path of the program using the event pipe.
    SetClient { path: String },
    /// Ask which program is recorded as using the event pipe.
    GetClient,
    /// Set whether Shift+right-click shows the native menu.
    SetShiftBypass { enabled: bool },
    /// Ask whether Shift+right-click shows the native menu.
    GetShiftBypass,
}

/// A reply from the shell extension. Exactly one is sent per [`Request`].
#[derive(Debug, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub(crate) enum Response {
    /// The command was applied.
    Ok,
    /// Reply to [`Request::Query`].
    State { enabled: bool },
    /// Reply to [`Request::GetLog`].
    LogLevel { level: LogLevel },
    /// Reply to [`Request::GetClient`]; `None` when nothing is registered.
    Client { path: Option<String> },
    /// Reply to [`Request::GetShiftBypass`].
    ShiftBypass { enabled: bool },
    /// The request could not be applied.
    Error { message: String },
}

// =============================================================================
// Framing
// =============================================================================

/// Maximum concurrent instances of one pipe.
///
/// Must stay below 255: that value is reserved for `PIPE_UNLIMITED_INSTANCES`
/// and `ServerOptions` rejects it.
const MAX_INSTANCES: usize = 16;

/// Create one pipe-server instance, restricted to the current user and
/// `Local System`.
///
/// Shared by both channels so the security and instance settings cannot drift
/// apart. Synchronous on purpose: [`crate::helpers::PipeSecurity`] owns raw
/// pointers and is not `Send`, so it must be dropped before the caller awaits.
///
/// `first` guards the name against squatting. Only the very first attempt may
/// set it — retries must not, or a failed attempt could never recover.
pub(crate) fn create_server(name: &str, first: bool) -> std::io::Result<NamedPipeServer> {
    let mut security = PipeSecurity::new();
    let mut options = ServerOptions::new();
    options.first_pipe_instance(first);
    options.max_instances(MAX_INSTANCES);
    options.pipe_mode(PipeMode::Byte);
    // Safety: `security` owns a valid SECURITY_ATTRIBUTES (or a null descriptor
    // on fallback) that outlives this call — the kernel copies the descriptor
    // into the pipe object.
    unsafe { options.create_with_security_attributes_raw(name, security.as_ptr()) }
}

/// Write one newline-terminated JSON message.
pub(crate) async fn write_message<W>(
    writer: &mut W,
    message: &(impl Serialize + ?Sized),
) -> Result<()>
where
    W: AsyncWrite + Unpin,
{
    let mut json = serde_json::to_vec(message)?;
    json.push(b'\n');
    writer.write_all(&json).await?;
    writer.flush().await?;
    Ok(())
}

// =============================================================================
// Control client (blocking)
// =============================================================================

/// Send a control request and read its single reply, blocking the caller.
///
/// Control commands are "send, wait, done", so plain blocking I/O keeps the
/// public API synchronous and usable without an async runtime. Only the event
/// stream is async, because it is long-lived.
///
/// The *connection* is retried until `timeout` (the extension may still be
/// starting); the *read* is not, which is fine for a cooperative server that
/// always answers.
pub(crate) fn request(request: &Request, timeout: Duration) -> Result<Response> {
    let deadline = Instant::now() + timeout;
    let mut pipe = loop {
        match std::fs::OpenOptions::new()
            .read(true)
            .write(true)
            .open(CONTROL_PIPE_NAME)
        {
            Ok(pipe) => break pipe,
            Err(err) => {
                if Instant::now() >= deadline {
                    return Err(RcmError::Environment(format!(
                        "the shell extension is not running (pipe '{CONTROL_PIPE_NAME}'): {err}"
                    )));
                }
                std::thread::sleep(CONNECT_RETRY);
            }
        }
    };

    let mut json = serde_json::to_vec(request)?;
    json.push(b'\n');
    std::io::Write::write_all(&mut pipe, &json)?;

    let mut reader = std::io::BufReader::new(pipe);
    let mut line = String::new();
    if std::io::BufRead::read_line(&mut reader, &mut line)? == 0 {
        return Err(RcmError::Environment(
            "the shell extension closed the connection".to_string(),
        ));
    }
    Ok(serde_json::from_str(line.trim_end())?)
}
