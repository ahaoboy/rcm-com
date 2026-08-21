//! Menu-blocking toggle — global `AtomicBool` + tokio control pipe.
//!
//! A background thread (spawned lazily on first right-click) listens on
//! `\\.\pipe\rcm_com_control`.  External programs send JSON commands like
//! `{"command":"disable"}` over this pipe to toggle the global flag.
//!
//! The CBT hook and `QueryContextMenu` consult [`is_enabled`] before
//! intercepting the native menu.
//!
//! Only two functions are publicly exposed: [`enable`] and [`disable`].

use serde::{Deserialize, Serialize};
use std::sync::OnceLock;
use std::sync::atomic::{AtomicBool, Ordering};
use std::thread;
use std::time::Duration;

use crate::consts::CONTROL_PIPE_NAME;
use crate::error::Result;

// =============================================================================
// Control commands (extensible via serde tagged enum)
// =============================================================================

/// A command sent over the control pipe.
///
/// Serialised as JSON with a `"command"` tag, e.g. `{"command":"enable"}`.
/// Add new variants here to extend the control protocol.
#[derive(Debug, Serialize, Deserialize)]
#[serde(tag = "command", rename_all = "lowercase")]
enum ControlCommand {
    Enable,
    Disable,
    /// Query the current blocking state — the server replies with `true`/`false`.
    Query,
}

// =============================================================================
// Global state
// =============================================================================

/// `true` = block the native context menu (default).
/// `false` = let the system menu appear normally.
static MENU_BLOCKING_ENABLED: AtomicBool = AtomicBool::new(true);

/// Ensures the control-pipe listener thread is spawned exactly once per process.
static LISTENER_STARTED: OnceLock<()> = OnceLock::new();

// =============================================================================
// DLL-internal check (crate-private)
// =============================================================================

/// Ensure the control-pipe listener thread is running.
///
/// Called from `DllMain(DLL_PROCESS_ATTACH)` so the control pipe exists as
/// soon as the DLL is loaded. Idempotent — the listener is spawned at most
/// once per process.
pub fn start() {
    LISTENER_STARTED.get_or_init(|| {
        thread::spawn(run_control_listener);
    });
}

/// Check whether menu blocking is currently enabled.
///
/// Called from the CBT hook and `QueryContextMenu` on every right-click.
/// The listener is already started by [`start`] during DLL load, so no lazy
/// spawn is needed here.
pub fn is_enabled() -> bool {
    MENU_BLOCKING_ENABLED.load(Ordering::Relaxed)
}

// =============================================================================
// Background pipe-listener thread (tokio)
// =============================================================================

/// Run in a dedicated thread: create a named-pipe server, wait for clients,
/// and update [`MENU_BLOCKING_ENABLED`] according to received commands.
///
/// The pipe is destroyed and recreated between connections because tokio's
/// `NamedPipeServer` does not expose `DisconnectNamedPipe`.  A 500 ms sleep
/// after drop gives the Windows kernel time to release the pipe name before
/// the next `create()` call — the DLL is only loaded once per Explorer
/// process so there is no contention from other instances.
fn run_control_listener() {
    let rt = match tokio::runtime::Builder::new_current_thread()
        .enable_io()
        .build()
    {
        Ok(rt) => rt,
        Err(_) => return,
    };

    rt.block_on(async {
        loop {
            let mut server = match tokio::net::windows::named_pipe::ServerOptions::new()
                .create(CONTROL_PIPE_NAME)
            {
                Ok(s) => s,
                Err(_) => {
                    tokio::time::sleep(Duration::from_millis(200)).await;
                    continue;
                }
            };

            if server.connect().await.is_err() {
                tokio::time::sleep(Duration::from_millis(200)).await;
                continue;
            }

            // Read a single command. A bounded read is used instead of
            // read_to_end so the server does not wait for the client to close
            // its write end — Query clients keep the connection open to read
            // the response back.
            let mut buf = [0u8; 64];
            let n = tokio::io::AsyncReadExt::read(&mut server, &mut buf).await;
            let Ok(n) = n else { continue };
            if n == 0 {
                continue;
            }
            if let Ok(cmd) = serde_json::from_slice::<ControlCommand>(&buf[..n]) {
                match cmd {
                    ControlCommand::Enable => MENU_BLOCKING_ENABLED.store(true, Ordering::Relaxed),
                    ControlCommand::Disable => {
                        MENU_BLOCKING_ENABLED.store(false, Ordering::Relaxed)
                    }
                    ControlCommand::Query => {
                        let state = MENU_BLOCKING_ENABLED.load(Ordering::Relaxed);
                        use tokio::io::AsyncWriteExt;
                        let _ = server
                            .write_all(serde_json::to_vec(&state).unwrap_or_default().as_slice())
                            .await;
                    }
                }
            }
        }
    });
}

// =============================================================================
// Public API — the only two functions exposed to consumers
// =============================================================================

/// Enable context-menu blocking (the default).
///
/// Sends `{"command":"enable"}` over the control named pipe.
/// The DLL must be loaded (right-click once in Explorer) for the pipe to exist.
pub fn enable() -> Result<()> {
    send_control(&ControlCommand::Enable)
}

/// Disable context-menu blocking.
///
/// Sends `{"command":"disable"}` over the control named pipe.
pub fn disable() -> Result<()> {
    send_control(&ControlCommand::Disable)
}

/// Query whether context-menu blocking is currently enabled.
///
/// Sends `{"command":"query"}` over the control named pipe and reads the
/// current state back from the DLL, so callers always get the *real* state
/// (unlike [`is_enabled`], which only reads this process's local copy).
pub fn query() -> Result<bool> {
    let cmd = serde_json::to_vec(&ControlCommand::Query)?;
    let max_attempts = 30;
    for _ in 0..max_attempts {
        match std::fs::OpenOptions::new()
            .read(true)
            .write(true)
            .open(CONTROL_PIPE_NAME)
        {
            Ok(mut pipe) => {
                std::io::Write::write_all(&mut pipe, &cmd)?;
                let mut buf = Vec::new();
                std::io::Read::read_to_end(&mut pipe, &mut buf)?;
                return serde_json::from_slice::<bool>(&buf).map_err(Into::into);
            }
            Err(_) => {
                thread::sleep(Duration::from_millis(100));
            }
        }
    }
    Err(crate::error::RcmError::Environment(format!(
        "Control pipe not available after {max_attempts} attempts — \
         right-click in Explorer first to load the DLL"
    )))
}

// =============================================================================
// Pipe client
// =============================================================================

/// Serialise a [`ControlCommand`] to JSON and send it over the control pipe.
///
/// Retries for up to 3 seconds — the pipe server sleeps 500 ms between
/// recreations, so 30 × 100 ms covers that window.
fn send_control(cmd: &ControlCommand) -> Result<()> {
    let json = serde_json::to_vec(cmd)?;
    let max_attempts = 30;
    for _ in 0..max_attempts {
        match std::fs::OpenOptions::new()
            .write(true)
            .open(CONTROL_PIPE_NAME)
        {
            Ok(mut pipe) => {
                std::io::Write::write_all(&mut pipe, &json)?;
                return Ok(());
            }
            Err(_) => {
                thread::sleep(Duration::from_millis(100));
            }
        }
    }
    Err(crate::error::RcmError::Environment(format!(
        "Control pipe not available after {max_attempts} attempts — \
         right-click in Explorer first to load the DLL"
    )))
}
