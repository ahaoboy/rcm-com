//! Menu-blocking state, the CLI-side control API, and the control-pipe server
//! that runs inside the shell extension.
//!
//! The state is process-global to the extension. Transport is the control pipe
//! (see [`crate::pipe`]); the extension hosts it so one-shot commands work
//! whenever the extension is loaded, without a listener process running.
//!
//! Public API: [`enable`], [`disable`], [`query`], [`is_enabled`], [`start`],
//! [`shift_bypass`], [`set_shift_bypass`], [`get_shift_bypass`],
//! [`get_log_level`], [`set_log_level`], [`try_set_remote_log_level`],
//! [`get_client`], and [`set_client`].

use std::ffi::c_void;
use std::os::windows::io::AsRawHandle;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::Duration;

use tokio::io::{AsyncBufReadExt, BufReader, WriteHalf};
use tokio::net::windows::named_pipe::NamedPipeServer;

use windows::Win32::Foundation::HANDLE;
use windows::Win32::Storage::FileSystem::FlushFileBuffers;

use crate::consts::CONTROL_PIPE_NAME;
use crate::error::{RcmError, Result};
use crate::logging::LogLevel;
use crate::pipe::{self, Request, Response};

/// Timeout for one-shot control commands.
const CLIENT_TIMEOUT: Duration = Duration::from_secs(3);
/// Short timeout for best-effort pushes such as the log level.
const NOTIFY_TIMEOUT: Duration = Duration::from_millis(200);
/// Delay before recreating the control pipe after an error.
const SERVER_RETRY: Duration = Duration::from_millis(200);
/// Delay before the supervisor restarts a stopped control server.
const SERVER_RESTART_DELAY: Duration = Duration::from_secs(1);

// =============================================================================
// State
// =============================================================================

/// `true` = block the native context menu (the default).
static MENU_BLOCKING_ENABLED: AtomicBool = AtomicBool::new(true);

/// `true` = let Shift+right-click show the native menu (the default).
///
/// On Windows 11 the classic menu is reached with Shift+right-click, so the
/// default preserves that escape hatch while a plain right-click stays
/// intercepted. Not persisted: an Explorer restart returns to the default.
static SHIFT_BYPASS: AtomicBool = AtomicBool::new(true);

/// Absolute path of the program registered as using the event pipe.
static CLIENT_PATH: std::sync::Mutex<Option<String>> = std::sync::Mutex::new(None);

/// Whether the control-pipe server thread is running.
static SERVER_ACTIVE: AtomicBool = AtomicBool::new(false);
static SERVER_STARTED: std::sync::OnceLock<()> = std::sync::OnceLock::new();

/// Whether menu blocking is currently enabled.
pub fn is_enabled() -> bool {
    MENU_BLOCKING_ENABLED.load(Ordering::Relaxed)
}

/// Update the blocking state (called by the control server).
pub(crate) fn set_enabled(enabled: bool) {
    MENU_BLOCKING_ENABLED.store(enabled, Ordering::Relaxed);
}

/// Whether Shift+right-click shows the native menu instead of being intercepted.
pub fn shift_bypass() -> bool {
    SHIFT_BYPASS.load(Ordering::Relaxed)
}

/// Update the Shift+right-click policy for this process.
pub(crate) fn apply_shift_bypass(enabled: bool) {
    SHIFT_BYPASS.store(enabled, Ordering::Relaxed);
}

/// Whether the control-pipe server thread is running.
pub(crate) fn server_active() -> bool {
    SERVER_ACTIVE.load(Ordering::Acquire)
}

fn set_client_path(path: String) {
    *CLIENT_PATH.lock().unwrap_or_else(|e| e.into_inner()) = Some(path);
}

fn client_path() -> Option<String> {
    CLIENT_PATH
        .lock()
        .unwrap_or_else(|e| e.into_inner())
        .clone()
}

// =============================================================================
// Server (inside the shell extension)
// =============================================================================

/// Start the control-pipe server, once per process.
///
/// Called from the COM entry points, which run on ordinary threads — **never**
/// from `DllMain`, where spawning threads or touching the registry can deadlock
/// Explorer. Idempotent.
pub fn start() {
    SERVER_STARTED.get_or_init(|| {
        if let Err(err) = std::thread::Builder::new()
            .name("rcm-control-server".into())
            .spawn(supervise)
        {
            log::error!("failed to start the control server thread: {err}");
        }
    });
}

/// Keep the control server alive for as long as the process lives.
///
/// The server owns the pipe name, so a silent exit would make every later
/// command fail until Explorer restarts. Restarting from one place keeps the
/// endpoint available.
fn supervise() {
    loop {
        run_server();
        log::warn!(
            "control server stopped; restarting in {} ms",
            SERVER_RESTART_DELAY.as_millis()
        );
        std::thread::sleep(SERVER_RESTART_DELAY);
    }
}

/// Build a runtime and serve until the server stops for any reason.
fn run_server() {
    let runtime = match tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
    {
        Ok(runtime) => runtime,
        Err(err) => {
            log::error!("failed to build the control runtime: {err}");
            return;
        }
    };
    // RAII so the flag is cleared on every exit path, panics included. A flag
    // stuck at `true` would make `DllCanUnloadNow` refuse forever, pinning the
    // DLL — and its file lock — inside Explorer even with no pipe left.
    let _active = ActiveGuard::enter();
    runtime.block_on(serve());
}

struct ActiveGuard;

impl ActiveGuard {
    fn enter() -> Self {
        SERVER_ACTIVE.store(true, Ordering::Release);
        Self
    }
}

impl Drop for ActiveGuard {
    fn drop(&mut self) {
        SERVER_ACTIVE.store(false, Ordering::Release);
    }
}

/// Accept control connections forever, one task per connection.
async fn serve() {
    // Claim the name before serving, retrying with `FIRST_PIPE_INSTANCE` until
    // this process owns it. Joining another process's instance would split the
    // extension's state — menu blocking, Shift policy, log level are all
    // per-process — so commands would reach only some Explorer processes.
    // Retrying (rather than failing) is right because Explorer processes turn
    // over, and this one must be able to take over when the previous owner exits.
    let mut announced = false;
    let mut pending = loop {
        match pipe::create_server(CONTROL_PIPE_NAME, true) {
            Ok(server) => break server,
            Err(err) => {
                if !announced {
                    announced = true;
                    log::info!("waiting to host the control pipe: {err}");
                }
                tokio::time::sleep(SERVER_RETRY).await;
            }
        }
    };

    loop {
        match pending.connect().await {
            Ok(()) => {
                // Detached: the task owns the instance and ends after one reply.
                tokio::spawn(handle_connection(pending));
            }
            Err(err) => {
                log::warn!("control pipe connect failed: {err}");
                tokio::time::sleep(SERVER_RETRY).await;
            }
        }
        // The name is ours now, so later instances are ordinary ones.
        pending = loop {
            match pipe::create_server(CONTROL_PIPE_NAME, false) {
                Ok(server) => break server,
                Err(err) => {
                    log::warn!("failed to create the control pipe: {err}");
                    tokio::time::sleep(SERVER_RETRY).await;
                }
            }
        };
    }
}

/// Read one request, apply it, and reply.
async fn handle_connection(server: NamedPipeServer) {
    // Grab the raw handle before splitting so the reply can be flushed to the
    // client before the pipe closes. Stored as an integer because `HANDLE` is
    // not `Send` and this runs on the runtime.
    let handle = server.as_raw_handle() as isize;
    let (read_half, write_half) = tokio::io::split(server);
    let mut reader = BufReader::new(read_half);
    let mut writer = write_half;

    let mut line = String::new();
    match reader.read_line(&mut line).await {
        Ok(0) | Err(_) => return,
        Ok(_) => {}
    }
    let response = match serde_json::from_str::<Request>(line.trim_end()) {
        Ok(request) => apply(request),
        Err(err) => {
            log::warn!("ignored malformed control request: {err}");
            return;
        }
    };
    reply(&mut writer, &response, handle).await;
}

/// Apply a request to the extension's state and produce the reply.
fn apply(request: Request) -> Response {
    match request {
        Request::Enable => {
            set_enabled(true);
            Response::Ok
        }
        Request::Disable => {
            set_enabled(false);
            Response::Ok
        }
        Request::Query => Response::State {
            enabled: is_enabled(),
        },
        Request::SetLog { level } => {
            crate::logging::apply_level(level);
            Response::Ok
        }
        Request::GetLog => Response::LogLevel {
            level: crate::logging::current_level(),
        },
        Request::SetClient { path } => {
            set_client_path(path);
            Response::Ok
        }
        Request::GetClient => Response::Client {
            path: client_path(),
        },
        Request::SetShiftBypass { enabled } => {
            apply_shift_bypass(enabled);
            Response::Ok
        }
        Request::GetShiftBypass => Response::ShiftBypass {
            enabled: shift_bypass(),
        },
    }
}

/// Write a reply and flush it to the client.
async fn reply(writer: &mut WriteHalf<NamedPipeServer>, response: &Response, handle: isize) {
    if pipe::write_message(writer, response).await.is_err() {
        return;
    }
    // Closing the pipe discards buffered data; FlushFileBuffers blocks until the
    // client has read it, which is what makes one-shot replies reliable.
    unsafe {
        let _ = FlushFileBuffers(HANDLE(handle as *mut c_void));
    }
}

// =============================================================================
// Client API (used by the CLI and by embedding programs)
// =============================================================================

/// Enable context-menu blocking (the default).
pub fn enable() -> Result<()> {
    expect_ok(Request::Enable)
}

/// Disable context-menu blocking.
pub fn disable() -> Result<()> {
    expect_ok(Request::Disable)
}

/// Query whether context-menu blocking is currently enabled.
///
/// Reads from the extension, so it reports the *real* state rather than this
/// process's own copy (which [`is_enabled`] returns).
pub fn query() -> Result<bool> {
    match pipe::request(&Request::Query, CLIENT_TIMEOUT)? {
        Response::State { enabled } => Ok(enabled),
        other => Err(unexpected(other)),
    }
}

/// Register `path` as the program using the event pipe.
pub fn set_client(path: String) -> Result<()> {
    expect_ok(Request::SetClient { path })
}

/// Query the program registered as using the event pipe.
pub fn get_client() -> Result<Option<String>> {
    match pipe::request(&Request::GetClient, CLIENT_TIMEOUT)? {
        Response::Client { path } => Ok(path),
        other => Err(unexpected(other)),
    }
}

/// Query the log level of the running extension.
pub fn get_log_level() -> Result<LogLevel> {
    match pipe::request(&Request::GetLog, CLIENT_TIMEOUT)? {
        Response::LogLevel { level } => Ok(level),
        other => Err(unexpected(other)),
    }
}

/// Change the log level of the running extension.
pub fn set_log_level(level: LogLevel) -> Result<()> {
    expect_ok(Request::SetLog { level })
}

/// Best-effort log-level push. `false` when the extension is not loaded.
pub fn try_set_remote_log_level(level: LogLevel) -> bool {
    matches!(
        pipe::request(&Request::SetLog { level }, NOTIFY_TIMEOUT),
        Ok(Response::Ok)
    )
}

/// Set the Shift+right-click policy on the running extension.
pub fn set_shift_bypass(enabled: bool) -> Result<()> {
    expect_ok(Request::SetShiftBypass { enabled })
}

/// Query the Shift+right-click policy of the running extension.
pub fn get_shift_bypass() -> Result<bool> {
    match pipe::request(&Request::GetShiftBypass, CLIENT_TIMEOUT)? {
        Response::ShiftBypass { enabled } => Ok(enabled),
        other => Err(unexpected(other)),
    }
}

/// Best-effort Shift+right-click push, used by [`crate::server::listen_with`].
pub(crate) fn try_set_shift_bypass(enabled: bool) -> Result<()> {
    match pipe::request(&Request::SetShiftBypass { enabled }, NOTIFY_TIMEOUT)? {
        Response::Ok => Ok(()),
        Response::Error { message } => Err(RcmError::Environment(message)),
        other => Err(unexpected(other)),
    }
}

fn expect_ok(request: Request) -> Result<()> {
    match pipe::request(&request, CLIENT_TIMEOUT)? {
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
