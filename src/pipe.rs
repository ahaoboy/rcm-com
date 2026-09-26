//! Unified named-pipe transport.
//!
//! Every interaction between the `rcm` CLI and the shell extension shares a
//! **single duplex pipe** (`\\.\pipe\rcm_com`), with the DLL as the server and
//! the CLI as the client:
//!
//! ```text
//!   rcm start  ─┐
//!   rcm enable ├─ client ──►  \\.\pipe\rcm_com  ──► DLL (server)
//!   rcm query  ─┘                 (duplex)
//! ```
//!
//! The DLL is the server because it owns all the shared state (menu blocking
//! and log level). Hosting the pipe there lets *any* CLI process connect
//! independently — `rcm enable` no longer needs `rcm start` to be running.
//!
//! Messages are newline-delimited JSON. [`Request`] travels client → server and
//! [`Response`] server → client, so one reader/writer pair covers control
//! commands, queries, log-level changes, and the event stream. JSON escapes
//! control characters, so newline framing is unambiguous for path data.
//!
//! ## Never block the Explorer UI thread
//!
//! [`broadcast_event`] is called from `IContextMenu::QueryContextMenu` on an
//! Explorer UI thread. It only does non-blocking `try_send`s into bounded
//! per-subscriber queues; the actual pipe writes happen on the server runtime.

use std::collections::VecDeque;
use std::ffi::c_void;
use std::os::windows::io::AsRawHandle;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use std::time::{Duration, Instant};

use serde::{Deserialize, Serialize};
use tokio::io::{
    AsyncBufReadExt, AsyncWrite, AsyncWriteExt, BufReader, ReadHalf, WriteHalf,
};
use tokio::net::windows::named_pipe::{
    ClientOptions, NamedPipeClient, NamedPipeServer, PipeMode, ServerOptions,
};

use windows::Win32::Foundation::HANDLE;
use windows::Win32::Storage::FileSystem::FlushFileBuffers;

use crate::consts::PIPE_NAME;
use crate::error::{RcmError, Result};
use crate::logging::LogLevel;
use crate::types::ContextMenuInfo;

// =============================================================================
// Wire protocol
// =============================================================================

/// A request sent from a CLI client to the DLL server.
#[derive(Debug, Serialize, Deserialize)]
#[serde(tag = "cmd", rename_all = "snake_case")]
pub(crate) enum Request {
    /// Block the native context menu.
    Enable,
    /// Stop blocking the native context menu.
    Disable,
    /// Ask whether menu blocking is currently enabled.
    Query,
    /// Stream context-menu events until the connection ends.
    ///
    /// The optional path is the client's own executable; the server records it
    /// as the program currently using the pipe.
    Subscribe {
        path: Option<String>,
        #[serde(default)]
        options: SubscribeOptions,
    },
    /// Change the log level of the running DLL.
    SetLog { level: LogLevel },
    /// Ask the running DLL for its current log level.
    GetLog,
    /// Record the absolute path of the program using the pipe.
    SetClient { path: String },
    /// Ask which program is recorded as using the pipe.
    GetClient,
    /// Set whether Shift+right-click shows the native menu.
    SetShiftBypass { enabled: bool },
    /// Ask whether Shift+right-click shows the native menu.
    GetShiftBypass,
}

/// Optional initialisation parameters a subscriber can pass with
/// [`Request::Subscribe`].
///
/// Every field is optional: an omitted field leaves the corresponding setting
/// unchanged, so a client only sends what it cares about.
#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub(crate) struct SubscribeOptions {
    /// Show the native menu on Shift+right-click. Applied to the DLL for this
    /// session only (not persisted) — the last subscriber to pass a value wins,
    /// since menu blocking is global.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shift_bypass: Option<bool>,
}

/// A message sent from the DLL server to a CLI client.
#[derive(Debug, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub(crate) enum Response {
    /// The control command was applied.
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
    /// A captured context-menu event, delivered to subscribers.
    Event { event: ContextMenuInfo },
}

// =============================================================================
// Tuning
// =============================================================================

/// Delay between connection attempts.
const CONNECT_RETRY: Duration = Duration::from_millis(100);
/// Delay before recreating the server pipe after an error.
const SERVER_RETRY: Duration = Duration::from_millis(200);
/// Per-subscriber queue depth — events are dropped if a client falls behind.
const EVENT_QUEUE_CAP: usize = 64;
/// Events retained for a subscriber that connects slightly late.
const PENDING_CAP: usize = 8;

// =============================================================================
// Server state
// =============================================================================

struct Subscriber {
    id: u64,
    tx: tokio::sync::mpsc::Sender<Arc<str>>,
}

/// Connected event subscribers, protected by a short-lived lock.
static SUBSCRIBERS: Mutex<Vec<Subscriber>> = Mutex::new(Vec::new());

/// Events captured while nobody was subscribed, replayed to the next
/// subscriber so the right-click that loaded the DLL is not lost.
static PENDING: Mutex<VecDeque<Arc<str>>> = Mutex::new(VecDeque::new());

static NEXT_SUBSCRIBER_ID: AtomicU64 = AtomicU64::new(1);
static SERVER_ACTIVE: AtomicBool = AtomicBool::new(false);
static SERVER_STARTED: OnceLock<()> = OnceLock::new();

/// Absolute path of the program currently registered as using the pipe.
///
/// Set explicitly by `rcm client set`, and automatically by every subscriber
/// (so `rcm start` registers itself).
static CLIENT_PATH: Mutex<Option<String>> = Mutex::new(None);

/// Whether the pipe server thread is currently running.
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
// Server
// =============================================================================

/// Start the pipe server, once per process.
///
/// Safe to call from any normal thread — **never** from `DllMain`, which runs
/// under the loader lock.
pub fn start_server() {
    SERVER_STARTED.get_or_init(|| {
        if let Err(err) = std::thread::Builder::new()
            .name("rcm-pipe-server".into())
            .spawn(run_server)
        {
            log::error!("failed to start the pipe server thread: {err}");
        }
    });
}

fn run_server() {
    let runtime = match tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
    {
        Ok(runtime) => runtime,
        Err(err) => {
            log::error!("failed to build the pipe runtime: {err}");
            return;
        }
    };
    SERVER_ACTIVE.store(true, Ordering::Release);
    runtime.block_on(serve());
    SERVER_ACTIVE.store(false, Ordering::Release);
}

/// Accept connections forever, one task per connection.
async fn serve() {
    let mut first = true;
    loop {
        // Restrict the pipe to the current user and Local System so other local
        // processes cannot read captured paths or change the blocking state.
        let mut security = crate::helpers::PipeSecurity::new();
        let mut options = ServerOptions::new();
        // `first_pipe_instance` prevents name squatting, but must be set only
        // for the very first instance — later ones legitimately coexist.
        options.first_pipe_instance(first);
        // Must stay below 255: that value is reserved for
        // PIPE_UNLIMITED_INSTANCES and `ServerOptions` rejects it.
        options.max_instances(16);
        options.pipe_mode(PipeMode::Byte);

        // Safety: `security` owns a valid SECURITY_ATTRIBUTES (or a null
        // descriptor on fallback) that outlives this call.
        let created = unsafe {
            options.create_with_security_attributes_raw(PIPE_NAME, security.as_ptr())
        };
        let server = match created {
            Ok(server) => server,
            Err(err) => {
                log::warn!("failed to create the pipe: {err}");
                tokio::time::sleep(SERVER_RETRY).await;
                continue;
            }
        };
        first = false;

        if let Err(err) = server.connect().await {
            log::warn!("pipe connect failed: {err}");
            tokio::time::sleep(SERVER_RETRY).await;
            continue;
        }
        tokio::spawn(handle_connection(server));
    }
}

/// Read the first request and dispatch it.
async fn handle_connection(server: NamedPipeServer) {
    // Capture the raw handle before splitting so one-shot replies can be
    // flushed to the client before the pipe closes. It is stored as an integer
    // because `HANDLE` is not `Send` and this handler is spawned on the runtime.
    let handle = server.as_raw_handle() as isize;
    let (read_half, write_half) = tokio::io::split(server);
    let mut reader = BufReader::new(read_half);
    let mut writer = write_half;

    let mut line = String::new();
    match reader.read_line(&mut line).await {
        Ok(0) | Err(_) => return,
        Ok(_) => {}
    }
    let request = match serde_json::from_str::<Request>(line.trim_end()) {
        Ok(request) => request,
        Err(err) => {
            log::warn!("ignored malformed pipe request: {err}");
            return;
        }
    };

    match request {
        Request::Subscribe { path, options } => handle_subscription(reader, writer, path, options).await,
        Request::Enable => {
            crate::control::set_enabled(true);
            reply(&mut writer, &Response::Ok, handle).await;
        }
        Request::Disable => {
            crate::control::set_enabled(false);
            reply(&mut writer, &Response::Ok, handle).await;
        }
        Request::Query => {
            let enabled = crate::control::is_enabled();
            reply(&mut writer, &Response::State { enabled }, handle).await;
        }
        Request::SetLog { level } => {
            crate::logging::apply_level(level);
            reply(&mut writer, &Response::Ok, handle).await;
        }
        Request::GetLog => {
            let level = crate::logging::current_level();
            reply(&mut writer, &Response::LogLevel { level }, handle).await;
        }
        Request::SetClient { path } => {
            set_client_path(path);
            reply(&mut writer, &Response::Ok, handle).await;
        }
        Request::GetClient => {
            let path = client_path();
            reply(&mut writer, &Response::Client { path }, handle).await;
        }
        Request::SetShiftBypass { enabled } => {
            crate::control::apply_shift_bypass(enabled);
            reply(&mut writer, &Response::Ok, handle).await;
        }
        Request::GetShiftBypass => {
            let enabled = crate::control::shift_bypass();
            reply(&mut writer, &Response::ShiftBypass { enabled }, handle).await;
        }
    }
}

/// Write a single response and flush it to the client.
async fn reply(writer: &mut WriteHalf<NamedPipeServer>, response: &Response, handle: isize) {
    if write_message(writer, response).await.is_err() {
        return;
    }
    // Closing the pipe can discard buffered data; FlushFileBuffers blocks until
    // the client has read everything, which is what makes one-shot replies
    // (notably `rcm query`) reliable.
    unsafe {
        let _ = FlushFileBuffers(HANDLE(handle as *mut c_void));
    }
}

/// Forward broadcast events to one subscriber until it disconnects.
async fn handle_subscription(
    mut reader: BufReader<ReadHalf<NamedPipeServer>>,
    mut writer: WriteHalf<NamedPipeServer>,
    client: Option<String>,
    options: SubscribeOptions,
) {
    // A subscriber announces its own executable, which is what
    // `rcm client get` reports as the program using the pipe.
    if let Some(path) = client {
        set_client_path(path);
    }
    // Subscription options are applied to the DLL for this session only.
    // Menu blocking is global, so the last subscriber to pass a value wins.
    if let Some(enabled) = options.shift_bypass {
        crate::control::apply_shift_bypass(enabled);
        log::info!(
            "subscriber set shift+right-click native menu to '{}'",
            if enabled { "on" } else { "off" }
        );
    }

    let id = NEXT_SUBSCRIBER_ID.fetch_add(1, Ordering::Relaxed);
    let (tx, mut rx) = tokio::sync::mpsc::channel::<Arc<str>>(EVENT_QUEUE_CAP);
    SUBSCRIBERS
        .lock()
        .unwrap_or_else(|e| e.into_inner())
        .push(Subscriber { id, tx });

    // Replay anything captured before this subscriber arrived.
    let pending: Vec<Arc<str>> = PENDING
        .lock()
        .unwrap_or_else(|e| e.into_inner())
        .drain(..)
        .collect();
    for line in pending {
        if writer.write_all(line.as_bytes()).await.is_err() {
            unregister(id);
            return;
        }
    }
    let _ = writer.flush().await;

    let mut scratch = String::new();
    loop {
        tokio::select! {
            message = rx.recv() => match message {
                Some(line) => {
                    if writer.write_all(line.as_bytes()).await.is_err() {
                        break;
                    }
                    let _ = writer.flush().await;
                }
                None => break,
            },
            // The subscriber sends nothing after subscribing; a read returning
            // 0 bytes is how we notice it disconnected.
            read = reader.read_line(&mut scratch) => match read {
                Ok(0) | Err(_) => break,
                Ok(_) => scratch.clear(),
            },
        }
    }
    unregister(id);
}

fn unregister(id: u64) {
    SUBSCRIBERS
        .lock()
        .unwrap_or_else(|e| e.into_inner())
        .retain(|subscriber| subscriber.id != id);
}

/// Fan a captured context-menu event out to every subscriber.
///
/// Called from the Explorer UI thread: it never blocks. With no subscribers the
/// event is retained (bounded by [`PENDING_CAP`]) for the next subscriber.
pub(crate) fn broadcast_event(event: ContextMenuInfo) {
    let line: Arc<str> = match serde_json::to_string(&Response::Event { event }) {
        Ok(json) => Arc::from(format!("{json}\n")),
        Err(err) => {
            log::warn!("failed to serialise context-menu event: {err}");
            return;
        }
    };

    let mut subscribers = SUBSCRIBERS.lock().unwrap_or_else(|e| e.into_inner());
    if subscribers.is_empty() {
        drop(subscribers);
        let mut pending = PENDING.lock().unwrap_or_else(|e| e.into_inner());
        if pending.len() >= PENDING_CAP {
            pending.pop_front();
        }
        pending.push_back(line);
        return;
    }
    subscribers.retain(|subscriber| {
        !matches!(
            subscriber.tx.try_send(line.clone()),
            Err(tokio::sync::mpsc::error::TrySendError::Closed(_))
        )
    });
}

// =============================================================================
// Client
// =============================================================================

/// Connect to the shell extension's pipe.
///
/// With `timeout == None` this retries forever, which is what `rcm start` wants
/// while it waits for the shell extension to be loaded.
async fn connect(timeout: Option<Duration>) -> Result<NamedPipeClient> {
    let deadline = timeout.map(|timeout| Instant::now() + timeout);
    let mut announced = false;
    loop {
        match ClientOptions::new().open(PIPE_NAME) {
            Ok(client) => return Ok(client),
            Err(err) => {
                if let Some(deadline) = deadline
                    && Instant::now() >= deadline
                {
                    return Err(RcmError::Environment(format!(
                        "the shell extension is not running (pipe '{PIPE_NAME}'): {err}"
                    )));
                }
                if !announced {
                    announced = true;
                    log::info!("waiting for the shell extension on pipe '{PIPE_NAME}'...");
                }
                tokio::time::sleep(CONNECT_RETRY).await;
            }
        }
    }
}

/// Write a newline-terminated JSON message.
async fn write_message<W>(writer: &mut W, message: &(impl Serialize + ?Sized)) -> Result<()>
where
    W: AsyncWrite + Unpin,
{
    let mut json = serde_json::to_vec(message)?;
    json.push(b'\n');
    writer.write_all(&json).await?;
    writer.flush().await?;
    Ok(())
}

/// Send a request and read the single response.
pub(crate) async fn request(request: &Request, timeout: Duration) -> Result<Response> {
    let mut client = connect(Some(timeout)).await?;
    write_message(&mut client, request).await?;

    let mut reader = BufReader::new(client);
    let mut line = String::new();
    if reader.read_line(&mut line).await? == 0 {
        return Err(RcmError::Environment(
            "the shell extension closed the connection".to_string(),
        ));
    }
    Ok(serde_json::from_str(line.trim_end())?)
}

/// Subscribe to the context-menu event stream, reconnecting if the shell
/// extension (or Explorer) restarts. Runs until the process exits.
///
/// `options` are sent with every (re)subscription; see [`SubscribeOptions`].
pub(crate) async fn subscribe<F>(
    mut on_event: F,
    options: Option<bool>,
) -> Result<()>
where
    F: FnMut(ContextMenuInfo),
{
    loop {
        let mut client = connect(None).await?;
        if write_message(
            &mut client,
            &Request::Subscribe {
                path: std::env::current_exe()
                    .ok()
                    .map(|path| path.to_string_lossy().into_owned()),
                options: SubscribeOptions {
                    shift_bypass: options,
                },
            },
        )
        .await
        .is_err()
        {
            continue;
        }
        log::info!("connected to the shell extension — waiting for context menu events");

        let mut reader = BufReader::new(client);
        let mut line = String::new();
        loop {
            line.clear();
            match reader.read_line(&mut line).await {
                Ok(0) | Err(_) => break,
                Ok(_) => match serde_json::from_str::<Response>(line.trim_end()) {
                    Ok(Response::Event { event }) => on_event(event),
                    Ok(other) => log::debug!("ignored unexpected pipe message: {other:?}"),
                    Err(err) => log::warn!("ignored malformed pipe message: {err}"),
                },
            }
        }
        log::warn!("lost the shell extension connection; reconnecting...");
    }
}
