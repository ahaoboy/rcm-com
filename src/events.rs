//! The context-menu **event** channel.
//!
//! The listener process hosts the pipe and each loaded shell-extension instance
//! connects as a client. This direction matters:
//!
//! * The pipe *name* belongs to the listener, so restarting Explorer does not
//!   destroy it — the extension simply reconnects when it is loaded again.
//! * Windows runs several `explorer.exe` processes, and each one loads its own
//!   copy of the extension. As clients they are independent connections, so
//!   their events all arrive at the single listener. (With the DLL as server,
//!   only the process that won the pipe name could ever be heard.)
//!
//! Delivery is **live only**. An event raised while no listener is connected is
//! dropped rather than buffered, so a client that reconnects never receives a
//! burst of stale events.
//!
//! ## Never block the Explorer UI thread
//!
//! [`send`] runs on an Explorer UI thread inside `IContextMenu`. It only does a
//! non-blocking `try_send` into a bounded queue; the pipe write, connection, and
//! reconnection happen on a dedicated thread.

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

use tokio::io::{AsyncBufReadExt, BufReader};
use tokio::net::windows::named_pipe::NamedPipeServer;

use crate::consts::EVENT_PIPE_NAME;
use crate::error::Result;
use crate::pipe::CONNECT_RETRY;
use crate::types::ContextMenuInfo;

/// Depth of the listener's hand-off queue. Bounds memory if a callback stalls.
const EVENT_QUEUE_CAP: usize = 64;
/// Depth of the DLL-side send queue. Only needs to absorb a burst of clicks;
/// a backlog is discarded rather than delivered late (see [`EVENT_TTL`]).
const SEND_QUEUE_CAP: usize = 64;
/// How stale an event may be and still be worth delivering.
///
/// An event captured while no listener was running must not be replayed when one
/// starts up: the user has already moved past that right-click. Any event older
/// than this is dropped instead of delivered, which is what keeps a freshly
/// started listener from receiving a burst of history. Normal delivery takes a
/// few milliseconds, so this only ever rejects genuinely stale events.
const EVENT_TTL: Duration = Duration::from_millis(500);
/// How long the sender thread lingers after the queue drains before exiting.
///
/// The thread exists only to keep pipe work off the Explorer UI thread, so it
/// shuts down once there is nothing to send. The linger avoids respawning it
/// for every click in a short burst.
const SENDER_LINGER: Duration = Duration::from_secs(5);
/// Connection attempts before an event is treated as undeliverable.
///
/// Kept small on purpose: retrying only needs to cover the listener's
/// instance-recreate gap, which is sub-millisecond. A long retry would burn
/// time while events pile up behind it.
const CONNECT_ATTEMPTS: u32 = 3;
/// Delay before recreating a listener instance after an error.
const SERVER_RETRY: Duration = Duration::from_millis(200);

// =============================================================================
// Listener side (the `rcm start` process)
// =============================================================================

/// Stream events from every connected shell-extension instance.
///
/// Never returns: if a client disconnects, the listener keeps accepting. The
/// callback runs on this task, so events reach it in arrival order.
pub(crate) async fn serve<F>(mut on_event: F) -> Result<()>
where
    F: FnMut(ContextMenuInfo),
{
    serve_on(EVENT_PIPE_NAME, &mut on_event).await
}

/// [`serve`] against an arbitrary pipe name.
///
/// Exists so the channel can be exercised without touching the real endpoint —
/// a stale instance of `\\.\pipe\rcm_com` (for example one left by an older
/// extension still loaded in Explorer) would otherwise be indistinguishable
/// from a bug.
async fn serve_on<F>(name: &'static str, on_event: &mut F) -> Result<()>
where
    F: FnMut(ContextMenuInfo),
{
    let (tx, mut rx) = tokio::sync::mpsc::channel::<ContextMenuInfo>(EVENT_QUEUE_CAP);
    tokio::spawn(accept_clients(name, tx));
    while let Some(event) = rx.recv().await {
        on_event(event);
    }
    Ok(())
}

/// Accept extension connections forever, one task per connection.
async fn accept_clients(name: &'static str, tx: tokio::sync::mpsc::Sender<ContextMenuInfo>) {
    let mut first = true;
    loop {
        let is_first_instance = first;
        first = false;

        let server = match crate::pipe::create_server(name, is_first_instance) {
            Ok(server) => server,
            Err(err) => {
                log::warn!("failed to create the event pipe: {err}");
                tokio::time::sleep(SERVER_RETRY).await;
                continue;
            }
        };

        if let Err(err) = server.connect().await {
            log::warn!("event pipe connect failed: {err}");
            tokio::time::sleep(SERVER_RETRY).await;
            continue;
        }
        tokio::spawn(read_events(server, tx.clone()));
    }
}

/// Read newline-delimited events from one extension instance.
async fn read_events(server: NamedPipeServer, tx: tokio::sync::mpsc::Sender<ContextMenuInfo>) {
    let mut reader = BufReader::new(server);
    let mut line = String::new();
    loop {
        line.clear();
        match reader.read_line(&mut line).await {
            // 0 bytes means the client closed; an error is equivalent.
            Ok(0) | Err(_) => return,
            Ok(_) => match serde_json::from_str::<ContextMenuInfo>(line.trim_end()) {
                Ok(event) => {
                    if tx.send(event).await.is_err() {
                        return;
                    }
                }
                Err(err) => log::warn!("ignored malformed event: {err}"),
            },
        }
    }
}

// =============================================================================
// Extension side (inside explorer.exe)
// =============================================================================

/// A serialised event plus when it was queued, so delivery can tell how stale it
/// has become while waiting.
struct Queued {
    enqueued: Instant,
    line: Arc<str>,
}

/// Queue an event for delivery to the listener. Never blocks.
///
/// Delivery is live-only: an event is dropped rather than delivered late, either
/// because it aged past [`EVENT_TTL`] or because no listener is connected. That
/// is deliberate — the alternative is stalling Explorer's UI thread or replaying
/// right-clicks the user has already moved past.
pub(crate) fn send(event: ContextMenuInfo) {
    let line: Arc<str> = match serde_json::to_string(&event) {
        Ok(json) => Arc::from(format!("{json}\n")),
        Err(err) => {
            log::warn!("failed to serialise event: {err}");
            return;
        }
    };

    let mut guard = SENDER.lock().unwrap_or_else(|e| e.into_inner());
    if guard.is_none() {
        *guard = spawn_sender();
    }
    if let Some(tx) = guard.as_ref()
        && tx
            .try_send(Queued {
                enqueued: Instant::now(),
                line,
            })
            .is_err()
    {
        // Full or disconnected: drop silently. Logging here would fire on every
        // right-click whenever `rcm start` is not running.
        log::trace!("event dropped (queue full or sender gone)");
    }
}

/// Whether the background sender thread is alive.
pub(crate) fn sender_active() -> bool {
    SENDER_ACTIVE.load(Ordering::Acquire)
}

/// Handle to the sender thread's queue; `None` once the thread has exited.
static SENDER: Mutex<Option<std::sync::mpsc::SyncSender<Queued>>> = Mutex::new(None);
static SENDER_ACTIVE: AtomicBool = AtomicBool::new(false);

/// Start the sender thread, returning its queue handle.
fn spawn_sender() -> Option<std::sync::mpsc::SyncSender<Queued>> {
    use std::sync::mpsc::sync_channel;

    let (tx, rx) = sync_channel::<Queued>(SEND_QUEUE_CAP);
    SENDER_ACTIVE.store(true, Ordering::Release);
    let spawned = std::thread::Builder::new()
        .name("rcm-event-sender".into())
        .spawn(move || {
            deliver_loop(rx);
            // Publish `None` under the lock so a concurrent `send` sees a
            // consistent state and respawns the thread if needed.
            let mut guard = SENDER.lock().unwrap_or_else(|e| e.into_inner());
            *guard = None;
            SENDER_ACTIVE.store(false, Ordering::Release);
        });
    if spawned.is_err() {
        SENDER_ACTIVE.store(false, Ordering::Release);
        return None;
    }
    Some(tx)
}

/// Send one event over its own connection. `false` if no listener accepted it.
///
/// A connection is opened per event and closed as soon as it is written. Right-
/// clicks are infrequent and each connect is local and cheap, so the cost is
/// negligible — while holding a connection open between clicks would mean
/// tracking its health, reconnecting when the listener restarts, and keeping a
/// handle alive inside Explorer for no benefit.
///
/// A fresh connection per event also means a listener that was restarted is
/// picked up automatically: the next click simply connects to the new one.
fn deliver(name: &str, line: &str) -> bool {
    let Some(mut pipe) = connect(name) else {
        log::trace!("no listener running; event dropped");
        return false;
    };
    if let Err(err) = std::io::Write::write_all(&mut pipe, line.as_bytes()) {
        log::debug!("event write failed: {err}");
        return false;
    }
    // Closing a pipe handle can discard data the far end has not read yet, so
    // flush before dropping it. `FlushFileBuffers` blocks until the listener has
    // consumed the line, which is what makes one-connection-per-event lossless.
    // It blocks only this sender thread, never the Explorer UI thread.
    unsafe {
        use std::os::windows::io::AsRawHandle;
        use windows::Win32::Foundation::HANDLE;
        let _ = windows::Win32::Storage::FileSystem::FlushFileBuffers(HANDLE(pipe.as_raw_handle()));
    }
    // `pipe` is dropped here, closing the connection immediately.
    true
}

/// Serve queued events, dropping any that have gone stale.
///
/// `recv_timeout` returns immediately while the queue is non-empty, so iterating
/// one item at a time re-checks staleness for each — no inner drain needed.
fn deliver_loop(rx: std::sync::mpsc::Receiver<Queued>) {
    deliver_loop_on(rx, EVENT_PIPE_NAME)
}

/// [`deliver_loop`] against an arbitrary pipe name, so the staleness rules can be
/// exercised without touching the real endpoint.
fn deliver_loop_on(rx: std::sync::mpsc::Receiver<Queued>, name: &str) {
    loop {
        // Idle for a full linger window: nothing left to send, so stop.
        let Ok(item) = rx.recv_timeout(SENDER_LINGER) else {
            return;
        };

        if item.enqueued.elapsed() > EVENT_TTL {
            // Captured while no listener was running. Dropping it — and the rest
            // of the backlog behind it — is the point: a listener starting later
            // must not replay right-clicks the user has long since moved past.
            discard_backlog(&rx);
            continue;
        }

        if !deliver(name, &item.line) {
            // No listener to receive it, so nothing else queued can be delivered
            // either; clear the backlog instead of letting it age into a burst.
            discard_backlog(&rx);
        }
    }
}

/// Drop everything still queued.
fn discard_backlog(rx: &std::sync::mpsc::Receiver<Queued>) {
    let mut dropped = 0usize;
    while rx.try_recv().is_ok() {
        dropped += 1;
    }
    if dropped > 0 {
        log::trace!("discarded {dropped} queued event(s)");
    }
}

/// Connect to a listener, retrying briefly. `None` when it is not running.
fn connect(name: &str) -> Option<std::fs::File> {
    for _ in 0..CONNECT_ATTEMPTS {
        match std::fs::OpenOptions::new().write(true).open(name) {
            Ok(pipe) => return Some(pipe),
            Err(_) => std::thread::sleep(CONNECT_RETRY),
        }
    }
    None
}
