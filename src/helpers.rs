//! DLL utility helpers — module handle, path resolution, rate-limited logging,
//! asynchronous pipe delivery, and named-pipe security descriptors.

use std::ffi::c_void;
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicUsize, Ordering};
use std::sync::Mutex;
use std::time::Duration;

use chrono::Utc;
use windows::Win32::Foundation::{CloseHandle, HANDLE, HLOCAL, HMODULE, LocalFree};
use windows::Win32::Security::Authorization::{
    ConvertSidToStringSidW, ConvertStringSecurityDescriptorToSecurityDescriptorW, SDDL_REVISION_1,
};
use windows::Win32::Security::{
    GetTokenInformation, PSECURITY_DESCRIPTOR, SECURITY_ATTRIBUTES, TOKEN_QUERY, TOKEN_USER,
    TokenUser,
};
use windows::Win32::System::LibraryLoader::{GetModuleFileNameW, GetModuleHandleExW};
use windows::Win32::System::Threading::{GetCurrentProcess, OpenProcessToken};
use windows::core::{BOOL, PCWSTR, PWSTR};

/// Handle of the loaded DLL module. Set during `DllMain(DLL_PROCESS_ATTACH)`.
pub(crate) static DLL_MODULE: AtomicUsize = AtomicUsize::new(0);

/// Global COM object reference count for `DllCanUnloadNow`.
pub(crate) static DLL_REF_COUNT: AtomicU32 = AtomicU32::new(0);

/// True while the background context-sender thread is alive.
static SENDER_ACTIVE: AtomicBool = AtomicBool::new(false);

// =============================================================================
// Module path resolution
// =============================================================================

/// Resolve (and cache) this DLL's module handle.
///
/// Falls back to `GetModuleHandleExW(FROM_ADDRESS)` when `DllMain` has not run
/// yet (defensive; normally `DLL_MODULE` is already populated).
fn current_module() -> Option<HMODULE> {
    let mut raw = DLL_MODULE.load(Ordering::Acquire);
    if raw == 0 {
        unsafe {
            let mut hmodule = HMODULE::default();
            let flags = 0x00000004 | 0x00000002; // FROM_ADDRESS | UNCHANGED_REFCOUNT
            let addr = current_module as *const c_void as *const u16;
            if GetModuleHandleExW(flags, PCWSTR(addr), &mut hmodule).is_ok()
                && !hmodule.is_invalid()
            {
                raw = hmodule.0 as usize;
                DLL_MODULE.store(raw, Ordering::Release);
            }
        }
    }
    if raw == 0 {
        None
    } else {
        Some(HMODULE(raw as *mut c_void))
    }
}

/// Read the file name of `module`, growing the buffer until the result is not
/// truncated. `GetModuleFileNameW` returns the buffer size on truncation, so a
/// fixed 1024-element buffer silently clipped very long paths.
fn module_file_name(module: HMODULE) -> Option<String> {
    let mut cap = 260usize;
    loop {
        let mut buf = vec![0u16; cap];
        let len = unsafe { GetModuleFileNameW(Some(module), &mut buf) } as usize;
        if len == 0 {
            return None;
        }
        // A result smaller than the buffer means it fit; a result equal to the
        // buffer size (minus the would-be terminator) means truncation.
        if len < cap - 1 || cap >= 32_768 {
            return Some(String::from_utf16_lossy(&buf[..len]));
        }
        cap = (cap * 2).min(32_768);
    }
}

/// Return the full path to the DLL file itself.
pub(crate) fn dll_path() -> Option<std::path::PathBuf> {
    let name = module_file_name(current_module()?)?;
    Some(std::path::PathBuf::from(name))
}

/// Return the directory containing the DLL.
pub(crate) fn dll_dir() -> Option<std::path::PathBuf> {
    dll_path()?.parent().map(|p| p.to_path_buf())
}

/// Return a UTC timestamp string for log entries and captured events.
pub(crate) fn timestamp() -> String {
    Utc::now().format("%Y-%m-%d %H:%M:%S UTC").to_string()
}

// =============================================================================
// Asynchronous context-menu delivery
// =============================================================================
//
// Historically `QueryContextMenu` opened the data pipe and wrote to it
// synchronously on the Explorer UI thread. If the listener was busy (the
// server handles one connection at a time) this blocked the UI, and an
// unresponsive listener could stall Explorer entirely. The payload is now
// handed to a dedicated sender thread through a bounded queue, so the COM
// callback never blocks.

/// Bounded queue depth — when full, newest events are dropped.
const SEND_QUEUE_CAP: usize = 32;
/// Idle time after which the sender thread exits (so the DLL can unload).
const SEND_IDLE_MS: u64 = 10_000;

static SENDER: Mutex<Option<std::sync::mpsc::SyncSender<String>>> = Mutex::new(None);

/// Whether the background sender thread is currently alive.
pub(crate) fn sender_active() -> bool {
    SENDER_ACTIVE.load(Ordering::Acquire)
}

/// Queue a context-menu JSON payload for asynchronous delivery.
///
/// Never blocks. If the queue is full or the sender could not be started the
/// event is dropped and a rate-limited log entry is written.
pub(crate) fn send_context(json: String) {
    use std::sync::mpsc::TrySendError;

    let mut guard = SENDER.lock().unwrap_or_else(|e| e.into_inner());
    if guard.is_none() {
        *guard = spawn_sender();
    }
    match guard.as_ref() {
        Some(tx) => match tx.try_send(json) {
            Ok(()) => {}
            Err(TrySendError::Full(_)) => {
                drop(guard);
                log::warn!("context pipe queue full — event dropped");
            }
            Err(TrySendError::Disconnected(_)) => {
                *guard = None;
                drop(guard);
                log::warn!("context sender thread gone — event dropped");
            }
        },
        None => {
            drop(guard);
            log::error!("failed to start context sender thread");
        }
    }
}

/// Spawn the sender thread, returning its queue handle.
fn spawn_sender() -> Option<std::sync::mpsc::SyncSender<String>> {
    use std::sync::mpsc::sync_channel;

    let (tx, rx) = sync_channel::<String>(SEND_QUEUE_CAP);
    SENDER_ACTIVE.store(true, Ordering::Release);
    let spawned = std::thread::Builder::new()
        .name("rcm-context-sender".into())
        .spawn(move || {
            // Idle timeout or channel closed ends the loop, so the DLL can be
            // unloaded again.
            while let Ok(msg) = rx.recv_timeout(Duration::from_millis(SEND_IDLE_MS)) {
                if let Err(err) = deliver(&msg) {
                    log::error!("context pipe delivery failed: {err}");
                }
            }
            // Publish `None` while holding the lock so a concurrent
            // `send_context` sees a consistent state and respawns if needed.
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

/// Blocking delivery to the listener — runs on the sender thread only.
fn deliver(json: &str) -> std::io::Result<()> {
    let mut pipe = std::fs::OpenOptions::new()
        .write(true)
        .open(crate::consts::PIPE_NAME)?;
    std::io::Write::write_all(&mut pipe, json.as_bytes())
}

// =============================================================================
// Named-pipe security
// =============================================================================

/// Owns a security descriptor whose DACL grants full control only to the
/// current user and `Local System`, and exposes it as `SECURITY_ATTRIBUTES`
/// for `CreateNamedPipeW`.
///
/// Without this, the default named-pipe DACL grants read access to `Everyone`,
/// letting any local process read captured paths/window data or toggle menu
/// blocking. If the descriptor cannot be built we fall back to a null
/// descriptor (the OS default), so we never make the pipe *less* restricted.
pub(crate) struct PipeSecurity {
    attrs: SECURITY_ATTRIBUTES,
    descriptor: Option<PSECURITY_DESCRIPTOR>,
}

impl PipeSecurity {
    pub(crate) fn new() -> Self {
        let descriptor = unsafe { current_user_security_descriptor() };
        Self {
            attrs: SECURITY_ATTRIBUTES {
                nLength: std::mem::size_of::<SECURITY_ATTRIBUTES>() as u32,
                lpSecurityDescriptor: descriptor
                    .map(|d| d.0)
                    .unwrap_or(std::ptr::null_mut()),
                bInheritHandle: BOOL(0),
            },
            descriptor,
        }
    }

    /// Raw pointer for `ServerOptions::create_with_security_attributes_raw`.
    ///
    /// Only valid for the duration of the create call; the kernel copies the
    /// descriptor into the pipe object.
    pub(crate) fn as_ptr(&mut self) -> *mut c_void {
        &mut self.attrs as *mut SECURITY_ATTRIBUTES as *mut c_void
    }
}

impl Drop for PipeSecurity {
    fn drop(&mut self) {
        if let Some(descriptor) = self.descriptor {
            unsafe {
                let _ = LocalFree(Some(HLOCAL(descriptor.0)));
            }
        }
    }
}

/// Build `D:P(A;;GA;;;<current user>)(A;;GA;;;SY)` as a security descriptor.
///
/// Returns `None` on any failure so the caller can fall back safely.
unsafe fn current_user_security_descriptor() -> Option<PSECURITY_DESCRIPTOR> {
    unsafe {
        let mut token = HANDLE::default();
        OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &mut token).ok()?;

        // Query required size for TOKEN_USER, then fetch it.
        let mut len = 0u32;
        let _ = GetTokenInformation(token, TokenUser, None, 0, &mut len);
        if len == 0 {
            let _ = CloseHandle(token);
            return None;
        }
        let mut buf = vec![0u8; len as usize];
        let info = GetTokenInformation(
            token,
            TokenUser,
            Some(buf.as_mut_ptr() as *mut c_void),
            len,
            &mut len,
        );
        let _ = CloseHandle(token);
        info.ok()?;

        let user = &*(buf.as_ptr() as *const TOKEN_USER);
        let mut sid_str = PWSTR::null();
        ConvertSidToStringSidW(user.User.Sid, &mut sid_str).ok()?;
        let sid = sid_str.to_string().ok()?;
        let _ = LocalFree(Some(HLOCAL(sid_str.0 as *mut c_void)));

        let sddl = format!("D:P(A;;GA;;;{sid})(A;;GA;;;SY)");
        let wide: Vec<u16> = sddl.encode_utf16().chain(std::iter::once(0)).collect();
        let mut descriptor = PSECURITY_DESCRIPTOR(std::ptr::null_mut());
        ConvertStringSecurityDescriptorToSecurityDescriptorW(
            PCWSTR(wide.as_ptr()),
            SDDL_REVISION_1,
            &mut descriptor,
            None,
        )
        .ok()?;
        Some(descriptor)
    }
}
