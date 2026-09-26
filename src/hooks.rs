//! WH_CBT hook — prevents the default Windows context menu from appearing.
//!
//! A per-thread WH_CBT hook monitors `HCBT_CREATEWND` and blocks creation of
//! the system popup-menu window (class atom `#32768`).
//!
//! ## Scope
//!
//! The hook is installed when a context-menu handler is initialized and only
//! blocks menus for a short window afterwards ([`MENU_BLOCK_TTL_MS`]). Without
//! this bound, every `#32768` window created on the thread — including
//! unrelated shell popups and dropdowns — was suppressed.
//!
//! Within that window a single invocation can opt out via
//! [`allow_native_menu`]: Shift+right-click reports `CMF_EXTENDEDVERBS`, and in
//! that case the user explicitly asked for the extended verb list, so the
//! native menu is shown instead of being intercepted.
//!
//! ## Lifecycle
//!
//! Hooks are installed per Explorer thread (`WH_CBT` is bound to the thread id
//! passed to `SetWindowsHookExW`). A background janitor thread unhooks every
//! remaining hook once the blocking window has elapsed and then exits, so the
//! DLL can become unloadable again. Each new `Initialize` refresh the window
//! and re-installs on the current thread as needed.

use std::cell::Cell;
use std::collections::HashMap;
use std::ffi::c_void;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};
use std::sync::{LazyLock, Mutex};
use std::time::Duration;

use windows::Win32::Foundation::*;
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::*;

use crate::helpers::{self, DLL_MODULE};

/// How long after a handler is initialized new popup menus are blocked.
const MENU_BLOCK_TTL_MS: u64 = 3_000;
/// How often the janitor checks whether the blocking window has elapsed.
const JANITOR_POLL_MS: u64 = 200;

/// Absolute (process-relative) millisecond deadline until which new popup
/// menus are blocked. `0` means "not blocking".
static BLOCK_DEADLINE_MS: AtomicU64 = AtomicU64::new(0);

/// Number of currently installed CBT hooks.
static ACTIVE_CBT_HOOKS: AtomicUsize = AtomicUsize::new(0);

/// Whether the janitor thread is currently running.
static JANITOR_RUNNING: AtomicBool = AtomicBool::new(false);

/// Installed hook handles keyed by the owning thread id.
static HOOKS: LazyLock<Mutex<HashMap<u32, isize>>> = LazyLock::new(|| Mutex::new(HashMap::new()));

thread_local! {
    /// Handle of the active WH_CBT hook for this thread (unused directly; kept
    /// so a thread's hook can be identified) — see [`HOOKS`] for the registry.
    static CBT_HOOK_HANDLE: Cell<isize> = const { Cell::new(0) };

    /// Set for the duration of one shell-extension invocation to let the
    /// native menu through even while blocking is enabled.
    static ALLOW_NATIVE_MENU: Cell<bool> = const { Cell::new(false) };
}

fn blocking_active() -> bool {
    let deadline = BLOCK_DEADLINE_MS.load(Ordering::Acquire);
    deadline != 0 && helpers::monotonic_millis() <= deadline
}

/// Allow the native context menu for the current invocation on this thread.
///
/// Called from `IContextMenu::QueryContextMenu` when Explorer reports
/// `CMF_EXTENDEDVERBS` (Shift+right-click, i.e. the user explicitly asked for
/// the extended verb list). The state is thread-local because Explorer calls
/// `QueryContextMenu` and then `TrackPopupMenu` — which creates the `#32768`
/// window this hook sees — on the same thread.
pub(crate) fn allow_native_menu() {
    ALLOW_NATIVE_MENU.with(|cell| cell.set(true));
}

/// Clear the per-invocation override.
///
/// Called from `IShellExtInit::Initialize` so a Shift from a previous
/// right-click cannot leak into the next, unrelated one.
pub(crate) fn reset_native_menu_override() {
    ALLOW_NATIVE_MENU.with(|cell| cell.set(false));
}

fn native_menu_allowed() -> bool {
    ALLOW_NATIVE_MENU.with(Cell::get)
}

/// CBT hook procedure — called before windows are created on our thread.
///
/// When `code == HCBT_CREATEWND` (3), `lparam` points to a `CBT_CREATEWNDW`
/// whose `lpcs->lpszClass` identifies the window class. A popup menu has class
/// atom 32768 (0x8000). We return 1 to prevent its creation.
unsafe extern "system" fn cbt_hook_proc(code: i32, _wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    // HCBT_CREATEWND = 3
    if code == 3 {
        unsafe {
            let cbt_ptr = lparam.0 as *const CBT_CREATEWNDW;
            if !cbt_ptr.is_null() && !(*cbt_ptr).lpcs.is_null() {
                let cs = &*(*cbt_ptr).lpcs;
                // For system classes, lpszClass is MAKEINTATOM(32768) = 0x8000.
                let is_system_menu = cs.lpszClass.0 as usize == 32768;
                if is_system_menu
                    && crate::control::is_enabled()
                    && blocking_active()
                    && !native_menu_allowed()
                {
                    return LRESULT(1);
                }
            }
        }
    }
    // Not our target — pass to the next hook in the chain.
    unsafe { CallNextHookEx(None, code, _wparam, lparam) }
}

/// Install (or refresh) the WH_CBT hook for the current Explorer thread and
/// extend the blocking window.
pub(crate) fn install_cbt_menu_blocker() {
    BLOCK_DEADLINE_MS.store(
        helpers::monotonic_millis().saturating_add(MENU_BLOCK_TTL_MS),
        Ordering::Release,
    );

    let tid = unsafe { GetCurrentThreadId() };
    let mut hooks = HOOKS.lock().unwrap_or_else(|e| e.into_inner());

    if let std::collections::hash_map::Entry::Vacant(entry) = hooks.entry(tid) {
        unsafe {
            let hinstance = HINSTANCE(DLL_MODULE.load(Ordering::Acquire) as *mut c_void);
            if let Ok(hook) = SetWindowsHookExW(WH_CBT, Some(cbt_hook_proc), Some(hinstance), tid) {
                entry.insert(hook.0 as isize);
                ACTIVE_CBT_HOOKS.fetch_add(1, Ordering::Relaxed);
                CBT_HOOK_HANDLE.with(|cell| cell.set(hook.0 as isize));
            }
        }
    }

    // Spawn the janitor while holding the lock so its shutdown path (which
    // also holds the lock) can never race with a fresh install.
    ensure_janitor_locked();
}

/// Whether any CBT hook is installed or a blocking window is still in effect.
pub(crate) fn has_active_cbt_hooks() -> bool {
    ACTIVE_CBT_HOOKS.load(Ordering::Relaxed) != 0 || JANITOR_RUNNING.load(Ordering::Acquire)
}

/// Must be called while holding the [`HOOKS`] lock.
fn ensure_janitor_locked() {
    if JANITOR_RUNNING.swap(true, Ordering::AcqRel) {
        return;
    }
    let spawned = std::thread::Builder::new()
        .name("rcm-hook-janitor".into())
        .spawn(janitor_loop);
    if spawned.is_err() {
        JANITOR_RUNNING.store(false, Ordering::Release);
    }
}

/// Periodically unhooks everything once the blocking window has elapsed, then
/// exits so the DLL is not kept resident by this thread.
fn janitor_loop() {
    loop {
        std::thread::sleep(Duration::from_millis(JANITOR_POLL_MS));

        let mut hooks = HOOKS.lock().unwrap_or_else(|e| e.into_inner());
        if blocking_active() {
            continue;
        }

        for (_, hook) in hooks.drain() {
            unhook_raw(hook);
            ACTIVE_CBT_HOOKS.fetch_sub(1, Ordering::Relaxed);
        }

        // Re-check under the lock: if a new install extended the window in the
        // meantime, keep running; otherwise shut down.
        if !blocking_active() {
            JANITOR_RUNNING.store(false, Ordering::Release);
            break;
        }
    }
}

fn unhook_raw(hook: isize) {
    unsafe {
        let _ = UnhookWindowsHookEx(HHOOK(hook as *mut c_void));
    }
}
