//! WH_CBT hook — prevents the default Windows context menu from appearing.
//!
//! A thread-local WH_CBT hook monitors `HCBT_CREATEWND` and blocks any
//! window of class `#32768` (the system popup-menu class).
//!
//! ## Lifecycle (critical — do not change without understanding)
//!
//! The hook is thread-local because `WH_CBT` hooks are installed per Explorer
//! thread. It is deliberately NOT uninstalled in `handler_release` because
//! `TrackPopupMenu` may be called by Explorer *after* the handler has been
//! released. Instead, each new `Initialize` call refreshes the hook for the
//! current thread only. The OS automatically cleans up hooks when their
//! installing threads exit.

use std::cell::Cell;
use std::ffi::c_void;
use std::sync::atomic::{AtomicUsize, Ordering};

use windows::Win32::Foundation::*;
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::*;

use crate::helpers::DLL_MODULE;

static ACTIVE_CBT_HOOK_THREADS: AtomicUsize = AtomicUsize::new(0);

thread_local! {
    /// Handle of the active WH_CBT hook for this Explorer thread.
    static CBT_HOOK_HANDLE: Cell<isize> = const { Cell::new(0) };
}

/// CBT hook procedure — called before windows are created on our thread.
///
/// When `code == HCBT_CREATEWND` (3), `lparam` points to a `CBT_CREATEWNDW`
/// whose `lpcs->lpszClass` identifies the window class. A popup menu has
/// class atom 32768 (0x8000). We return 1 to prevent its creation.
///
/// The hook never unhooks itself; it stays alive across handler lifecycles.
/// See module-level docs for the rationale.
unsafe extern "system" fn cbt_hook_proc(code: i32, _wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    // HCBT_CREATEWND = 3
    if code == 3 {
        unsafe {
            let cbt_ptr = lparam.0 as *const CBT_CREATEWNDW;
            if !cbt_ptr.is_null() {
                let cs = &*(*cbt_ptr).lpcs;
                // For system classes, lpszClass is MAKEINTATOM(32768) = 0x8000
                if cs.lpszClass.0 as usize == 32768 {
                    // Block this popup menu window — but keep the hook alive
                    // so it can catch subsequent menu windows from other handlers.
                    return LRESULT(1);
                }
            }
        }
    }
    // Not our target — pass to the next hook in the chain.
    unsafe { CallNextHookEx(None, code, _wparam, lparam) }
}

/// Install a WH_CBT hook for the current Explorer thread.
///
/// A process-global hook handle is incorrect here: `WH_CBT` is bound to the
/// thread id passed to `SetWindowsHookExW`. Explorer can create context menus
/// on different COM apartment threads over time, so each thread needs its own
/// hook handle. Otherwise one thread can accidentally uninstall another
/// thread's hook and native menus start leaking through intermittently.
pub(crate) fn install_cbt_menu_blocker() {
    unsafe {
        let hinstance = HINSTANCE(DLL_MODULE.load(Ordering::Acquire) as *mut c_void);
        let hook = SetWindowsHookExW(
            WH_CBT,
            Some(cbt_hook_proc),
            Some(hinstance),
            GetCurrentThreadId(),
        );
        if let Ok(hook) = hook {
            CBT_HOOK_HANDLE.with(|cell| {
                let old_hook = cell.replace(hook.0 as isize);
                if old_hook == 0 {
                    ACTIVE_CBT_HOOK_THREADS.fetch_add(1, Ordering::Relaxed);
                }
                if old_hook != 0 && old_hook != hook.0 as isize {
                    unhook_raw(old_hook);
                }
            });
        }
    }
}

pub(crate) fn has_active_cbt_hooks() -> bool {
    ACTIVE_CBT_HOOK_THREADS.load(Ordering::Relaxed) != 0
}

fn unhook_raw(hook: isize) {
    unsafe {
        let _ = UnhookWindowsHookEx(HHOOK(hook as *mut c_void));
    }
}
