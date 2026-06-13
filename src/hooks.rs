//! WH_CBT hook — prevents the default Windows context menu from appearing.
//!
//! A thread-local WH_CBT hook monitors `HCBT_CREATEWND` and blocks any
//! window of class `#32768` (the system popup-menu class).
//!
//! ## Lifecycle (critical — do not change without understanding)
//!
//! The hook is installed on the first `Initialize` call and remains active
//! **across handler lifecycles**. It is deliberately NOT uninstalled in
//! `handler_release` because `TrackPopupMenu` may be called by Explorer
//! *after* the handler has been released. Instead, each new `Initialize`
//! call atomically swaps out the old hook for a fresh one. The OS
//! automatically cleans up the last hook when the installing thread exits.

use std::ffi::c_void;
use std::sync::atomic::{AtomicIsize, Ordering};

use windows::Win32::Foundation::*;
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::*;

use crate::helpers::DLL_MODULE;

/// Handle of the active WH_CBT hook (0 when not installed).
/// Atomic to prevent data races between the hook callback and the
/// installer running on different COM apartment threads.
static CBT_HOOK_HANDLE: AtomicIsize = AtomicIsize::new(0);

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

/// Install a thread-local WH_CBT hook.  Atomically replaces any previous hook
/// (old one is uninstalled, new one installed) so only one is active at a time.
pub(crate) fn install_cbt_menu_blocker() {
    unsafe {
        // Uninstall any previous hook first (safety: clean stale hooks).
        uninstall_cbt_menu_blocker();

        let hinstance = HINSTANCE(DLL_MODULE.load(Ordering::Acquire) as *mut c_void);
        let hook = SetWindowsHookExW(
            WH_CBT,
            Some(cbt_hook_proc),
            Some(hinstance),
            GetCurrentThreadId(),
        );
        if let Ok(hook) = hook {
            CBT_HOOK_HANDLE.store(hook.0 as isize, Ordering::Release);
        }
    }
}

/// Uninstall the WH_CBT hook.  Called from `install_cbt_menu_blocker`
/// to atomically swap hooks, and from `uninstall_cbt_menu_blocker` publicly.
pub(crate) fn uninstall_cbt_menu_blocker() {
    let hook = CBT_HOOK_HANDLE.swap(0, Ordering::Acquire);
    if hook != 0 {
        unsafe {
            let _ = UnhookWindowsHookEx(HHOOK(hook as *mut c_void));
        }
    }
}
