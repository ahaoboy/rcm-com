//! WH_CBT hook — prevents the default Windows context menu from appearing.
//!
//! When a popup menu (window class #32768) is about to be created, the
//! `HCBT_CREATEWND` notification fires. Returning non-zero from the CBT
//! hook procedure prevents the window from being created entirely — the
//! menu never appears, no flash, no timing issues.
//!
//! The hook is installed per-right-click in `IShellExtInit::Initialize` and
//! unhooks itself on the first popup-menu creation it catches.

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
unsafe extern "system" fn cbt_hook_proc(code: i32, _wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    // HCBT_CREATEWND = 3
    if code == 3 {
        unsafe {
            // lparam → CBT_CREATEWNDW* → lpcs → CREATESTRUCTW*
            let cbt_ptr = lparam.0 as *const CBT_CREATEWNDW;
            if !cbt_ptr.is_null() {
                let cs = &*(*cbt_ptr).lpcs;
                // For system classes, lpszClass contains MAKEINTATOM(32768) = 0x8000
                if cs.lpszClass.0 as usize == 32768 {
                    // Prevent this popup menu window from being created.
                    // Then unhook — we only need to catch the first one.
                    let hook = CBT_HOOK_HANDLE.swap(0, Ordering::Acquire);
                    if hook != 0 {
                        let _ = UnhookWindowsHookEx(HHOOK(hook as *mut c_void));
                    }
                    return LRESULT(1);
                }
            }
        }
    }
    // Not our target — pass to the next hook in the chain.
    unsafe { CallNextHookEx(None, code, _wparam, lparam) }
}

/// Install a thread-local WH_CBT hook that blocks the next popup menu.
pub(crate) fn install_cbt_menu_blocker() {
    unsafe {
        // Already installed — nothing to do.
        if CBT_HOOK_HANDLE.load(Ordering::Acquire) != 0 {
            return;
        }

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
