//! DLL utility helpers — module handle, logging, and path resolution.

use std::ffi::c_void;
use std::io::Write;
use std::sync::atomic::{AtomicU32, AtomicUsize, Ordering};

use chrono::Utc;
use windows::Win32::Foundation::HMODULE;
use windows::Win32::System::LibraryLoader::{GetModuleFileNameW, GetModuleHandleExW};

/// Handle of the loaded DLL module. Set during `DllMain(DLL_PROCESS_ATTACH)`.
pub(crate) static DLL_MODULE: AtomicUsize = AtomicUsize::new(0);

/// Global COM object reference count for `DllCanUnloadNow`.
pub(crate) static DLL_REF_COUNT: AtomicU32 = AtomicU32::new(0);

/// Return the directory containing the DLL.
pub(crate) fn dll_dir() -> Option<std::path::PathBuf> {
    let mut raw = DLL_MODULE.load(Ordering::Acquire);
    if raw == 0 {
        unsafe {
            let mut hmodule = HMODULE::default();
            let flags = 0x00000004 | 0x00000002; // FROM_ADDRESS | UNCHANGED_REFCOUNT
            let addr = dll_dir as *const c_void as *const u16;
            if GetModuleHandleExW(flags, windows::core::PCWSTR(addr), &mut hmodule).is_ok()
                && !hmodule.is_invalid()
            {
                raw = hmodule.0 as usize;
                DLL_MODULE.store(raw, Ordering::Release);
            }
        }
    }

    if raw == 0 {
        return None;
    }
    let module = HMODULE(raw as *mut c_void);
    let mut buf = vec![0u16; 1024];
    let len = unsafe { GetModuleFileNameW(Some(module), &mut buf) } as usize;
    if len == 0 {
        return None;
    }
    let path = std::path::PathBuf::from(String::from_utf16_lossy(&buf[..len]));
    path.parent().map(|p| p.to_path_buf())
}

/// Return the full path to the DLL file itself.
#[allow(dead_code)]
pub(crate) fn dll_path() -> Option<std::path::PathBuf> {
    let mut raw = DLL_MODULE.load(Ordering::Acquire);
    if raw == 0 {
        unsafe {
            let mut hmodule = HMODULE::default();
            let flags = 0x00000004 | 0x00000002;
            let addr = dll_path as *const c_void as *const u16;
            if GetModuleHandleExW(flags, windows::core::PCWSTR(addr), &mut hmodule).is_ok()
                && !hmodule.is_invalid()
            {
                raw = hmodule.0 as usize;
                DLL_MODULE.store(raw, Ordering::Release);
            }
        }
    }
    if raw == 0 {
        return None;
    }
    let module = HMODULE(raw as *mut c_void);
    let mut buf = vec![0u16; 1024];
    let len = unsafe { GetModuleFileNameW(Some(module), &mut buf) } as usize;
    if len == 0 {
        return None;
    }
    Some(std::path::PathBuf::from(String::from_utf16_lossy(
        &buf[..len],
    )))
}

/// Write a diagnostic log entry.
///
/// Always writes to `err.txt` as a fallback so errors are never silently lost.
pub(crate) fn write_log(err: impl std::fmt::Display) {
    if let Some(path) = dll_dir().map(|d| d.join("err.txt"))
        && let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(path)
    {
        let _ = writeln!(file, "[{}] {}", timestamp(), err);
    }
}

/// Return a UTC timestamp string for log entries.
pub(crate) fn timestamp() -> String {
    Utc::now().format("%Y-%m-%d %H:%M:%S UTC").to_string()
}
