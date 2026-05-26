//! Windows 11 context menu style switching and Explorer restart utilities.
//!
//! Provides functions to toggle between Windows 11 and Windows 10 (classic)
//! right-click context menu styles, and to restart Windows Explorer.

use crate::error::{RcmError, Result};
use std::process::Command;
use std::thread;
use std::time::Duration;

// ── Registry paths ────────────────────────────────────────────────────────

/// The CLSID key that controls whether Windows 11 uses its new compact
/// context menu or falls back to the classic Windows 10 style.
const WIN11_MENU_CLSID: &str =
    r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}";

const INPROC_SERVER32: &str = "InprocServer32";

// ── Public API ────────────────────────────────────────────────────────────

/// Switch between Windows 11 and Windows 10 right-click context menu styles.
///
/// When `use_win10` is `true`, the registry is modified to use the classic
/// Windows 10 expanded context menu.  When `false`, the default Windows 11
/// compact menu is restored.
///
/// **Requires writing to `HKEY_CURRENT_USER`.**  A restart of Explorer is
/// recommended afterwards — call [`restart_explorer`] to do so automatically.
pub fn set_win11_menu_style(use_win10: bool) -> Result<()> {
    use windows::Win32::System::Registry::*;
    use windows::core::PCWSTR;

    let clsid_wide: Vec<u16> = WIN11_MENU_CLSID
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();

    if use_win10 {
        // Create the key to enable classic Win10 menu:
        // HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32
        // with empty default value.
        let inproc_path = format!("{WIN11_MENU_CLSID}\\{INPROC_SERVER32}");
        let inproc_wide: Vec<u16> = inproc_path
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();

        unsafe {
            let mut key = HKEY::default();
            RegCreateKeyW(
                HKEY_CURRENT_USER,
                PCWSTR(inproc_wide.as_ptr()),
                &mut key,
            )
            .ok()
            .map_err(|e| {
                RcmError::Registry(format!("Failed to create Win10 menu key: {e}"))
            })?;

            // Set empty default value
            let empty: [u8; 2] = [0, 0]; // null-terminated empty wide string
            RegSetValueExW(
                key,
                PCWSTR::null(),
                None,
                REG_SZ,
                Some(&empty),
            )
            .ok()
            .map_err(|e| {
                RcmError::Registry(format!("Failed to set Win10 menu value: {e}"))
            })?;

            let _ = RegCloseKey(key);
        }

        println!("✅ Switched to Windows 10 classic context menu style.");
    } else {
        // Delete the key to restore default Win11 menu:
        // HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}
        unsafe {
            // First try to delete the InprocServer32 subkey specifically
            let inproc_path = format!("{WIN11_MENU_CLSID}\\{INPROC_SERVER32}");
            let inproc_wide: Vec<u16> = inproc_path
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect();
            let _ = RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(inproc_wide.as_ptr()));

            // Also delete the parent CLSID key if it's empty
            let _ = RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(clsid_wide.as_ptr()));
        }

        println!("✅ Switched to Windows 11 default context menu style.");
    }

    Ok(())
}

/// Set the default context menu to use the classic (Windows 10) style.
///
/// When `use_classic` is `true`, the registry is configured so that the
/// classic expanded context menu is the default.  When `false`, the
/// Windows 11 compact menu becomes the default.
///
/// This is a semantic alias for [`set_win11_menu_style`] — the underlying
/// mechanism is the same.
pub fn set_default_classic_menu(use_classic: bool) -> Result<()> {
    set_win11_menu_style(use_classic)
}

/// Restart Windows Explorer.
///
/// Kills all `explorer.exe` processes, waits 5 seconds for them to fully
/// terminate, then launches a new Explorer instance.  This is useful after
/// making registry changes that affect the shell.
///
/// # Platform
/// Windows only.
pub fn restart_explorer() -> Result<()> {
    println!("🔄 Stopping Explorer...");

    // Kill all explorer.exe processes
    let kill_status = Command::new("taskkill")
        .args(["/f", "/im", "explorer.exe"])
        .status();

    match kill_status {
        Ok(status) if status.success() => {
            println!("   Explorer stopped.");
        }
        Ok(status) => {
            // Exit code 128 means "no such process" — that's fine
            if status.code() != Some(128) {
                eprintln!("   Warning: taskkill exited with code {:?}", status.code());
            }
        }
        Err(e) => {
            return Err(RcmError::Environment(format!(
                "Failed to stop Explorer: {e}"
            )));
        }
    }

    // Wait 5 seconds
    println!("⏳ Waiting 5 seconds...");
    thread::sleep(Duration::from_secs(5));

    // Restart Explorer
    println!("🚀 Starting Explorer...");
    let start_status = Command::new("explorer.exe").spawn();

    match start_status {
        Ok(_) => {
            println!("✅ Explorer restarted successfully.");
        }
        Err(e) => {
            return Err(RcmError::Environment(format!(
                "Failed to start Explorer: {e}"
            )));
        }
    }

    Ok(())
}

// ── Helpers ───────────────────────────────────────────────────────────────

/// Get the current right-click context menu style.
///
/// Returns `"Win10"` if the classic expanded menu is active,
/// or `"Win11"` if the default compact menu is active.
pub fn get_menu_style() -> &'static str {
    if is_default_classic() {
        "Win10"
    } else {
        "Win11"
    }
}

/// Check whether the classic (Windows 10) context menu is set as the default.
///
/// Returns `true` if the classic expanded menu is the current default,
/// `false` if the Windows 11 compact menu is the default.
pub fn is_default_classic() -> bool {
    use windows::Win32::System::Registry::*;
    use windows::core::PCWSTR;

    let inproc_path = format!("{WIN11_MENU_CLSID}\\{INPROC_SERVER32}");
    let inproc_wide: Vec<u16> = inproc_path
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();

    unsafe {
        let mut key = HKEY::default();
        let result = RegOpenKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(inproc_wide.as_ptr()),
            Some(0),
            KEY_READ,
            &mut key,
        );
        if result.is_ok() {
            let _ = RegCloseKey(key);
            true
        } else {
            false
        }
    }
}
