//! Configuration module — reads settings from a JSON file next to the DLL.
//!
//! The config file has the same stem as the DLL (e.g. `rcm_com.dll` → `rcm_com.json`).
//!
//! ## Fields
//! - `log` (bool, default `false`): when enabled, program logs are written to `{dll_stem}.log`.
//! - `block_win11_menu` (bool, default `false`): when enabled, `QueryContextMenu` returns
//!   `E_FAIL` to prevent the default Windows 11 right-click menu from appearing.

use serde::{Deserialize, Serialize};

/// Holds the user-configurable behaviour knobs loaded from the JSON file.
#[derive(Debug, Clone, Deserialize, Serialize)]
#[derive(Default)]
pub struct RcmConfig {
    /// If true, write diagnostic logs to `{dll_name}.log`.
    #[serde(default)]
    pub log: bool,
    /// If true, block (intercept) the default Windows 11 right-click context menu.
    #[serde(default)]
    pub block_win11_menu: bool,
}


/// Read config from disk on every call (no caching).
/// Returns the on-disk config, or the default if the file can't be read.
pub fn get_config() -> RcmConfig {
    load_config().unwrap_or_default()
}

// ── internal helpers ──────────────────────────────────────────────────────

fn load_config() -> Option<RcmConfig> {
    let path = config_path()?;
    let text = std::fs::read_to_string(&path).ok()?;
    serde_json::from_str::<RcmConfig>(&text).ok()
}

/// Returns the full path to `{dll_stem}.json`, or `None` if the DLL directory
/// cannot be determined.
pub fn config_path() -> Option<std::path::PathBuf> {
    let dll_path = crate::dll_path()?;
    let stem = dll_path.file_stem()?;
    Some(dll_path.parent()?.join(format!("{}.json", stem.to_string_lossy())))
}

/// Returns the full path to `{dll_stem}.log`, or `None` if the DLL directory
/// cannot be determined.
pub fn log_path() -> Option<std::path::PathBuf> {
    let dll_path = crate::dll_path()?;
    let stem = dll_path.file_stem()?;
    Some(dll_path.parent()?.join(format!("{}.log", stem.to_string_lossy())))
}
