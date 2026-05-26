//! Configuration module — reads settings from a JSON file next to the DLL.
//!
//! The config file has the same stem as the DLL (e.g. `rcm_com.dll` → `rcm_com.json`).
//!
//! ## Fields
//! - `log` (bool, default `false`): when enabled, program logs are written to `{dll_stem}.log`.
//! - `block_win11_menu` (bool, default `false`): when enabled, `QueryContextMenu` returns
//!   `E_FAIL` to prevent the default Windows 11 right-click menu from appearing.

use serde::{Deserialize, Serialize};
use std::sync::OnceLock;

/// Holds the user-configurable behaviour knobs loaded from the JSON file.
#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct RcmConfig {
    /// If true, write diagnostic logs to `{dll_name}.log`.
    #[serde(default)]
    pub log: bool,
    /// If true, block (intercept) the default Windows 11 right-click context menu.
    #[serde(default)]
    pub block_win11_menu: bool,
}

impl Default for RcmConfig {
    fn default() -> Self {
        Self {
            log: false,
            block_win11_menu: false,
        }
    }
}

/// Lazily-initialised static config.  Loading happens once on first access.
static CONFIG: OnceLock<RcmConfig> = OnceLock::new();

/// Return a reference to the global config, loading it from disk if necessary.
pub fn get_config() -> &'static RcmConfig {
    CONFIG.get_or_init(|| load_config().unwrap_or_default())
}

/// Force a reload of the config (useful for `status` subcommand).
/// Returns `None` when the DLL directory cannot be determined or the JSON is malformed.
pub fn reload_config() -> Option<RcmConfig> {
    let cfg = load_config()?;
    // OnceLock only supports `set` once, so we can't truly reload.
    // For the purposes of the `status` command we return the fresh value directly;
    // the caller can display it without overwriting the cached one.
    Some(cfg)
}

// ── internal helpers ──────────────────────────────────────────────────────

fn load_config() -> Option<RcmConfig> {
    let path = config_path()?;

    match std::fs::read_to_string(&path) {
        Ok(text) => match serde_json::from_str::<RcmConfig>(&text) {
            Ok(cfg) => return Some(cfg),
            Err(e) => {
                let _ = std::fs::OpenOptions::new()
                    .create(true)
                    .append(true)
                    .open(path.parent()?.join("err.txt"))
                    .and_then(|mut f| {
                        use std::io::Write;
                        writeln!(f, "[config] Failed to parse {}: {e}", path.display())
                    });
                return None;
            }
        },
        Err(_) => {
            // File doesn't exist — create it with default values.
            let cfg = RcmConfig::default();
            if let Ok(json) = serde_json::to_string_pretty(&cfg) {
                let _ = std::fs::write(&path, json);
            }
            return Some(cfg);
        }
    }
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
