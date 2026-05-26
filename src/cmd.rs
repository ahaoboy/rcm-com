use std::fmt::Display;
use std::path::PathBuf;
use crate::error::{RcmError, Result};
use crate::consts::*;
use windows::Win32::Foundation::ERROR_FILE_NOT_FOUND;
use windows::Win32::System::Registry::*;
use windows::Win32::UI::Shell::*;
use windows::core::PCWSTR;

// ── helpers ────────────────────────────────────────────────────────────────

/// Convert `&str` to a null-terminated wide string.
fn to_wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

/// RAII wrapper that calls `RegCloseKey` on drop.
struct RegKeyGuard(HKEY);

impl RegKeyGuard {
    fn new(key: HKEY) -> Self {
        Self(key)
    }
}

impl Drop for RegKeyGuard {
    fn drop(&mut self) {
        if !self.0.is_invalid() {
            unsafe { let _ = RegCloseKey(self.0); }
        }
    }
}

fn dll_path() -> Result<PathBuf> {
    let exe = std::env::current_exe().map_err(|e| RcmError::Environment(format!("Failed to get exe path: {e}")))?;
    let dir = exe
        .parent()
        .ok_or_else(|| RcmError::Environment("Failed to get exe directory".to_string()))?;
    let dll = dir.join("rcm_com.dll");
    if !dll.exists() {
        return Err(RcmError::Environment(format!("DLL not found at {}", dll.display())));
    }
    Ok(dll)
}

fn set_reg_value(key: HKEY, name: Option<&str>, value: &str) -> Result<()> {
    let wide_val = to_wide(value);
    let name_wide: Option<Vec<u16>> =
        name.map(to_wide);
    let name_pcwstr = name_wide
        .as_ref()
        .map(|v| PCWSTR(v.as_ptr()))
        .unwrap_or(PCWSTR::null());

    unsafe {
        RegSetValueExW(
            key,
            name_pcwstr,
            None,
            REG_SZ,
            Some(std::slice::from_raw_parts(
                wide_val.as_ptr() as *const u8,
                wide_val.len() * 2,
            )),
        )
        .ok()
        .map_err(|e| RcmError::Registry(format!("RegSetValueExW failed: {e}")))
    }
}

fn create_key(parent: HKEY, subkey: &str) -> Result<HKEY> {
    let wide = to_wide(subkey);
    let mut key = HKEY::default();
    unsafe {
        RegCreateKeyW(parent, PCWSTR(wide.as_ptr()), &mut key)
            .ok()
            .map_err(|e| RcmError::Registry(format!("RegCreateKeyW({subkey}) failed: {e}")))?;
    }
    Ok(key)
}

fn delete_key(parent: HKEY, subkey: &str) -> Result<()> {
    let wide = to_wide(subkey);
    unsafe {
        RegDeleteTreeW(parent, PCWSTR(wide.as_ptr()))
            .ok()
            .map_err(|e| RcmError::Registry(format!("RegDeleteTreeW({subkey}) failed: {e}")))
    }
}

fn open_key(parent: HKEY, subkey: &str) -> Result<HKEY> {
    let wide = to_wide(subkey);
    let mut key = HKEY::default();
    unsafe {
        let res = RegOpenKeyExW(parent, PCWSTR(wide.as_ptr()), Some(0), KEY_READ, &mut key);
        if res.is_err() {
            if res == ERROR_FILE_NOT_FOUND {
                return Err(RcmError::RegistryKeyNotFound(subkey.to_string()));
            }
            return Err(RcmError::Registry(format!("RegOpenKeyExW({subkey}) failed: {res:?}")));
        }
    }
    Ok(key)
}

fn get_reg_value(key: HKEY, name: Option<&str>) -> Result<String> {
    let name_wide: Option<Vec<u16>> =
        name.map(to_wide);
    let name_pcwstr = name_wide
        .as_ref()
        .map(|v| PCWSTR(v.as_ptr()))
        .unwrap_or(PCWSTR::null());

    let mut buf_len = 0u32;
    unsafe {
        RegQueryValueExW(key, name_pcwstr, None, None, None, Some(&mut buf_len)).ok()
            .map_err(|e| RcmError::Registry(format!("RegQueryValueExW length failed: {e}")))?;

        let mut buf = vec![0u8; buf_len as usize];
        RegQueryValueExW(key, name_pcwstr, None, None, Some(buf.as_mut_ptr()), Some(&mut buf_len)).ok()
            .map_err(|e| RcmError::Registry(format!("RegQueryValueExW failed: {e}")))?;

        Ok(String::from_utf16_lossy(std::slice::from_raw_parts(buf.as_ptr() as *const u16, (buf_len / 2) as usize))
            .trim_matches(char::from(0))
            .to_string())
    }
}

fn get_handler_paths() -> Vec<String> {
    vec![
        format!("*\\shellex\\ContextMenuHandlers\\{HANDLER_NAME}"),
        format!("Directory\\shellex\\ContextMenuHandlers\\{HANDLER_NAME}"),
        format!("Directory\\Background\\shellex\\ContextMenuHandlers\\{HANDLER_NAME}"),
    ]
}

pub fn register() -> Result<()> {
    let dll = dll_path()?;
    let dll_str = dll.to_string_lossy();

    println!("Registering shell extension...");
    println!("  CLSID: {CLSID_STR}");
    println!("  DLL:   {dll_str}");

    // HKCR\CLSID\{GUID}
    let clsid_path = format!("CLSID\\{CLSID_STR}");
    {
        let _key = RegKeyGuard::new(create_key(HKEY_CLASSES_ROOT, &clsid_path)?);
        set_reg_value(_key.0, None, "RCM Context Menu Handler")?;
    }

    // HKCR\CLSID\{GUID}\InProcServer32
    let inproc_path = format!("CLSID\\{CLSID_STR}\\InProcServer32");
    {
        let _key = RegKeyGuard::new(create_key(HKEY_CLASSES_ROOT, &inproc_path)?);
        set_reg_value(_key.0, None, &dll_str)?;
        set_reg_value(_key.0, Some("ThreadingModel"), "Apartment")?;
    }

    // Context menu handler registrations
    for path in get_handler_paths() {
        let _key = RegKeyGuard::new(create_key(HKEY_CLASSES_ROOT, &path)?);
        set_reg_value(_key.0, None, CLSID_STR)?;
    }

    // Approved shell extensions
    let approved_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    if let Ok(key) = create_key(HKEY_LOCAL_MACHINE, approved_path) {
        let _key = RegKeyGuard::new(key);
        let _ = set_reg_value(_key.0, Some(CLSID_STR), "RCM Context Menu Handler");
    }

    // Notify shell of changes
    unsafe {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }

    println!("Registration successful. Restart Explorer to apply.");
    Ok(())
}

pub fn unregister() -> Result<()> {
    println!("Unregistering shell extension...");

    // Remove handler registrations
    for path in get_handler_paths() {
        delete_key(HKEY_CLASSES_ROOT, &path)?;
    }

    // Remove CLSID registration
    delete_key(HKEY_CLASSES_ROOT, &format!("CLSID\\{CLSID_STR}"))?;

    // Remove from Approved list
    let approved_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    let wide_clsid = to_wide(CLSID_STR);
    if let Ok(key) = create_key(HKEY_LOCAL_MACHINE, approved_path) {
        let _key = RegKeyGuard::new(key);
        unsafe {
            let _ = RegDeleteValueW(_key.0, PCWSTR(wide_clsid.as_ptr()));
        }
    }

    // Notify shell of changes
    unsafe {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }

    println!("Unregistration successful. Restart Explorer to apply.");
    Ok(())
}

pub struct Status {
    pub pipe_name: String,
    pub dll_path: Option<PathBuf>,
    pub clsid_exists: bool,
    pub clsid_name: Option<String>,
    pub inproc_path: Option<String>,
    pub threading_model: Option<String>,
    pub handler_star_ok: bool,
    pub handler_directory_ok: bool,
    pub handler_background_ok: bool,
    pub is_approved: bool,
    pub config_path: Option<PathBuf>,
    pub config_log_enabled: bool,
    pub config_block_win11_menu: bool,
}

impl Status {
    pub fn is_valid(&self) -> bool {
        self.dll_path.is_some()
            && self.clsid_exists
            && self.inproc_path.is_some()
            && self.handler_star_ok
            && self.handler_directory_ok
            && self.handler_background_ok
            && self.is_approved
    }
}

impl Display for Status {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "RCM Context Menu Status")?;
        writeln!(f, "=======================")?;
        writeln!(f, "Pipe Name:      {}", self.pipe_name)?;
        if let Some(dll) = &self.dll_path {
            writeln!(f, "Expected DLL:   {}", dll.display())?;
        } else {
            writeln!(f, "Expected DLL:   Not found in current directory")?;
        }
        writeln!(f)?;

        writeln!(f, "Registry Check:")?;
        if self.clsid_exists {
            writeln!(
                f,
                "  ✅ CLSID: {CLSID_STR} ({})",
                self.clsid_name.as_deref().unwrap_or("Unknown")
            )?;
            if let Some(path) = &self.inproc_path {
                writeln!(f, "    ✅ InProcServer32: {path}")?;
            } else {
                writeln!(f, "    ❌ InProcServer32 key missing")?;
            }
            if let Some(model) = &self.threading_model {
                writeln!(f, "    ✅ ThreadingModel: {model}")?;
            }
        } else {
            writeln!(f, "  ❌ CLSID key missing")?;
        }

        writeln!(f, "  Handlers:")?;
        let print_handler = |f: &mut std::fmt::Formatter<'_>, ok: bool, path: &str| {
            if ok {
                writeln!(f, "    ✅ {path}")
            } else {
                writeln!(f, "    ❌ {path} (Missing or Mismatch)")
            }
        };

        let paths = get_handler_paths();
        print_handler(f, self.handler_star_ok, &paths[0])?;
        print_handler(f, self.handler_directory_ok, &paths[1])?;
        print_handler(f, self.handler_background_ok, &paths[2])?;

        if self.is_approved {
            writeln!(f, "  ✅ Approved")?;
        } else {
            writeln!(f, "  ❌ Not in Approved list")?;
        }

        writeln!(f)?;
        writeln!(f, "Config:")?;
        if let Some(ref cfg_path) = self.config_path {
            if cfg_path.exists() {
                writeln!(f, "  ✅ Config file: {}", cfg_path.display())?;
                writeln!(f, "     log:              {}", self.config_log_enabled)?;
                writeln!(f, "     block_win11_menu: {}", self.config_block_win11_menu)?;
            } else {
                writeln!(f, "  ⚠️  Config file not found: {}", cfg_path.display())?;
                writeln!(f, "     (defaults: log=false, block_win11_menu=false)")?;
            }
        } else {
            writeln!(f, "  ⚠️  Cannot determine config file path (DLL location unknown)")?;
        }

        writeln!(f)?;
        if self.is_valid() {
            writeln!(f, "Overall Status: ✅ All items are valid.")?;
        } else {
            writeln!(f, "Overall Status: ❌ Some items are missing or invalid.")?;
        }

        Ok(())
    }
}

pub fn status() -> Result<Status> {
    let dll = dll_path().ok();

    let mut status = Status {
        pipe_name: crate::PIPE_NAME.to_string(),
        dll_path: dll,
        clsid_exists: false,
        clsid_name: None,
        inproc_path: None,
        threading_model: None,
        handler_star_ok: false,
        handler_directory_ok: false,
        handler_background_ok: false,
        is_approved: false,
        config_path: None,
        config_log_enabled: false,
        config_block_win11_menu: false,
    };

    // Config file path & values
    #[cfg(feature = "config")]
    {
        status.config_path = crate::config::config_path();
        let cfg = crate::config::get_config();
        status.config_log_enabled = cfg.log;
        status.config_block_win11_menu = cfg.block_win11_menu;
    }

    // Registry: CLSID
    let clsid_path = format!("CLSID\\{CLSID_STR}");
    if let Ok(key) = open_key(HKEY_CLASSES_ROOT, &clsid_path) {
        let _key = RegKeyGuard::new(key);
        status.clsid_exists = true;
        status.clsid_name = get_reg_value(_key.0, None).ok();

        // InProcServer32
        let inproc_path = format!("{clsid_path}\\InProcServer32");
        if let Ok(key) = open_key(HKEY_CLASSES_ROOT, &inproc_path) {
            let _key = RegKeyGuard::new(key);
            status.inproc_path = get_reg_value(_key.0, None).ok();
            status.threading_model = get_reg_value(_key.0, Some("ThreadingModel")).ok();
        }
    }

    // Handlers
    let handler_paths = get_handler_paths();
    let check_handler = |path: &str| -> bool {
        if let Ok(key) = open_key(HKEY_CLASSES_ROOT, path) {
            let _key = RegKeyGuard::new(key);
            let val = get_reg_value(_key.0, None).unwrap_or_default();
            val == CLSID_STR
        } else {
            false
        }
    };

    status.handler_star_ok = check_handler(&handler_paths[0]);
    status.handler_directory_ok = check_handler(&handler_paths[1]);
    status.handler_background_ok = check_handler(&handler_paths[2]);

    // Approved
    let approved_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    if let Ok(key) = open_key(HKEY_LOCAL_MACHINE, approved_path) {
        let _key = RegKeyGuard::new(key);
        if get_reg_value(_key.0, Some(CLSID_STR)).is_ok() {
            status.is_approved = true;
        }
    }

    Ok(status)
}
