use std::path::PathBuf;
use windows::Win32::System::Registry::*;
use windows::Win32::UI::Shell::*;
use windows::core::PCWSTR;

use crate::error::{RcmError, Result};

use crate::consts::*;

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
    let wide_val: Vec<u16> = value.encode_utf16().chain(std::iter::once(0)).collect();
    let name_wide: Option<Vec<u16>> =
        name.map(|n| n.encode_utf16().chain(std::iter::once(0)).collect());
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
    let wide: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
    let mut key = HKEY::default();
    unsafe {
        RegCreateKeyW(parent, PCWSTR(wide.as_ptr()), &mut key)
            .ok()
            .map_err(|e| RcmError::Registry(format!("RegCreateKeyW({subkey}) failed: {e}")))?;
    }
    Ok(key)
}

fn delete_key(parent: HKEY, subkey: &str) {
    let wide: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        let _ = RegDeleteTreeW(parent, PCWSTR(wide.as_ptr()));
    }
}

fn open_key(parent: HKEY, subkey: &str) -> Result<HKEY> {
    let wide: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
    let mut key = HKEY::default();
    unsafe {
        RegOpenKeyExW(parent, PCWSTR(wide.as_ptr()), Some(0), KEY_READ, &mut key)
            .ok()
            .map_err(|e| RcmError::Registry(format!("RegOpenKeyExW({subkey}) failed: {e}")))?;
    }
    Ok(key)
}

fn get_reg_value(key: HKEY, name: Option<&str>) -> Result<String> {
    let name_wide: Option<Vec<u16>> =
        name.map(|n| n.encode_utf16().chain(std::iter::once(0)).collect());
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
    let key = create_key(HKEY_CLASSES_ROOT, &clsid_path)?;
    set_reg_value(key, None, "RCM Context Menu Handler")?;
    unsafe {
        let _ = RegCloseKey(key);
    }

    // HKCR\CLSID\{GUID}\InProcServer32
    let inproc_path = format!("CLSID\\{CLSID_STR}\\InProcServer32");
    let key = create_key(HKEY_CLASSES_ROOT, &inproc_path)?;
    set_reg_value(key, None, &dll_str)?;
    set_reg_value(key, Some("ThreadingModel"), "Apartment")?;
    unsafe {
        let _ = RegCloseKey(key);
    }

    // Context menu handler registrations
    for path in get_handler_paths() {
        let key = create_key(HKEY_CLASSES_ROOT, &path)?;
        set_reg_value(key, None, CLSID_STR)?;
        unsafe {
            let _ = RegCloseKey(key);
        }
    }

    // Approved shell extensions
    let approved_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    if let Ok(key) = create_key(HKEY_LOCAL_MACHINE, approved_path) {
        let _ = set_reg_value(key, Some(CLSID_STR), "RCM Context Menu Handler");
        unsafe {
            let _ = RegCloseKey(key);
        }
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
        delete_key(HKEY_CLASSES_ROOT, &path);
    }

    // Remove CLSID registration
    delete_key(HKEY_CLASSES_ROOT, &format!("CLSID\\{CLSID_STR}"));

    // Remove from Approved list
    let approved_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    let wide_clsid: Vec<u16> = CLSID_STR.encode_utf16().chain(std::iter::once(0)).collect();
    if let Ok(key) = create_key(HKEY_LOCAL_MACHINE, approved_path) {
        unsafe {
            let _ = RegDeleteValueW(key, PCWSTR(wide_clsid.as_ptr()));
            let _ = RegCloseKey(key);
        }
    }

    // Notify shell of changes
    unsafe {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }

    println!("Unregistration successful. Restart Explorer to apply.");
    Ok(())
}

pub fn status() -> Result<()> {
    println!("RCM Context Menu Status");
    println!("=======================");

    // Basic Info
    println!("Pipe Name:      {}", crate::PIPE_NAME);
    if let Ok(dll) = dll_path() {
        println!("Expected DLL:   {}", dll.display());
    } else {
        println!("Expected DLL:   Not found in current directory");
    }
    println!();

    // Registry: CLSID
    println!("Registry Check:");
    let clsid_path = format!("CLSID\\{CLSID_STR}");
    match open_key(HKEY_CLASSES_ROOT, &clsid_path) {
        Ok(key) => {
            let name = get_reg_value(key, None).unwrap_or_else(|_| "Unknown".to_string());
            println!("  ✅ CLSID: {CLSID_STR} ({name})");
            unsafe { let _ = RegCloseKey(key); }

            // InProcServer32
            let inproc_path = format!("{clsid_path}\\InProcServer32");
            match open_key(HKEY_CLASSES_ROOT, &inproc_path) {
                Ok(key) => {
                    let path = get_reg_value(key, None).unwrap_or_else(|_| "Unknown".to_string());
                    let thread = get_reg_value(key, Some("ThreadingModel")).unwrap_or_else(|_| "Unknown".to_string());
                    println!("    ✅ InProcServer32: {path}");
                    println!("    ✅ ThreadingModel: {thread}");
                    unsafe { let _ = RegCloseKey(key); }
                }
                Err(_) => println!("    ❌ InProcServer32 key missing"),
            }
        }
        Err(_) => println!("  ❌ CLSID key missing"),
    }

    // Registry: ContextMenuHandlers
    println!("  Handlers:");
    for path in get_handler_paths() {
        match open_key(HKEY_CLASSES_ROOT, &path) {
            Ok(key) => {
                let val = get_reg_value(key, None).unwrap_or_else(|_| "Unknown".to_string());
                if val == CLSID_STR {
                    println!("    ✅ {path}");
                } else {
                    println!("    ⚠️ {path} (Wrong CLSID: {val})");
                }
                unsafe { let _ = RegCloseKey(key); }
            }
            Err(_) => println!("    ❌ {path} (Missing)"),
        }
    }

    // Registry: Approved
    let approved_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    match open_key(HKEY_LOCAL_MACHINE, approved_path) {
        Ok(key) => {
            match get_reg_value(key, Some(CLSID_STR)) {
                Ok(val) => println!("  ✅ Approved: {val}"),
                Err(_) => println!("  ❌ Not in Approved list"),
            }
            unsafe { let _ = RegCloseKey(key); }
        }
        Err(_) => println!("  ❌ Approved key list inaccessible"),
    }

    Ok(())
}
