use windows::core::PCWSTR;
use windows::Win32::System::Registry::*;
use windows::Win32::UI::Shell::*;

const CLSID_STR: &str = "{B8A0E19C-4C6D-4A82-9F3B-6E8E7D1F2A5C}";
const HANDLER_NAME: &str = "RcmContextMenu";

fn dll_path() -> Result<std::path::PathBuf, String> {
    let exe = std::env::current_exe().map_err(|e| format!("Failed to get exe path: {e}"))?;
    let dir = exe
        .parent()
        .ok_or_else(|| "Failed to get exe directory".to_string())?;
    let dll = dir.join("rcm_com.dll");
    if !dll.exists() {
        return Err(format!("DLL not found at {}", dll.display()));
    }
    Ok(dll)
}

fn set_reg_value(key: HKEY, name: Option<&str>, value: &str) -> Result<(), String> {
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
        .map_err(|e| format!("RegSetValueExW failed: {e}"))
    }
}

fn create_key(parent: HKEY, subkey: &str) -> Result<HKEY, String> {
    let wide: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
    let mut key = HKEY::default();
    unsafe {
        RegCreateKeyW(parent, PCWSTR(wide.as_ptr()), &mut key)
            .ok()
            .map_err(|e| format!("RegCreateKeyW({subkey}) failed: {e}"))?;
    }
    Ok(key)
}

fn delete_key(parent: HKEY, subkey: &str) {
    let wide: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        let _ = RegDeleteTreeW(parent, PCWSTR(wide.as_ptr()));
    }
}

fn register() -> Result<(), String> {
    let dll = dll_path()?;
    let dll_str = dll.to_string_lossy();

    println!("Registering shell extension...");
    println!("  CLSID: {CLSID_STR}");
    println!("  DLL:   {dll_str}");

    // HKCR\CLSID\{GUID}
    let clsid_path = format!("CLSID\\{CLSID_STR}");
    let key = create_key(HKEY_CLASSES_ROOT, &clsid_path)?;
    set_reg_value(key, None, "RCM Context Menu Handler")?;
    unsafe { let _ = RegCloseKey(key); }

    // HKCR\CLSID\{GUID}\InProcServer32
    let inproc_path = format!("CLSID\\{CLSID_STR}\\InProcServer32");
    let key = create_key(HKEY_CLASSES_ROOT, &inproc_path)?;
    set_reg_value(key, None, &dll_str)?;
    set_reg_value(key, Some("ThreadingModel"), "Apartment")?;
    unsafe { let _ = RegCloseKey(key); }

    // Context menu handler registrations
    let handler_paths = [
        format!("*\\shellex\\ContextMenuHandlers\\{HANDLER_NAME}"),
        format!("Directory\\shellex\\ContextMenuHandlers\\{HANDLER_NAME}"),
        format!("Directory\\Background\\shellex\\ContextMenuHandlers\\{HANDLER_NAME}"),
    ];

    for path in &handler_paths {
        let key = create_key(HKEY_CLASSES_ROOT, path)?;
        set_reg_value(key, None, CLSID_STR)?;
        unsafe { let _ = RegCloseKey(key); }
    }

    // Approved shell extensions
    let approved_path =
        "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    if let Ok(key) = create_key(HKEY_LOCAL_MACHINE, approved_path) {
        let _ = set_reg_value(key, Some(CLSID_STR), "RCM Context Menu Handler");
        unsafe { let _ = RegCloseKey(key); }
    }

    // Notify shell of changes
    unsafe {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }

    println!("Registration successful. Restart Explorer to apply.");
    Ok(())
}

fn unregister() -> Result<(), String> {
    println!("Unregistering shell extension...");

    // Remove handler registrations
    let handler_paths = [
        format!("*\\shellex\\ContextMenuHandlers\\{HANDLER_NAME}"),
        format!("Directory\\shellex\\ContextMenuHandlers\\{HANDLER_NAME}"),
        format!("Directory\\Background\\shellex\\ContextMenuHandlers\\{HANDLER_NAME}"),
    ];
    for path in &handler_paths {
        delete_key(HKEY_CLASSES_ROOT, path);
    }

    // Remove CLSID registration
    delete_key(HKEY_CLASSES_ROOT, &format!("CLSID\\{CLSID_STR}"));

    // Remove from Approved list
    let approved_path =
        "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
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

fn main() {
    let args: Vec<String> = std::env::args().collect();

    if args.len() < 2 {
        print_usage();
        return;
    }

    let result = match args[1].as_str() {
        "register" | "reg" | "install" => register(),
        "unregister" | "unreg" | "uninstall" => unregister(),
        _ => {
            print_usage();
            Ok(())
        }
    };

    if let Err(e) = result {
        eprintln!("Error: {e}");
        std::process::exit(1);
    }
}

fn print_usage() {
    println!("RCM Context Menu - Shell Extension Registration Tool");
    println!();
    println!("Usage: rcm <command>");
    println!();
    println!("Commands:");
    println!("  register    Register the shell extension (requires admin)");
    println!("  unregister  Unregister the shell extension (requires admin)");
    println!();
    println!("The DLL (rcm_com.dll) must be in the same directory as this executable.");
}
