#![allow(non_snake_case)]

use serde::{Deserialize, Serialize};
use std::ffi::c_void;
use std::io::Write;
use std::sync::atomic::{AtomicU32, AtomicUsize, Ordering};
pub mod cmd;
pub mod consts;
pub mod error;
pub mod server;
use windows::Win32::Foundation::*;
use windows::Win32::System::LibraryLoader::*;
use windows::Win32::System::SystemServices::*;
use windows::Win32::UI::Shell::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::{GUID, HRESULT};

// =============================================================================
// Constants
// =============================================================================

use crate::consts::*;

static DLL_MODULE: AtomicUsize = AtomicUsize::new(0);
static DLL_REF_COUNT: AtomicU32 = AtomicU32::new(0);

// =============================================================================
// Helpers
// =============================================================================

fn dll_dir() -> Option<std::path::PathBuf> {
    let mut raw = DLL_MODULE.load(Ordering::SeqCst);
    if raw == 0 {
        unsafe {
            let mut hmodule = HMODULE::default();
            let flags = 0x00000004 | 0x00000002; // FROM_ADDRESS | UNCHANGED_REFCOUNT
            let addr = dll_dir as *const c_void as *const u16;
            if GetModuleHandleExW(flags, windows::core::PCWSTR(addr), &mut hmodule).is_ok()
                && !hmodule.is_invalid()
            {
                raw = hmodule.0 as usize;
                DLL_MODULE.store(raw, Ordering::SeqCst);
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


fn write_error(err: impl std::fmt::Display) {
    if let Some(path) = dll_dir().map(|d| d.join("err.txt"))
        && let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(path)
    {
        let _ = writeln!(file, "[{}] Error: {}", timestamp(), err);
    }
}

/// Compute a UTC timestamp string without external dependencies.
/// Uses Howard Hinnant's civil_from_days algorithm.
fn timestamp() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();

    let z = (secs / 86400) as i64 + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = (z - era * 146097) as u32;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };

    let tod = secs % 86400;
    let h = tod / 3600;
    let min = (tod % 3600) / 60;
    let s = tod % 60;
    format!("{y:04}-{m:02}-{d:02} {h:02}:{min:02}:{s:02} UTC")
}

// =============================================================================
// ContextMenuInfo - all captured right-click context data
// =============================================================================

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type")]
pub enum Event {
    LeftClickSelect { flags: u32 },
    RightClickMenu { flags: u32 },
    ShiftSelect { flags: u32 },
}

impl Default for Event {
    fn default() -> Self {
        Event::RightClickMenu {
            flags: 0,
        }
    }
}

impl Event {
    pub fn flags(&self) -> u32 {
        match self {
            Event::LeftClickSelect { flags } => *flags,
            Event::RightClickMenu { flags } => *flags,
            Event::ShiftSelect { flags } => *flags,
        }
    }

    pub fn flags_str(&self) -> String {
        let uflags = self.flags();
        let mut flags_str = Vec::new();
        if uflags == 0 { flags_str.push("CMF_NORMAL"); }
        if uflags & 0x00000001 != 0 { flags_str.push("CMF_DEFAULTONLY"); }
        if uflags & 0x00000002 != 0 { flags_str.push("CMF_VERBSONLY"); }
        if uflags & 0x00000004 != 0 { flags_str.push("CMF_EXPLORE"); }
        if uflags & 0x00000008 != 0 { flags_str.push("CMF_NOVERBS"); }
        if uflags & 0x00000010 != 0 { flags_str.push("CMF_CANRENAME"); }
        if uflags & 0x00000020 != 0 { flags_str.push("CMF_NODEFAULT"); }
        if uflags & 0x00000040 != 0 { flags_str.push("CMF_INCLUDESTATIC"); }
        if uflags & 0x00000080 != 0 { flags_str.push("CMF_ITEMMENU"); }
        if uflags & 0x00000100 != 0 { flags_str.push("CMF_EXTENDEDVERBS"); }
        if uflags & 0x00000200 != 0 { flags_str.push("CMF_DISABLEDVERBS"); }
        if uflags & 0x00000400 != 0 { flags_str.push("CMF_ASYNCVERBSTATE"); }
        if uflags & 0x00000800 != 0 { flags_str.push("CMF_OPTIMIZEFORINVOKE"); }
        if uflags & 0x00001000 != 0 { flags_str.push("CMF_SYNCCASCADEMENU"); }
        if uflags & 0x00002000 != 0 { flags_str.push("CMF_DONOTPICKDEFAULT"); }
        if uflags & 0x00010000 != 0 { flags_str.push("CMF_DVFILE"); }
        flags_str.join(" | ")
    }
}

impl std::fmt::Display for Event {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = match self {
            Event::LeftClickSelect { .. } => "LeftClickSelect",
            Event::RightClickMenu { .. } => "RightClickMenu",
            Event::ShiftSelect { .. } => "ShiftSelect",
        };
        write!(f, "{} ({} - {})", name, self.flags(), self.flags_str())
    }
}

#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub struct ContextMenuInfo {
    pub cid: String,
    pub timestamp: String,
    pub cursor_x: i32,
    pub cursor_y: i32,
    pub folder_path: String,
    pub selected_files: Vec<String>,
    pub file_count: u32,
    pub is_background: bool,
    pub window_handle: usize,
    pub window_class: String,
    pub process_id: u32,
    pub event: Event,
}

impl std::fmt::Display for ContextMenuInfo {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "[{}]", self.timestamp)?;
        writeln!(f, "Position: ({}, {})", self.cursor_x, self.cursor_y)?;
        writeln!(f, "Directory: {}", self.folder_path)?;
        writeln!(f, "Background: {}", self.is_background)?;
        writeln!(f, "File Count: {}", self.file_count)?;
        writeln!(f, "Window: 0x{:X}", self.window_handle)?;
        writeln!(f, "Window Class: {}", self.window_class)?;
        writeln!(f, "Process ID: {}", self.process_id)?;
        writeln!(f, "Event: {}", self.event)?;
        if !self.selected_files.is_empty() {
            writeln!(f, "Selected Files:")?;
            for file in &self.selected_files {
                writeln!(f, "  - {file}")?;
            }
        }
        writeln!(f, "---")?;
        Ok(())
    }
}

// =============================================================================
// Raw COM vtable definitions (C ABI compatible)
// =============================================================================

#[repr(C)]
struct IUnknownVtbl {
    QueryInterface:
        unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> HRESULT,
    AddRef: unsafe extern "system" fn(*mut c_void) -> u32,
    Release: unsafe extern "system" fn(*mut c_void) -> u32,
}

#[repr(C)]
struct IClassFactoryVtbl {
    base: IUnknownVtbl,
    CreateInstance: unsafe extern "system" fn(
        *mut c_void,
        *mut c_void,
        *const GUID,
        *mut *mut c_void,
    ) -> HRESULT,
    LockServer: unsafe extern "system" fn(*mut c_void, i32) -> HRESULT, // BOOL = i32
}

#[repr(C)]
struct IShellExtInitVtbl {
    base: IUnknownVtbl,
    Initialize: unsafe extern "system" fn(
        *mut c_void,   // this
        *const c_void, // pidlFolder (PCIDLIST_ABSOLUTE)
        *mut c_void,   // pDataObj (IDataObject*)
        isize,         // hKeyProgID (HKEY)
    ) -> HRESULT,
}

#[repr(C)]
struct IContextMenuVtbl {
    base: IUnknownVtbl,
    QueryContextMenu: unsafe extern "system" fn(*mut c_void, isize, u32, u32, u32, u32) -> HRESULT,
    InvokeCommand: unsafe extern "system" fn(*mut c_void, *const c_void) -> HRESULT,
    GetCommandString:
        unsafe extern "system" fn(*mut c_void, usize, u32, *const u32, *mut u8, u32) -> HRESULT,
}

/// Raw FORMATETC for IDataObject::GetData call
#[repr(C)]
struct RawFormatEtc {
    cf_format: u16,
    ptd: *mut c_void,
    dw_aspect: u32,
    lindex: i32,
    tymed: u32,
}

/// Raw STGMEDIUM for IDataObject::GetData call
#[repr(C)]
struct RawStgMedium {
    tymed: u32,
    data: *mut c_void, // union field (hGlobal for TYMED_HGLOBAL)
    punk_for_release: *mut c_void,
}

// =============================================================================
// ClassFactory
// =============================================================================

#[repr(C)]
struct ClassFactory {
    vtbl: *const IClassFactoryVtbl,
    ref_count: AtomicU32,
}

static CLASS_FACTORY_VTBL: IClassFactoryVtbl = IClassFactoryVtbl {
    base: IUnknownVtbl {
        QueryInterface: cf_query_interface,
        AddRef: cf_add_ref,
        Release: cf_release,
    },
    CreateInstance: cf_create_instance,
    LockServer: cf_lock_server,
};

unsafe extern "system" fn cf_query_interface(
    this: *mut c_void,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        if ppv.is_null() {
            return E_POINTER;
        }
        *ppv = std::ptr::null_mut();
        let iid = &*riid;
        if *iid == IID_IUNKNOWN || *iid == IID_ICLASSFACTORY {
            *ppv = this;
            cf_add_ref(this);
            return S_OK;
        }
        E_NOINTERFACE
    }
}

unsafe extern "system" fn cf_add_ref(this: *mut c_void) -> u32 {
    unsafe {
        let cf = &*(this as *const ClassFactory);
        cf.ref_count.fetch_add(1, Ordering::SeqCst) + 1
    }
}

unsafe extern "system" fn cf_release(this: *mut c_void) -> u32 {
    unsafe {
        let cf = &*(this as *const ClassFactory);
        let count = cf.ref_count.fetch_sub(1, Ordering::SeqCst) - 1;
        if count == 0 {
            drop(Box::from_raw(this as *mut ClassFactory));
            DLL_REF_COUNT.fetch_sub(1, Ordering::SeqCst);
        }
        count
    }
}

unsafe extern "system" fn cf_create_instance(
    _this: *mut c_void,
    punk_outer: *mut c_void,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        if ppv.is_null() {
            return E_POINTER;
        }
        *ppv = std::ptr::null_mut();
        if !punk_outer.is_null() {
            return CLASS_E_NOAGGREGATION;
        }
        let handler = ContextMenuHandler::new();
        let ptr = Box::into_raw(Box::new(handler));
        let hr = handler_query_interface(ptr as *mut c_void, riid, ppv);
        // Release the initial ref since QI added one
        handler_release(ptr as *mut c_void);
        hr
    }
}

unsafe extern "system" fn cf_lock_server(_this: *mut c_void, lock: i32) -> HRESULT {
    if lock != 0 {
        DLL_REF_COUNT.fetch_add(1, Ordering::SeqCst);
    } else {
        DLL_REF_COUNT.fetch_sub(1, Ordering::SeqCst);
    }
    S_OK
}

// =============================================================================
// ContextMenuHandler - implements IShellExtInit + IContextMenu
// =============================================================================

#[repr(C)]
struct ContextMenuHandler {
    vtbl_init: *const IShellExtInitVtbl,
    vtbl_menu: *const IContextMenuVtbl,
    ref_count: AtomicU32,
    info: std::sync::Mutex<ContextMenuInfo>,
}

static SHELL_EXT_INIT_VTBL: IShellExtInitVtbl = IShellExtInitVtbl {
    base: IUnknownVtbl {
        QueryInterface: handler_query_interface,
        AddRef: handler_add_ref,
        Release: handler_release,
    },
    Initialize: handler_initialize,
};

static CONTEXT_MENU_VTBL: IContextMenuVtbl = IContextMenuVtbl {
    base: IUnknownVtbl {
        QueryInterface: handler_menu_query_interface,
        AddRef: handler_menu_add_ref,
        Release: handler_menu_release,
    },
    QueryContextMenu: handler_query_context_menu,
    InvokeCommand: handler_invoke_command,
    GetCommandString: handler_get_command_string,
};

impl ContextMenuHandler {
    fn new() -> Self {
        DLL_REF_COUNT.fetch_add(1, Ordering::SeqCst);
        Self {
            vtbl_init: &SHELL_EXT_INIT_VTBL,
            vtbl_menu: &CONTEXT_MENU_VTBL,
            ref_count: AtomicU32::new(1),
            info: std::sync::Mutex::new(ContextMenuInfo::default()),
        }
    }
}

/// Recover ContextMenuHandler* from IContextMenu interface pointer.
/// vtbl_menu is the second pointer field in the struct.
unsafe fn handler_from_menu_ptr(this: *mut c_void) -> *mut ContextMenuHandler {
    unsafe {
        (this as *mut u8).sub(std::mem::offset_of!(ContextMenuHandler, vtbl_menu))
            as *mut ContextMenuHandler
    }
}

// --- IShellExtInit IUnknown (primary interface at offset 0) ---

unsafe extern "system" fn handler_query_interface(
    this: *mut c_void,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        if ppv.is_null() {
            return E_POINTER;
        }
        *ppv = std::ptr::null_mut();
        let iid = &*riid;
        let handler = this as *mut ContextMenuHandler;

        if *iid == IID_IUNKNOWN || *iid == IID_ISHELLEXTINIT {
            *ppv = this;
            handler_add_ref(this);
            return S_OK;
        }
        if *iid == IID_ICONTEXTMENU {
            *ppv = std::ptr::addr_of_mut!((*handler).vtbl_menu) as *mut c_void;
            handler_add_ref(this);
            return S_OK;
        }
        E_NOINTERFACE
    }
}

unsafe extern "system" fn handler_add_ref(this: *mut c_void) -> u32 {
    unsafe {
        let handler = &*(this as *const ContextMenuHandler);
        handler.ref_count.fetch_add(1, Ordering::SeqCst) + 1
    }
}

unsafe extern "system" fn handler_release(this: *mut c_void) -> u32 {
    unsafe {
        let handler = this as *mut ContextMenuHandler;
        let count = (*handler).ref_count.fetch_sub(1, Ordering::SeqCst) - 1;
        if count == 0 {
            drop(Box::from_raw(handler));
            DLL_REF_COUNT.fetch_sub(1, Ordering::SeqCst);
        }
        count
    }
}

// --- IContextMenu IUnknown (secondary interface at offset 8) ---

unsafe extern "system" fn handler_menu_query_interface(
    this: *mut c_void,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        let handler = handler_from_menu_ptr(this);
        handler_query_interface(handler as *mut c_void, riid, ppv)
    }
}

unsafe extern "system" fn handler_menu_add_ref(this: *mut c_void) -> u32 {
    unsafe {
        let handler = handler_from_menu_ptr(this);
        handler_add_ref(handler as *mut c_void)
    }
}

unsafe extern "system" fn handler_menu_release(this: *mut c_void) -> u32 {
    unsafe {
        let handler = handler_from_menu_ptr(this);
        handler_release(handler as *mut c_void)
    }
}

// --- IShellExtInit::Initialize ---

unsafe extern "system" fn handler_initialize(
    this: *mut c_void,
    pidl_folder: *const c_void,
    p_data_obj: *mut c_void,
    _hkey_prog_id: isize,
) -> HRESULT {
    unsafe {
        let handler = &*(this as *const ContextMenuHandler);
        let Ok(mut info) = handler.info.lock() else {
            return E_FAIL;
        };
        *info = ContextMenuInfo::default();

        info.timestamp = timestamp();
        info.process_id = std::process::id();

        // Cursor position
        let mut pt = POINT::default();
        let _ = GetCursorPos(&mut pt);
        info.cursor_x = pt.x;
        info.cursor_y = pt.y;

        // Folder path from PIDL
        if !pidl_folder.is_null() {
            let mut buf = [0u16; 260];
            if SHGetPathFromIDListW(pidl_folder as *const _, &mut buf).as_bool() {
                let len = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
                info.folder_path = String::from_utf16_lossy(&buf[..len]);
            }
        }

        // Extract selected files via IDataObject::GetData (raw vtable call)
        if !p_data_obj.is_null() {
            extract_selected_files(p_data_obj, &mut info);
        }

        // Context menus invoked on files (e.g. from HKCR\*) often pass a NULL pidlFolder.
        // We can recover the directory path from the parent of the first selected file.
        if info.folder_path.is_empty() && !info.selected_files.is_empty()
            && let Some(first_file) = info.selected_files.first()
                && let Some(parent) = std::path::Path::new(first_file).parent() {
                    info.folder_path = parent.to_string_lossy().into_owned();
                }

        info.is_background = info.selected_files.is_empty() && !info.folder_path.is_empty();

        // Window information
        let hwnd = GetForegroundWindow();
        info.window_handle = hwnd.0 as usize;
        let mut class_buf = [0u16; 256];
        let class_len = GetClassNameW(hwnd, &mut class_buf);
        if class_len > 0 {
            info.window_class = String::from_utf16_lossy(&class_buf[..class_len as usize]);
        }

        S_OK
    }
}

/// Extract selected file paths from IDataObject using CF_HDROP format.
/// Uses raw COM vtable call to avoid windows crate feature issues with GetData.
unsafe fn extract_selected_files(p_data_obj: *mut c_void, info: &mut ContextMenuInfo) {
    unsafe {
        // IDataObject vtable: [QI, AddRef, Release, GetData, ...]
        // GetData is index 3
        let vtbl_ptr = *(p_data_obj as *const *const usize);

        type GetDataFn = unsafe extern "system" fn(
            *mut c_void,
            *const RawFormatEtc,
            *mut RawStgMedium,
        ) -> HRESULT;
        let get_data: GetDataFn = std::mem::transmute(*(vtbl_ptr.add(3)));

        let fmt = RawFormatEtc {
            cf_format: 15, // CF_HDROP
            ptd: std::ptr::null_mut(),
            dw_aspect: 1, // DVASPECT_CONTENT
            lindex: -1,
            tymed: 1, // TYMED_HGLOBAL
        };
        let mut medium = RawStgMedium {
            tymed: 0,
            data: std::ptr::null_mut(),
            punk_for_release: std::ptr::null_mut(),
        };

        let hr = get_data(p_data_obj, &fmt, &mut medium);
        if hr != S_OK || medium.data.is_null() {
            return;
        }

        let hdrop = HDROP(medium.data);
        let count = DragQueryFileW(hdrop, 0xFFFFFFFF, None);
        info.file_count = count;

        for i in 0..count {
            let len = DragQueryFileW(hdrop, i, None);
            if len > 0 {
                let mut buf = vec![0u16; (len + 1) as usize];
                DragQueryFileW(hdrop, i, Some(&mut buf));
                info.selected_files
                    .push(String::from_utf16_lossy(&buf[..len as usize]));
            }
        }

        // Release the STGMEDIUM by calling ReleaseStgMedium via ole32
        // For TYMED_HGLOBAL with null pUnkForRelease, GlobalFree is sufficient
        if medium.punk_for_release.is_null() && medium.tymed == 1 {
            // TYMED_HGLOBAL - we don't own it if pUnkForRelease is null, do nothing
            // The shell manages the memory
        }
    }
}

// --- IContextMenu methods ---

pub const PIPE_NAME: &str = r"\\.\pipe\rcm_com_pipe";

unsafe extern "system" fn handler_query_context_menu(
    this: *mut c_void,
    _hmenu: isize,
    _index_menu: u32,
    _id_cmd_first: u32,
    _id_cmd_last: u32,
    uflags: u32,
) -> HRESULT {
    unsafe {
        let handler = &*handler_from_menu_ptr(this);
        if let Ok(mut info) = handler.info.lock() {
            if uflags & 0x00000001 != 0 {
                info.event = Event::LeftClickSelect { flags: uflags };
            } else if uflags & 0x00000100 != 0 {
                info.event = Event::ShiftSelect { flags: uflags };
            } else {
                info.event = Event::RightClickMenu { flags: uflags };
            }

            let execute_result = (|| -> crate::error::Result<()> {
                let json_str = serde_json::to_string(&*info)?;

                let mut pipe = std::fs::OpenOptions::new()
                    .write(true)
                    .open(PIPE_NAME)?;
                
                std::io::Write::write_all(&mut pipe, json_str.as_bytes())?;
                    
                Ok(())
            })();
            
            if let Err(err) = execute_result {
                write_error(err);
            }
        }
        HRESULT(0) // 0 items added
    }
}

unsafe extern "system" fn handler_invoke_command(
    _this: *mut c_void,
    _pici: *const c_void,
) -> HRESULT {
    S_OK
}

unsafe extern "system" fn handler_get_command_string(
    _this: *mut c_void,
    _id_cmd: usize,
    _u_type: u32,
    _preserved: *const u32,
    _psz_name: *mut u8,
    _cch_max: u32,
) -> HRESULT {
    E_NOTIMPL
}

// =============================================================================
// DLL entry points
// =============================================================================

#[unsafe(no_mangle)]
unsafe extern "system" fn DllMain(hinstance: HMODULE, reason: u32, _reserved: *mut c_void) -> i32 {
    // BOOL = i32; TRUE = 1
    unsafe {
        if reason == DLL_PROCESS_ATTACH {
            DLL_MODULE.store(hinstance.0 as usize, Ordering::SeqCst);
            let _ = DisableThreadLibraryCalls(hinstance);
        }
        1 // TRUE
    }
}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllGetClassObject(
    rclsid: *const GUID,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        if ppv.is_null() {
            return E_POINTER;
        }
        *ppv = std::ptr::null_mut();

        if *rclsid != CLSID_RCM {
            return CLASS_E_CLASSNOTAVAILABLE;
        }

        let factory = Box::new(ClassFactory {
            vtbl: &CLASS_FACTORY_VTBL,
            ref_count: AtomicU32::new(1),
        });
        DLL_REF_COUNT.fetch_add(1, Ordering::SeqCst);

        let ptr = Box::into_raw(factory) as *mut c_void;
        let hr = cf_query_interface(ptr, riid, ppv);
        // Release initial ref (QI already added one)
        cf_release(ptr);
        hr
    }
}

#[unsafe(no_mangle)]
extern "system" fn DllCanUnloadNow() -> HRESULT {
    if DLL_REF_COUNT.load(Ordering::SeqCst) == 0 {
        S_OK
    } else {
        S_FALSE
    }
}
