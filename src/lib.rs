#![allow(non_snake_case)]

use serde::{Deserialize, Serialize};
use std::ffi::c_void;
use std::io::Write;
use std::sync::atomic::{AtomicIsize, AtomicU32, AtomicUsize, Ordering};
pub mod cmd;
pub mod consts;
pub mod error;
pub mod server;
#[cfg(feature = "config")]
pub mod config;
use windows::Win32::Foundation::*;
use windows::Win32::System::LibraryLoader::*;
use windows::Win32::System::SystemServices::*;
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::Shell::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::{GUID, HRESULT};
use chrono::Utc;
// =============================================================================
// Constants
// =============================================================================

use crate::consts::*;

static DLL_MODULE: AtomicUsize = AtomicUsize::new(0);
static DLL_REF_COUNT: AtomicU32 = AtomicU32::new(0);

// ---------------------------------------------------------------------------
// WH_CBT hook — prevents context menu windows from being created
// ---------------------------------------------------------------------------
//
// When a popup menu (window class #32768) is about to be created, the
// HCBT_CREATEWND notification fires.  Returning non-zero from the CBT
// hook procedure PREVENTS the window from being created entirely – the
// menu never appears, no flash, no timing issues.
//
// Unlike inline hooks, this approach does NOT modify any code in memory.
// The hook is installed per-right-click in IShellExtInit::Initialize and
// unhooks itself on the first popup-menu creation it catches.

/// Handle of the active WH_CBT hook (0 when not installed).
/// Atomic to prevent data races between the hook callback and the
/// installer running on different COM apartment threads.
static CBT_HOOK_HANDLE: AtomicIsize = AtomicIsize::new(0);

/// CBT hook procedure — called before windows are created on our thread.
///
/// When code == HCBT_CREATEWND (3), lparam points to a CBT_CREATEWNDW
/// whose lpcs->lpszClass identifies the window class.  A popup menu has
/// class atom 32768 (0x8000).  We return 1 to prevent its creation.
unsafe extern "system" fn cbt_hook_proc(
    code: i32,
    _wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
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
                    // Then unhook – we only need to catch the first one.
                    let hook = CBT_HOOK_HANDLE.swap(0, Ordering::Acquire);
                    if hook != 0 {
                        let _ = UnhookWindowsHookEx(HHOOK(hook as *mut c_void));
                    }
                    return LRESULT(1);
                }
            }
        }
    }
    // Not our target – pass to the next hook in the chain.
    unsafe { CallNextHookEx(None, code, _wparam, lparam) }
}

/// Install a thread-local WH_CBT hook that blocks the next popup menu.
fn install_cbt_menu_blocker() {
    unsafe {
        // Already installed – nothing to do.
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

// =============================================================================
// Helpers
// =============================================================================

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


/// Returns the full path to the DLL file itself, not just its directory.
#[cfg_attr(not(feature = "config"), allow(dead_code))]
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
    Some(std::path::PathBuf::from(String::from_utf16_lossy(&buf[..len])))
}

/// Write a diagnostic log entry.  When the `config` feature is active and
/// `config.log == true`, the entry is appended to `{dll_stem}.log`.
/// Otherwise it is written to `err.txt` (error-only fallback).
fn write_log(err: impl std::fmt::Display) {
    #[cfg(feature = "config")]
    {
        if crate::config::get_config().log
            && let Some(path) = crate::config::log_path()
                && let Ok(mut file) = std::fs::OpenOptions::new()
                    .create(true)
                    .append(true)
                    .open(&path)
            {
                let _ = writeln!(file, "[{}] {}", timestamp(), err);
                return;
            }
    }
    // Fallback: always write to err.txt so we don't lose error visibility.
    if let Some(path) = dll_dir().map(|d| d.join("err.txt"))
        && let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(path)
    {
        let _ = writeln!(file, "[{}] {}", timestamp(), err);
    }
}

/// Compute a UTC timestamp string for log entries.
fn timestamp() -> String {
    Utc::now().format("%Y-%m-%d %H:%M:%S UTC").to_string()
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
        cf.ref_count.fetch_add(1, Ordering::Relaxed) + 1
    }
}

unsafe extern "system" fn cf_release(this: *mut c_void) -> u32 {
    unsafe {
        let cf = &*(this as *const ClassFactory);
        let count = cf.ref_count.fetch_sub(1, Ordering::Relaxed) - 1;
        if count == 0 {
            drop(Box::from_raw(this as *mut ClassFactory));
            DLL_REF_COUNT.fetch_sub(1, Ordering::Relaxed);
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
        DLL_REF_COUNT.fetch_add(1, Ordering::Relaxed);
    } else {
        DLL_REF_COUNT.fetch_sub(1, Ordering::Relaxed);
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
        DLL_REF_COUNT.fetch_add(1, Ordering::Relaxed);
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
        handler.ref_count.fetch_add(1, Ordering::Relaxed) + 1
    }
}

unsafe extern "system" fn handler_release(this: *mut c_void) -> u32 {
    unsafe {
        let handler = this as *mut ContextMenuHandler;
        let count = (*handler).ref_count.fetch_sub(1, Ordering::Relaxed) - 1;
        if count == 0 {
            drop(Box::from_raw(handler));
            DLL_REF_COUNT.fetch_sub(1, Ordering::Relaxed);
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
        // Install the inline hook on TrackPopupMenu once (first call).
        // This intercepts ALL subsequent TrackPopupMenu calls in the
        // explorer process, preventing the system menu from appearing.
        install_cbt_menu_blocker();

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
        let vtbl = *(p_data_obj as *const *const usize);
        if vtbl.is_null() {
            return;
        }

        type GetDataFn = unsafe extern "system" fn(
            *mut c_void,
            *const RawFormatEtc,
            *mut RawStgMedium,
        ) -> HRESULT;
        let get_data: GetDataFn = std::mem::transmute(*(vtbl.add(3)));

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
            // Release STGMEDIUM even on failure (it may hold partial data).
            release_stg_medium(&mut medium);
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

        // Release the STGMEDIUM — the shell allocated it, we must free it.
        release_stg_medium(&mut medium);
    }
}

/// Call ole32!ReleaseStgMedium to free resources held by an STGMEDIUM.
unsafe fn release_stg_medium(medium: &mut RawStgMedium) {
    // Declare the external function from ole32.dll.
    #[link(name = "ole32")]
    unsafe extern "system" {
        fn ReleaseStgMedium(pmedium: *mut RawStgMedium);
    }
    unsafe {
        ReleaseStgMedium(medium);
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
                write_log(err);
            }
        }

        // -----------------------------------------------------------------
        // Block the system context menu from appearing.
        //
        // The inline hook on TrackPopupMenu (installed during
        // IShellExtInit::Initialize) intercepts the API call before any
        // menu window is created, returning FALSE immediately.
        // So the menu never appears — no flash, no timing issues.
        //
        // We still return E_FAIL to tell the shell we contributed nothing.
        // -----------------------------------------------------------------
        let should_block = {
            #[cfg(feature = "config")]
            { crate::config::get_config().block_win11_menu }
            #[cfg(not(feature = "config"))]
            { true }
        };

        if should_block {
            return E_FAIL;
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
            DLL_MODULE.store(hinstance.0 as usize, Ordering::Release);
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
        DLL_REF_COUNT.fetch_add(1, Ordering::Relaxed);

        let ptr = Box::into_raw(factory) as *mut c_void;
        let hr = cf_query_interface(ptr, riid, ppv);
        // Release initial ref (QI already added one)
        cf_release(ptr);
        hr
    }
}

#[unsafe(no_mangle)]
extern "system" fn DllCanUnloadNow() -> HRESULT {
    if DLL_REF_COUNT.load(Ordering::Relaxed) == 0 {
        S_OK
    } else {
        S_FALSE
    }
}
