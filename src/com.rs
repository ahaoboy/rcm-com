//! COM shell extension handler — implements IShellExtInit and IContextMenu.
//!
//! The `ContextMenuHandler` struct captures right-click context data from
//! Explorer and sends it over a named pipe to the listening process. It also
//! blocks the default Windows 11 context menu via a WH_CBT hook.

use std::ffi::c_void;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::Win32::Foundation::*;
use windows::Win32::UI::Shell::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::{GUID, HRESULT};

use crate::consts::*;
use crate::helpers::{self, DLL_REF_COUNT};
use crate::hooks::install_cbt_menu_blocker;
use crate::types::{ContextMenuInfo, Event};

// =============================================================================
// Raw COM vtable definitions (C ABI compatible)
// =============================================================================

#[repr(C)]
pub(crate) struct IUnknownVtbl {
    pub(crate) QueryInterface:
        unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> HRESULT,
    AddRef: unsafe extern "system" fn(*mut c_void) -> u32,
    pub(crate) Release: unsafe extern "system" fn(*mut c_void) -> u32,
}

#[repr(C)]
pub(crate) struct IShellExtInitVtbl {
    pub(crate) base: IUnknownVtbl,
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

/// Raw FORMATETC for IDataObject::GetData call.
#[repr(C)]
struct RawFormatEtc {
    cf_format: u16,
    ptd: *mut c_void,
    dw_aspect: u32,
    lindex: i32,
    tymed: u32,
}

/// Raw STGMEDIUM for IDataObject::GetData call.
#[repr(C)]
struct RawStgMedium {
    tymed: u32,
    data: *mut c_void, // union field (hGlobal for TYMED_HGLOBAL)
    punk_for_release: *mut c_void,
}

// =============================================================================
// ContextMenuHandler — implements IShellExtInit + IContextMenu
// =============================================================================

#[repr(C)]
pub(crate) struct ContextMenuHandler {
    pub(crate) vtbl_init: *const IShellExtInitVtbl,
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
    pub(crate) fn new() -> Self {
        DLL_REF_COUNT.fetch_add(1, Ordering::Relaxed);
        Self {
            vtbl_init: &SHELL_EXT_INIT_VTBL,
            vtbl_menu: &CONTEXT_MENU_VTBL,
            ref_count: AtomicU32::new(1),
            info: std::sync::Mutex::new(ContextMenuInfo::default()),
        }
    }
}

// =============================================================================
// Pointer arithmetic helpers
// =============================================================================

/// Recover `ContextMenuHandler*` from an `IContextMenu` interface pointer.
/// `vtbl_menu` is the second pointer field in the struct (after `vtbl_init`).
unsafe fn handler_from_menu_ptr(this: *mut c_void) -> *mut ContextMenuHandler {
    unsafe {
        (this as *mut u8).sub(std::mem::offset_of!(ContextMenuHandler, vtbl_menu))
            as *mut ContextMenuHandler
    }
}

// =============================================================================
// IShellExtInit IUnknown (primary interface at offset 0)
// =============================================================================

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

// =============================================================================
// IContextMenu IUnknown (secondary interface at offset 8)
// =============================================================================

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

// =============================================================================
// IShellExtInit::Initialize
// =============================================================================

unsafe extern "system" fn handler_initialize(
    this: *mut c_void,
    pidl_folder: *const c_void,
    p_data_obj: *mut c_void,
    _hkey_prog_id: isize,
) -> HRESULT {
    unsafe {
        // Install the CBT hook on first call to block the system menu.
        install_cbt_menu_blocker();

        let handler = &*(this as *const ContextMenuHandler);
        let Ok(mut info) = handler.info.lock() else {
            return E_FAIL;
        };
        *info = ContextMenuInfo::default();

        info.timestamp = helpers::timestamp();
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

        // Context menus invoked on files (e.g. from HKCR\*) often pass a NULL
        // pidlFolder. Recover the directory from the first selected file's parent.
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

// =============================================================================
// IContextMenu methods
// =============================================================================

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

            // Send the info over the named pipe.
            let execute_result = (|| -> crate::error::Result<()> {
                let json_str = serde_json::to_string(&*info)?;
                let mut pipe = std::fs::OpenOptions::new()
                    .write(true)
                    .open(crate::consts::PIPE_NAME)?;
                std::io::Write::write_all(&mut pipe, json_str.as_bytes())?;
                Ok(())
            })();

            if let Err(err) = execute_result {
                helpers::write_log(err);
            }
        }

        // Block the system context menu from appearing.
        // The CBT hook (installed during IShellExtInit::Initialize) intercepts
        // TrackPopupMenu before any menu window is created.
        // We also return E_FAIL to tell the shell we contributed nothing.
        let should_block = {
            #[cfg(feature = "config")]
            {
                crate::config::get_config().block_win11_menu
            }
            #[cfg(not(feature = "config"))]
            {
                true
            }
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
// IDataObject helpers (raw vtable)
// =============================================================================

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
    #[link(name = "ole32")]
    unsafe extern "system" {
        fn ReleaseStgMedium(pmedium: *mut RawStgMedium);
    }
    unsafe {
        ReleaseStgMedium(medium);
    }
}
