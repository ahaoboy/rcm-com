//! ContextMenuHandler — implements IShellExtInit + IContextMenu.
//!
//! Captures right-click context data from Explorer and sends it over a named
//! pipe to the listening process. Always blocks the native context menu from
//! appearing via a WH_CBT hook (Win10) and by returning E_FAIL from
//! QueryContextMenu (Win11).

use std::ffi::c_void;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::Win32::Foundation::*;
use windows::Win32::UI::Shell::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::{GUID, HRESULT};

use crate::com::dataobj::extract_selected_files;
use crate::com::vtable::{IContextMenuVtbl, IShellExtInitVtbl, IUnknownVtbl};
use crate::consts::*;
use crate::helpers::{self, DLL_REF_COUNT};
use crate::hooks::install_cbt_menu_blocker;
use crate::types::{ContextMenuInfo, Event};

// =============================================================================
// ContextMenuHandler struct
// =============================================================================

#[repr(C)]
pub(crate) struct ContextMenuHandler {
    pub(crate) vtbl_init: *const IShellExtInitVtbl,
    pub(crate) vtbl_menu: *const IContextMenuVtbl,
    ref_count: AtomicU32,
    pub(crate) info: std::sync::Mutex<ContextMenuInfo>,
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
// Pointer arithmetic helper
// =============================================================================

/// Recover `ContextMenuHandler*` from an `IContextMenu` interface pointer.
/// `vtbl_menu` is the second pointer field in the struct (after `vtbl_init`).
pub(crate) unsafe fn handler_from_menu_ptr(this: *mut c_void) -> *mut ContextMenuHandler {
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
        // Always install the CBT hook to block native context menu windows.
        // This works for both Win10 (TrackPopupMenu) and Win11 (new menu).
        install_cbt_menu_blocker();

        let handler = &*(this as *const ContextMenuHandler);
        let Ok(mut info) = handler.info.lock() else {
            return E_FAIL;
        };
        *info = ContextMenuInfo::default();

        info.ts = helpers::timestamp();
        info.pid = std::process::id();

        // Cursor position
        let mut pt = POINT::default();
        let _ = GetCursorPos(&mut pt);
        info.x = pt.x;
        info.y = pt.y;

        // Folder path from PIDL
        if !pidl_folder.is_null() {
            let mut buf = [0u16; 260];
            if SHGetPathFromIDListW(pidl_folder as *const _, &mut buf).as_bool() {
                let len = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
                info.dir = String::from_utf16_lossy(&buf[..len]);
            }
        }

        // Extract selected files via IDataObject::GetData (raw vtable call)
        if !p_data_obj.is_null() {
            extract_selected_files(p_data_obj, &mut info);
        }

        // Context menus invoked on files (e.g. from HKCR\*) often pass a NULL
        // pidlFolder. Recover the directory from the first selected file's parent.
        if info.dir.is_empty() && !info.files.is_empty()
            && let Some(first_file) = info.files.first()
                && let Some(parent) = std::path::Path::new(first_file).parent() {
                    info.dir = parent.to_string_lossy().into_owned();
                }

        info.bg = info.files.is_empty() && !info.dir.is_empty();

        // Window information
        let hwnd = GetForegroundWindow();
        info.hwnd = hwnd.0 as usize;
        let mut class_buf = [0u16; 256];
        let class_len = GetClassNameW(hwnd, &mut class_buf);
        if class_len > 0 {
            info.class = String::from_utf16_lossy(&class_buf[..class_len as usize]);
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
            // Determine event type from flags
            if uflags & 0x00000001 != 0 {
                info.event = Event::Click { flags: uflags };
            } else if uflags & 0x00000100 != 0 {
                info.event = Event::Shift { flags: uflags };
            } else {
                info.event = Event::Menu { flags: uflags };
            }

            // Send the info over the named pipe to the listening process.
            let send_result = (|| -> crate::error::Result<()> {
                let json_str = serde_json::to_string(&*info)?;
                let mut pipe = std::fs::OpenOptions::new()
                    .write(true)
                    .open(crate::consts::PIPE_NAME)?;
                std::io::Write::write_all(&mut pipe, json_str.as_bytes())?;
                Ok(())
            })();

            if let Err(err) = send_result {
                helpers::write_log(err);
            }
        }

        // Always block the native context menu — both Win10 and Win11.
        // The CBT hook (installed during IShellExtInit::Initialize) intercepts
        // TrackPopupMenu before any menu window is created. Returning E_FAIL
        // tells the shell we contributed no items.
        E_FAIL
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
