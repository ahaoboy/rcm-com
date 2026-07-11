//! Raw COM IDataObject helpers — extract selected file paths from Explorer
//! using CF_HDROP and CFSTR_SHELLIDLIST formats via raw vtable calls.

use std::ffi::c_void;

use windows::Win32::Foundation::*;
use windows::Win32::System::DataExchange::RegisterClipboardFormatW;
use windows::Win32::UI::Shell::{DragQueryFileW, HDROP, SHGetPathFromIDListW, ILCombine, ILFree};
use windows::core::{HRESULT, PCWSTR};

use crate::helpers::write_log;
use crate::types::ContextMenuInfo;

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

/// CIDA header preceding PIDL data in CFSTR_SHELLIDLIST format.
#[repr(C)]
struct CidaHeader {
    cidl: u32,         // count of child PIDLs (not including parent)
    aoffset: [u32; 1], // variable-length: aoffset[0] → parent, aoffset[1..] → children
}

/// Extract selected file / namespace-item paths from IDataObject.
///
/// Tries **CF_HDROP** first (works for filesystem files and folders).  If no
/// files are obtained falls back to **CFSTR_SHELLIDLIST** which is used for
/// namespace objects such as drives shown in "This PC" (My Computer).
///
/// Diagnostic entries are written to `err.txt` so you can verify which format
/// was used and how many files were extracted.
pub(crate) unsafe fn extract_selected_files(p_data_obj: *mut c_void, info: &mut ContextMenuInfo) {
    unsafe {
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

        // ── 1. Try CF_HDROP (filesystem paths) ──────────────────────────
        if try_hdrop(p_data_obj, get_data, info) {
            write_log(format!("dataobj: CF_HDROP → {} file(s)", info.files.len()));
            return;
        }

        // ── 2. Fallback: CFSTR_SHELLIDLIST (namespace items like drives) ──
        try_shell_idlist(p_data_obj, get_data, info);
        write_log(format!(
            "dataobj: CFSTR_SHELLIDLIST → {} item(s)",
            info.files.len()
        ));
    }
}

// =============================================================================
// CF_HDROP extraction
// =============================================================================

unsafe fn try_hdrop(
    p_data_obj: *mut c_void,
    get_data: unsafe extern "system" fn(
        *mut c_void,
        *const RawFormatEtc,
        *mut RawStgMedium,
    ) -> HRESULT,
    info: &mut ContextMenuInfo,
) -> bool {
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

    let hr = unsafe { get_data(p_data_obj, &fmt, &mut medium) };
    if hr != S_OK || medium.data.is_null() {
        unsafe {
            release_stg_medium(&mut medium);
        }
        return false;
    }

    let hdrop = HDROP(medium.data);
    let count = unsafe { DragQueryFileW(hdrop, 0xFFFFFFFF, None) };

    for i in 0..count {
        let len = unsafe { DragQueryFileW(hdrop, i, None) };
        if len > 0 {
            let mut buf = vec![0u16; (len + 1) as usize];
            unsafe { DragQueryFileW(hdrop, i, Some(&mut buf)) };
            let name = String::from_utf16_lossy(&buf[..len as usize]);
            info.files.push(name);
        }
    }

    unsafe {
        release_stg_medium(&mut medium);
    }
    !info.files.is_empty()
}

// =============================================================================
// CFSTR_SHELLIDLIST extraction (namespace items — drives in "This PC", etc.)
// =============================================================================

unsafe fn try_shell_idlist(
    p_data_obj: *mut c_void,
    get_data: unsafe extern "system" fn(
        *mut c_void,
        *const RawFormatEtc,
        *mut RawStgMedium,
    ) -> HRESULT,
    info: &mut ContextMenuInfo,
) {
    // Register (or retrieve) the clipboard format for "Shell IDList Array".
    let name_wide: Vec<u16> = "Shell IDList Array\0".encode_utf16().collect();
    let cf = unsafe { RegisterClipboardFormatW(PCWSTR::from_raw(name_wide.as_ptr())) };
    if cf == 0 {
        write_log("dataobj: RegisterClipboardFormatW('Shell IDList Array') failed");
        return;
    }

    let fmt = RawFormatEtc {
        cf_format: cf as u16,
        ptd: std::ptr::null_mut(),
        dw_aspect: 1,
        lindex: -1,
        tymed: 1, // TYMED_HGLOBAL
    };
    let mut medium = RawStgMedium {
        tymed: 0,
        data: std::ptr::null_mut(),
        punk_for_release: std::ptr::null_mut(),
    };

    let hr = unsafe { get_data(p_data_obj, &fmt, &mut medium) };
    if hr != S_OK || medium.data.is_null() {
        write_log(format!(
            "dataobj: GetData(CFSTR_SHELLIDLIST) failed hr={hr:?}"
        ));
        unsafe {
            release_stg_medium(&mut medium);
        }
        return;
    }

    // The data is a CIDA header followed by PIDLs.
    let cida = medium.data as *const CidaHeader;
    let cidl = unsafe { (*cida).cidl };

    // Get the parent folder PIDL (relative to Desktop, i.e. absolute)
    let parent_offset = unsafe { (*cida).aoffset[0] };
    let parent_pidl = unsafe { medium.data.byte_add(parent_offset as usize) } as *const c_void;

    let mut count = 0u32;
    for i in 0..cidl {
        // aoffset[0] is the parent folder PIDL; aoffset[1..cidl] are children (relative to parent).
        let offset = unsafe { *(*cida).aoffset.as_ptr().add(i as usize + 1) };
        let child_pidl = unsafe { medium.data.byte_add(offset as usize) } as *const c_void;

        // Combine parent PIDL and child PIDL to get absolute PIDL
        let combined_pidl = unsafe { ILCombine(Some(parent_pidl as *const _), Some(child_pidl as *const _)) };

        if !combined_pidl.is_null() {
            let mut buf = [0u16; 260]; // MAX_PATH
            if unsafe { SHGetPathFromIDListW(combined_pidl as *const _, &mut buf) }.as_bool() {
                let len = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
                let path = String::from_utf16_lossy(&buf[..len]);
                info.files.push(path);
                count += 1;
            }
            unsafe {
                ILFree(Some(combined_pidl as *const _));
            }
        }
    }

    if count == 0 && cidl > 0 {
        write_log(format!(
            "dataobj: SHGetPathFromIDListW failed for all {cidl} PIDL(s)"
        ));
    }

    unsafe {
        release_stg_medium(&mut medium);
    }
}

/// Release an STGMEDIUM structure.
/// For TYMED_HGLOBAL (1), frees the global memory handle.
/// For other types with a release pointer, calls IUnknown::Release.
unsafe fn release_stg_medium(medium: &mut RawStgMedium) {
    // TYMED_HGLOBAL — free the global memory
    if medium.tymed == 1 && !medium.data.is_null() {
        unsafe {
            // GlobalFree is in kernel32; call it via FFI
            unsafe extern "system" {
                fn GlobalFree(hMem: *mut c_void) -> *mut c_void;
            }
            GlobalFree(medium.data);
        }
    }
    // Release any punkForRelease
    if !medium.punk_for_release.is_null() {
        unsafe {
            let unknown: *mut *const crate::com::vtable::IUnknownVtbl =
                medium.punk_for_release as *mut _;
            let vtbl = *unknown;
            if !vtbl.is_null() {
                let release = (*vtbl).Release;
                release(medium.punk_for_release);
            }
        }
    }
}
