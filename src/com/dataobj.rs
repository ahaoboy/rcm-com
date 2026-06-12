//! Raw COM IDataObject helpers — extract selected file paths from Explorer
//! using CF_HDROP format via raw vtable calls.

use std::ffi::c_void;

use windows::Win32::Foundation::*;
use windows::Win32::UI::Shell::{DragQueryFileW, HDROP};
use windows::core::HRESULT;

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

/// Extract selected file paths from IDataObject using CF_HDROP format.
/// Uses raw COM vtable call to avoid windows crate feature issues with GetData.
pub(crate) unsafe fn extract_selected_files(p_data_obj: *mut c_void, info: &mut ContextMenuInfo) {
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
            release_stg_medium(&mut medium);
            return;
        }

        let hdrop = HDROP(medium.data);
        let count = DragQueryFileW(hdrop, 0xFFFFFFFF, None);

        for i in 0..count {
            let len = DragQueryFileW(hdrop, i, None);
            if len > 0 {
                let mut buf = vec![0u16; (len + 1) as usize];
                DragQueryFileW(hdrop, i, Some(&mut buf));
                let name = String::from_utf16_lossy(&buf[..len as usize]);
                info.files.push(name);
            }
        }

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
