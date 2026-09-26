//! Raw COM IDataObject helpers — extract selected file paths from Explorer
//! using CF_HDROP format via raw vtable calls.

use std::ffi::c_void;

use windows::Win32::Foundation::*;
use windows::Win32::System::Com::STGMEDIUM;
use windows::Win32::System::Ole::ReleaseStgMedium;
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

/// Extract selected file paths from IDataObject using CF_HDROP format.
pub(crate) unsafe fn extract_selected_files(p_data_obj: *mut c_void, info: &mut ContextMenuInfo) {
    unsafe {
        let vtbl = *(p_data_obj as *const *const usize);
        if vtbl.is_null() {
            return;
        }

        type GetDataFn =
            unsafe extern "system" fn(*mut c_void, *const RawFormatEtc, *mut STGMEDIUM) -> HRESULT;
        let get_data: GetDataFn = std::mem::transmute(*(vtbl.add(3)));

        let fmt = RawFormatEtc {
            cf_format: 15, // CF_HDROP
            ptd: std::ptr::null_mut(),
            dw_aspect: 1, // DVASPECT_CONTENT
            lindex: -1,
            tymed: 1, // TYMED_HGLOBAL
        };
        let mut medium = STGMEDIUM::default();

        let hr = get_data(p_data_obj, &fmt, &mut medium);
        if hr != S_OK {
            // GetData failed — per contract nothing was allocated.
            return;
        }

        let hdrop = HDROP(medium.u.hGlobal.0);
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

        // Use the OS routine so the medium is released with the exact semantics
        // required: it either frees the `hGlobal` *or* calls `pUnkForRelease`'s
        // `Release` — never both. The previous hand-rolled version could free an
        // HGLOBAL *and* release the IUnknown for the same medium.
        ReleaseStgMedium(&mut medium);
    }
}
