#![allow(non_snake_case)]

// ── public modules ───────────────────────────────────────────────────────
pub mod cmd;
pub mod consts;
pub mod error;
pub mod server;

// ── private modules ──────────────────────────────────────────────────────
pub(crate) mod com;
pub(crate) mod helpers;
pub(crate) mod hooks;
pub(crate) mod types;

// ── public re-exports ────────────────────────────────────────────────────
pub use consts::PIPE_NAME;
pub use types::{ContextMenuInfo, Event};

use std::ffi::c_void;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::Win32::Foundation::*;
use windows::Win32::System::LibraryLoader::DisableThreadLibraryCalls;
use windows::Win32::System::SystemServices::*;
use windows::core::{GUID, HRESULT};

use crate::com::handler::ContextMenuHandler;
use crate::com::vtable::IClassFactoryVtbl;
use crate::com::vtable::IUnknownVtbl;
use crate::consts::*;
use crate::helpers::DLL_MODULE;

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
            helpers::DLL_REF_COUNT.fetch_sub(1, Ordering::Relaxed);
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
        // QueryInterface via handler IShellExtInit vtable
        let vtbl = &*(*ptr).vtbl_init;
        let qi = vtbl.base.QueryInterface;
        let hr = qi(ptr as *mut c_void, riid, ppv);
        // Release the initial reference since QI added one
        let release = vtbl.base.Release;
        release(ptr as *mut c_void);
        hr
    }
}

unsafe extern "system" fn cf_lock_server(_this: *mut c_void, lock: i32) -> HRESULT {
    if lock != 0 {
        helpers::DLL_REF_COUNT.fetch_add(1, Ordering::Relaxed);
    } else {
        helpers::DLL_REF_COUNT.fetch_sub(1, Ordering::Relaxed);
    }
    S_OK
}

// =============================================================================
// DLL entry points
// =============================================================================

#[unsafe(no_mangle)]
unsafe extern "system" fn DllMain(hinstance: HMODULE, reason: u32, _reserved: *mut c_void) -> i32 {
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
        helpers::DLL_REF_COUNT.fetch_add(1, Ordering::Relaxed);

        let ptr = Box::into_raw(factory) as *mut c_void;
        let hr = cf_query_interface(ptr, riid, ppv);
        // Release initial ref (QI already added one)
        cf_release(ptr);
        hr
    }
}

#[unsafe(no_mangle)]
extern "system" fn DllCanUnloadNow() -> HRESULT {
    if helpers::DLL_REF_COUNT.load(Ordering::Relaxed) == 0 {
        S_OK
    } else {
        S_FALSE
    }
}
