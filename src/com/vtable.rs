//! Raw COM vtable definitions (C ABI compatible) for IUnknown, IShellExtInit,
//! IContextMenu, and IClassFactory.

use std::ffi::c_void;

use windows::core::{GUID, HRESULT};

// =============================================================================
// IUnknown
// =============================================================================

#[repr(C)]
pub(crate) struct IUnknownVtbl {
    pub(crate) QueryInterface:
        unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> HRESULT,
    pub(crate) AddRef: unsafe extern "system" fn(*mut c_void) -> u32,
    pub(crate) Release: unsafe extern "system" fn(*mut c_void) -> u32,
}

// =============================================================================
// IShellExtInit
// =============================================================================

#[repr(C)]
pub(crate) struct IShellExtInitVtbl {
    pub(crate) base: IUnknownVtbl,
    pub(crate) Initialize: unsafe extern "system" fn(
        *mut c_void,   // this
        *const c_void, // pidlFolder (PCIDLIST_ABSOLUTE)
        *mut c_void,   // pDataObj (IDataObject*)
        isize,         // hKeyProgID (HKEY)
    ) -> HRESULT,
}

// =============================================================================
// IContextMenu
// =============================================================================

#[repr(C)]
pub(crate) struct IContextMenuVtbl {
    pub(crate) base: IUnknownVtbl,
    pub(crate) QueryContextMenu: unsafe extern "system" fn(*mut c_void, isize, u32, u32, u32, u32) -> HRESULT,
    pub(crate) InvokeCommand: unsafe extern "system" fn(*mut c_void, *const c_void) -> HRESULT,
    pub(crate) GetCommandString:
        unsafe extern "system" fn(*mut c_void, usize, u32, *const u32, *mut u8, u32) -> HRESULT,
}

// =============================================================================
// IClassFactory
// =============================================================================

#[repr(C)]
pub(crate) struct IClassFactoryVtbl {
    pub(crate) base: IUnknownVtbl,
    pub(crate) CreateInstance: unsafe extern "system" fn(
        *mut c_void,
        *mut c_void,
        *const GUID,
        *mut *mut c_void,
    ) -> HRESULT,
    pub(crate) LockServer: unsafe extern "system" fn(*mut c_void, i32) -> HRESULT,
}
