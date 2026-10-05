//! Project-wide constants: CLSID, IIDs, handler name, and pipe name.

use windows::core::GUID;

// UUID v5 of "https://github.com/ahaoboy/rcm-com.git"
pub const CLSID_STR: &str = "{F96C1A16-22B8-5B5F-AEF4-B5E45A312B00}";
pub const CLSID_RCM: GUID = GUID::from_u128(0xF96C1A16_22B8_5B5F_AEF4_B5E45A312B00);

pub const IID_IUNKNOWN: GUID = GUID::from_u128(0x00000000_0000_0000_C000_000000000046);
pub const IID_ICLASSFACTORY: GUID = GUID::from_u128(0x00000001_0000_0000_C000_000000000046);
pub const IID_ISHELLEXTINIT: GUID = GUID::from_u128(0x000214E8_0000_0000_C000_000000000046);
pub const IID_ICONTEXTMENU: GUID = GUID::from_u128(0x000214E4_0000_0000_C000_000000000046);

pub const HANDLER_NAME: &str = "RcmContextMenu";

/// User-scoped settings key (`HKEY_CURRENT_USER\Software\RcmCom`).
///
/// Holds persisted preferences such as the log level and the Shift+right-click
/// behaviour, read by both the CLI and the DLL.
pub const CONFIG_REG_KEY: &str = r"Software\RcmCom";

/// Config value name for the persisted log level (see [`crate::logging`]).
pub const CONFIG_LOG_LEVEL: &str = "LogLevel";

/// Event pipe: the **listener** hosts it and each loaded shell-extension
/// instance connects as a client.
///
/// The listener owns the name, so restarting Explorer does not destroy this
/// channel — and because every Explorer process connects separately, events
/// from all of them reach the one listener. See [`crate::events`].
pub const EVENT_PIPE_NAME: &str = r"\\.\pipe\rcm_com";

/// Control pipe: the shell extension hosts it and the `rcm` CLI connects as a
/// client. Short-lived request/response traffic only; see [`crate::pipe`].
pub const CONTROL_PIPE_NAME: &str = r"\\.\pipe\rcm_com_control";
