use windows::core::GUID;

// python -c "import uuid; print(str(uuid.uuid5(uuid.NAMESPACE_URL, 'https://github.com/ahaoboy/rcm-com.git')).upper())"
// UUID v5 of "https://github.com/ahaoboy/rcm-com.git"
pub const CLSID_STR: &str = "{F96C1A16-22B8-5B5F-AEF4-B5E45A312B00}";
pub const CLSID_RCM: GUID = GUID::from_u128(0xF96C1A16_22B8_5B5F_AEF4_B5E45A312B00);

pub const IID_IUNKNOWN: GUID = GUID::from_u128(0x00000000_0000_0000_C000_000000000046);
pub const IID_ICLASSFACTORY: GUID = GUID::from_u128(0x00000001_0000_0000_C000_000000000046);
pub const IID_ISHELLEXTINIT: GUID = GUID::from_u128(0x000214E8_0000_0000_C000_000000000046);
pub const IID_ICONTEXTMENU: GUID = GUID::from_u128(0x000214E4_0000_0000_C000_000000000046);

pub const HANDLER_NAME: &str = "RcmContextMenu";
