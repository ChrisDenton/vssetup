use core::fmt;
use core::mem::ManuallyDrop;
use windows_result::HRESULT;
use windows_strings::BSTR;

// Windows.Win32.Foundation.FILETIME
#[repr(C)]
#[derive(Default, Debug, Clone, Copy)]
pub struct FILETIME {
    pub dwLowDateTime: u32,
    pub dwHighDateTime: u32,
}

impl FILETIME {
    pub fn as_u64(&self) -> u64 {
        ((self.dwHighDateTime as u64) << 32) | (self.dwLowDateTime as u64)
    }
}

// Windows.Win32.System.Com.SAFEARRAYBOUND
#[repr(C)]
pub struct SAFEARRAYBOUND {
    pub cElements: u32,
    pub lLbound: i32,
}
// Windows.Win32.System.Com.SAFEARRAY
#[repr(C)]
pub struct SAFEARRAY {
    pub cDims: u16,
    pub fFeatures: u16,
    pub cbElements: u32,
    pub cLocks: u32,
    pub pvData: *mut (),
    pub rgsabound: [SAFEARRAYBOUND; 1],
}

pub type LCID = u32;
pub type LPCOLESTR = *const u16;
pub type VARIANT_BOOL = i16;

// VARIANT stuff
// We only need to support a subset of all possible VARIANT types

type VARTYPE = u16;
pub const VT_BSTR: VARTYPE = 8;
pub const VT_BOOL: VARTYPE = 11;
pub const VT_I1: VARTYPE = 16;
pub const VT_I2: VARTYPE = 2;
pub const VT_I4: VARTYPE = 3;
pub const VT_I8: VARTYPE = 20;
pub const VT_UI1: VARTYPE = 17;
pub const VT_UI2: VARTYPE = 18;
pub const VT_UI4: VARTYPE = 19;
pub const VT_UI8: VARTYPE = 21;

pub enum Variant {
    Bstr(BSTR),
    Bool(bool),
    Signed(i64),
    Unsigned(u64),
    Unknown,
}

impl fmt::Debug for Variant {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Bstr(bstr) => core::write!(f, "{bstr}"),
            Self::Bool(bool) => core::write!(f, "{bool}"),
            Self::Signed(i64) => core::write!(f, "[int]{i64}"),
            Self::Unsigned(u64) => core::write!(f, "[uint]{u64}"),
            Self::Unknown => core::write!(f, "<unknown>"),
        }
    }
}

impl fmt::Display for Variant {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Bstr(bstr) => core::write!(f, "{bstr}"),
            Self::Bool(bool) => core::write!(f, "{bool}"),
            Self::Signed(i64) => core::write!(f, "{i64}"),
            Self::Unsigned(u64) => core::write!(f, "{u64}"),
            Self::Unknown => core::write!(f, "<unknown>"),
        }
    }
}

// Windows.Win32.System.Variant.VARIANT
#[repr(C)]
pub struct VARIANT {
    vt: VARTYPE,
    wReserved1: u16,
    wReserved2: u16,
    wReserved3: u16,
    data: VARIANT_DATA,
}
impl VARIANT {
    pub fn into_variant(mut self) -> Variant {
        match self.vt {
            VT_BSTR => Variant::Bstr(unsafe { ManuallyDrop::take(&mut self.data.bstrVal) }),
            VT_BOOL => Variant::Bool(unsafe { self.data.boolVal != 0 }),
            VT_I1 | VT_I2 | VT_I4 | VT_I8 => Variant::Signed(unsafe { self.data.llVal as i64 }),
            VT_UI1 | VT_UI2 | VT_UI4 | VT_UI8 => Variant::Unsigned(unsafe { self.data.llVal }),
            // This should not be reachable when using the API exposed by this crate.
            _ => {
                if cfg!(debug_assertions) {
                    panic!("unhandled variant type: {}", self.vt)
                }
                Variant::Unknown
            }
        }
    }
}
impl Drop for VARIANT {
    fn drop(&mut self) {
        if self.vt == VT_BSTR {
            unsafe {
                ManuallyDrop::drop(&mut self.data.bstrVal);
            }
        }
    }
}

#[repr(C)]
pub union VARIANT_DATA {
    llVal: u64,
    boolVal: VARIANT_BOOL,
    bstrVal: ManuallyDrop<BSTR>,
    // This is necessary to correctly size the union for types we don't support.
    __unknown__: [*mut (); 2],
}

pub const CLSCTX_ALL: u32 = 23;
pub const S_OK: HRESULT = HRESULT(0);
pub const S_FALSE: HRESULT = HRESULT(0x1);
pub const E_POINTER: HRESULT = HRESULT(0x80004003_u32 as i32);
pub const E_INVALIDARG: HRESULT = HRESULT(0x80070057_u32 as i32);
pub const E_UNEXPECTED: HRESULT = HRESULT(0x8000FFFF_u32 as i32);

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    pub fn variant_size_align() {
        #[cfg(target_pointer_width = "64")]
        assert_eq!(size_of::<VARIANT>(), 24);
        #[cfg(target_pointer_width = "32")]
        assert_eq!(size_of::<VARIANT>(), 16);

        assert_eq!(align_of::<VARIANT>(), 8);
    }
}
