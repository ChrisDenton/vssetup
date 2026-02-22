use core::ffi::c_void;
use core::fmt;

use crate::{OkHresult, defs::*};

use windows_result::HRESULT;
use windows_strings::BSTR;

#[repr(transparent)]
#[derive(Debug, PartialEq, Eq, Clone, Copy)]
pub struct InstanceState {
    value: i32,
}

impl InstanceState {
    pub const eNone: Self = Self { value: 0 };
    pub const eLocal: Self = Self { value: 1 };
    pub const eRegistered: Self = Self { value: 2 };
    pub const eNoRebootRequired: Self = Self { value: 4 };
    pub const eNoErrors: Self = Self { value: 8 };
    pub const eComplete: Self = Self {
        value: u32::MAX as i32,
    };
}

impl fmt::Display for InstanceState {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if *self == InstanceState::eNone {
            f.write_str("None")
        } else if *self == InstanceState::eComplete {
            f.write_str("Complete")
        } else {
            // TODO: Do better than a raw value
            f.write_fmt(format_args!("Incomplete({})", self.value))
        }
    }
}

macro_rules! com_interface {
    ($(
        #[interface($iid:literal)]
        pub unsafe interface $interface:ident: $parent:ident {
            $(
                $vis:vis fn $method:ident(&self $(, $arg:ident:$ty:ty)*$(,)?) -> $rtn:ty;
            )*
        }
    )+) => {
        use crate::raw as interface;
        $(
        #[repr(transparent)]
        #[derive(Clone)]
        pub struct $interface($parent);
        #[allow(unused)]
        impl $interface {
            $(
            #[inline(always)]
            pub unsafe fn $method(
                &self,
                $(
                    $arg: $ty,
                )*
            ) -> $rtn {
                unsafe {
                    let (vtable, raw) = Interface::vtable(self);
                    ((**vtable).$method)(raw, $($arg,)*)
                }
            }
            )*
        }

        unsafe impl Interface for $interface {
            const IID: GUID = GUID::from_u128($iid);
            type Vtable = vtable::$interface;
        }
        )*

        mod vtable {
            use super::*;
            use crate::raw as interface;
            type IUnknown = IUnknown_Vtbl;

            $(
                #[repr(C)]
                pub struct $interface {
                    pub base__: $parent,
                    $(
                        pub $method: unsafe extern "system" fn(this: *mut ::core::ffi::c_void, $($arg:$ty,)*) -> $rtn,
                    )*
                }
            )+
        }
    };
}

// Replacement for the windows-rs proc macro
// This is purely for compile-time performance.
com_interface!(
    #[interface(0x_b41463c3_8866_43b5_bc33_2b0676f7f42e)]
    pub unsafe interface ISetupInstance: IUnknown {
        pub fn GetInstanceId(&self, pbstrInstanceId: *mut BSTR) -> HRESULT;
        pub fn GetInstallDate(&self, pInstallDate: *mut FILETIME) -> HRESULT;
        pub fn GetInstallationName(&self, pbstrInstallationName: *mut BSTR) -> HRESULT;
        pub fn GetInstallationPath(&self, pbstrInstallationPath: *mut BSTR) -> HRESULT;
        pub fn GetInstallationVersion(&self, pbstrInstallationVersion: *mut BSTR) -> HRESULT;
        pub fn GetDisplayName(&self, lcid: LCID, pbstrDisplayName: *mut BSTR) -> HRESULT;
        pub fn GetDescription(&self, lcid: LCID, pbstrDescription: *mut BSTR) -> HRESULT;
        pub fn ResolvePath(&self, pwszRelativePath: LPCOLESTR, pbstrAbsolutePath: *mut BSTR)
        -> HRESULT;
    }

    #[interface(0x_89143c9a_05af_49b0_b717_72e218a2185c)]
    pub unsafe interface ISetupInstance2: ISetupInstance {
        pub fn GetState(&self, pState: *mut InstanceState) -> HRESULT;
        pub fn GetPackages(&self, ppsaPackages: *mut *mut SAFEARRAY) -> HRESULT;
        pub fn GetProduct(&self, ppPackage: *mut Option<interface::ISetupPackageReference>) -> HRESULT;
        pub fn GetProductPath(&self, pbstrProductPath: *mut BSTR) -> HRESULT;
        pub fn GetErrors(&self, ppErrorState: *mut Option<interface::ISetupErrorState>) -> HRESULT;
        pub fn IsLaunchable(&self, pfIsLaunchable: *mut VARIANT_BOOL) -> HRESULT;
        pub fn IsComplete(&self, pfIsComplete: *mut VARIANT_BOOL) -> HRESULT;
        pub fn GetProperties(&self, ppProperties: *mut Option<interface::ISetupPropertyStore>) -> HRESULT;
        pub fn GetEnginePath(&self, pbstrEnginePath: *mut BSTR) -> HRESULT;
    }

    #[interface(0xda8d8a16_b2b6_4487_a2f1_594ccccd6bf5)]
    pub unsafe interface ISetupPackageReference: IUnknown {
        pub fn GetId(&self, pbstrId: *mut BSTR) -> HRESULT;
        pub fn GetVersion(&self, pbstrVersion: *mut BSTR) -> HRESULT;
        pub fn GetChip(&self, pbstrChip: *mut BSTR) -> HRESULT;
        pub fn GetLanguage(&self, pbstrLanguage: *mut BSTR) -> HRESULT;
        pub fn GetBranch(&self, pbstrBranch: *mut BSTR) -> HRESULT;
        pub fn GetType(&self, pbstrType: *mut BSTR) -> HRESULT;
        pub fn GetUniqueId(&self, pbstrUniqueId: *mut BSTR) -> HRESULT;
        pub fn GetIsExtension(&self, pfIsExtension: *mut VARIANT_BOOL) -> HRESULT;
    }

    #[interface(0x_46dccd94_a287_476a_851e_dfbc2ffdbc20)]
    pub unsafe interface ISetupErrorState: IUnknown {
        pub fn GetFailedPackages(&self, ppsaFailedPackages: *mut *mut SAFEARRAY) -> HRESULT;
        pub fn GetSkippedPackages(&self, ppsaSkippedPackages: *mut *mut SAFEARRAY) -> HRESULT;
    }

    #[interface(0x_9871385b_ca69_48f2_bc1f_7a37cbf0b1ef)]
    pub unsafe interface ISetupErrorState2: ISetupErrorState {
        pub fn GetErrorLogFilePath(&self, pbstrErrorLogFilePath: *mut BSTR) -> HRESULT;
        pub fn GetLogFilePath(&self, pbstrLogFilePath: *mut BSTR) -> HRESULT;
    }

    #[interface(0x290019ad_28e2_46d5_9de5_da4b6bcf8057)]
    pub unsafe interface ISetupErrorState3: ISetupErrorState2 {
        pub fn GetRuntimeError(&self, ppErrorInfo: *mut Option<interface::ISetupErrorInfo>) -> HRESULT;
    }

    #[interface(0x_e73559cd_7003_4022_b134_27dc650b280f)]
    pub unsafe interface ISetupFailedPackageReference: ISetupPackageReference {}

    #[interface(0x0fad873e_e874_42e3_b268_4fe2f096b9ca)]
    pub unsafe interface ISetupFailedPackageReference2: ISetupFailedPackageReference {
        pub fn GetLogFilePath(&self, pbstrLogFilePath: *mut BSTR) -> HRESULT;
        pub fn GetDescription(&self, pbstrDescription: *mut BSTR) -> HRESULT;
        pub fn GetSignature(&self, pbstrSignature: *mut BSTR) -> HRESULT;
        pub fn GetDetails(&self, ppsaDetails: *mut *mut SAFEARRAY) -> HRESULT;
        pub fn GetAffectedPackages(&self, ppsaAffectedPackages: *mut *mut SAFEARRAY) -> HRESULT;
    }

    #[interface(0x_ebc3ae68_ad15_44e8_8377_39dbf0316f6c)]
    pub unsafe interface ISetupFailedPackageReference3: ISetupFailedPackageReference2 {
        pub fn GetAction(&self, pbstrAction: *mut BSTR) -> HRESULT;
        pub fn GetReturnCode(&self, pbstrReturnCode: *mut BSTR) -> HRESULT;
    }

    #[interface(0x_a170b5ef_223d_492b_b2d4_945032980685)]
    pub unsafe interface ISetupProductReference: ISetupPackageReference {
        pub fn GetIsInstalled(&self, pfIsInstalled: *mut VARIANT_BOOL) -> HRESULT;
    }

    #[interface(0x_279a5db3_7503_444b_b34d_308f961b9a06)]
    pub unsafe interface ISetupProductReference2: ISetupProductReference {
        pub fn GetSupportsExtensions(&self, pfSupportsExtensions: *mut VARIANT_BOOL) -> HRESULT;
    }

    #[interface(0x_6380bcff_41d3_4b2e_8b2e_bf8a6810c848)]
    pub unsafe interface IEnumSetupInstances: IUnknown {
        pub fn Next(
            &self,
            celt: u32,
            rgelt: *mut Option<interface::ISetupInstance>,
            pceltFetched: *mut u32,
        ) -> HRESULT;
        pub fn Skip(&self, celt: u32) -> HRESULT;
        pub fn Reset(&self) -> HRESULT;
        pub fn Clone(&self, ppenum: *mut Option<interface::IEnumSetupInstances>) -> HRESULT;
    }

    #[interface(0x_c601c175_a3be_44bc_91f6_4568d230fc83)]
    pub unsafe interface ISetupPropertyStore: IUnknown {
        pub fn GetNames(&self, ppsaNames: *mut *mut SAFEARRAY) -> HRESULT;
        pub fn GetValue(&self, pwszName: LPCOLESTR, pvtValue: *mut VARIANT) -> HRESULT;
    }

    #[interface(0x_9ad8e40f_39a2_40f1_bf64_0a6c50dd9eeb)]
    pub unsafe interface ISetupInstanceCatalog: IUnknown {
        pub fn GetCatalogInfo(&self, ppCatalogInfo: *mut Option<interface::ISetupPropertyStore>) -> HRESULT;
        pub fn IsPrerelease(&self, pfIsPrerelease: *mut VARIANT_BOOL) -> HRESULT;
    }

    #[interface(0x_f4bd7382_fe27_4ab4_b974_9905b2a148b0)]
    pub unsafe interface ISetupLocalizedProperties: IUnknown {
        pub fn GetLocalizedProperties(
            &self,
            ppLocalizedProperties: *mut Option<interface::ISetupLocalizedPropertyStore>,
        ) -> HRESULT;
        pub fn GetLocalizedChannelProperties(
            &self,
            ppLocalizedChannelProperties: *mut Option<interface::ISetupLocalizedPropertyStore>,
        ) -> HRESULT;
    }

    #[interface(0x5bb53126_e0d5_43df_80f1_6b161e5c6f6c)]
    pub unsafe interface ISetupLocalizedPropertyStore: IUnknown {
        pub fn GetNames(&self, lcid: LCID, ppsaNames: *mut *mut SAFEARRAY) -> HRESULT;
        pub fn GetValue(&self, pwszName: LPCOLESTR, lcid: LCID, pvtValue: *mut VARIANT) -> HRESULT;
    }

    #[interface(0x_42843719_db4c_46c2_8e7c_64f1816efd5b)]
    pub unsafe interface ISetupConfiguration: IUnknown {
        pub fn EnumInstances(&self, ppEnumInstances: *mut Option<interface::IEnumSetupInstances>) -> HRESULT;
        pub fn GetInstanceForCurrentProcess(&self, ppInstance: *mut Option<interface::ISetupInstance>) -> HRESULT;
        pub fn GetInstanceForPath(
            &self,
            wzPath: *const u16,
            ppInstance: *mut Option<interface::ISetupInstance>,
        ) -> HRESULT;
    }

    #[interface(0x_26aab78c_4a60_49d6_af3b_3c35bc93365d)]
    pub unsafe interface ISetupConfiguration2: ISetupConfiguration {
        pub fn EnumAllInstances(&self, ppEnumInstances: *mut Option<interface::IEnumSetupInstances>) -> HRESULT;
    }

    #[interface(0x_e1da4cbd_64c4_4c44_821d_98fab64c4da7)]
    pub unsafe interface ISetupPolicy: IUnknown {
        pub fn GetSharedInstallationPath(&self, pbstrSharedInstallationPath: *mut BSTR) -> HRESULT;
        pub fn GetValue(&self, pwszName: LPCOLESTR, pvtValue: *mut VARIANT) -> HRESULT;
    }

    #[interface(0x_2a2f3292_958e_4905_b36e_013be84e27ab)]
    pub unsafe interface ISetupErrorInfo: IUnknown {
        pub fn GetErrorHResult(&self, plHResult: *mut HRESULT) -> HRESULT;
        pub fn GetErrorClassName(&self, pbstrClassName: *mut BSTR) -> HRESULT;
        pub fn GetErrorMessage(&self, pbstrMessage: *mut BSTR) -> HRESULT;
    }

    #[interface(0x_42b21b78_6192_463e_87bf_d577838f1d5c)]
    pub unsafe interface ISetupHelper: IUnknown {
        pub fn ParseVersion(&self, pwszVersion: LPCOLESTR, pullVersion: *mut u64) -> HRESULT;
        pub fn ParseVersionRange(
            &self,
            pwszVersionRange: LPCOLESTR,
            pullMinVersion: *mut u64,
            pullMaxVersion: *mut u64,
        ) -> HRESULT;
    }
);

pub const SETUP_CONFIGURATION_CLSID: GUID = GUID::from_u128(0x177F0C4A_1CD3_4DE7_A32C_71DBBB9FA36D);

pub(crate) unsafe trait Interface: Sized {
    const IID: GUID;
    type Vtable;

    #[inline(always)]
    unsafe fn vtable(&self) -> (*const *mut Self::Vtable, *mut c_void) {
        unsafe {
            let raw = *(core::ptr::from_ref(self).cast::<*mut c_void>());
            let vtable = raw.cast::<*mut Self::Vtable>();
            (vtable, raw)
        }
    }

    #[inline(always)]
    fn cast<I: Interface>(&self) -> Result<I, HRESULT> {
        unsafe {
            let (vtable, raw) = self.vtable();
            let vtable = vtable.cast::<*mut IUnknown_Vtbl>();
            let mut interface = None;
            ((**vtable).QueryInterface)(raw, &I::IID, core::ptr::from_mut(&mut interface).cast())
                .ok_hresult()?;
            interface.ok_or(E_POINTER)
        }
    }

    unsafe fn from_raw(raw: *mut c_void) -> Self {
        unsafe { core::mem::transmute_copy(&raw) }
    }
}
