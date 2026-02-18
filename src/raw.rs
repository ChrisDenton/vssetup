use crate::defs::*;

use windows_core::{interface, IUnknown, IUnknown_Vtbl, GUID};
use windows_result::HRESULT;
use windows_strings::BSTR;

#[repr(transparent)]
#[derive(Debug, PartialEq, Eq)]
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

#[interface("b41463c3-8866-43b5-bc33-2b0676f7f42e")]
pub unsafe trait ISetupInstance: IUnknown {
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

#[interface("89143c9a-05af-49b0-b717-72e218a2185c")]
pub unsafe trait ISetupInstance2: ISetupInstance {
    pub fn GetState(&self, pState: *mut InstanceState) -> HRESULT;
    pub fn GetPackages(&self, ppsaPackages: *mut *mut SAFEARRAY) -> HRESULT;
    pub fn GetProduct(&self, ppPackage: *mut Option<ISetupPackageReference>) -> HRESULT;
    pub fn GetProductPath(&self, pbstrProductPath: *mut BSTR) -> HRESULT;
    pub fn GetErrors(&self, ppErrorState: *mut Option<ISetupErrorState>) -> HRESULT;
    pub fn IsLaunchable(&self, pfIsLaunchable: *mut VARIANT_BOOL) -> HRESULT;
    pub fn IsComplete(&self, pfIsComplete: *mut VARIANT_BOOL) -> HRESULT;
    pub fn GetProperties(&self, ppProperties: *mut Option<ISetupPropertyStore>) -> HRESULT;
    pub fn GetEnginePath(&self, pbstrEnginePath: *mut BSTR) -> HRESULT;
}

#[interface("da8d8a16-b2b6-4487-a2f1-594ccccd6bf5")]
pub unsafe trait ISetupPackageReference: IUnknown {
    pub fn GetId(&self, pbstrId: *mut BSTR) -> HRESULT;
    pub fn GetVersion(&self, pbstrVersion: *mut BSTR) -> HRESULT;
    pub fn GetChip(&self, pbstrChip: *mut BSTR) -> HRESULT;
    pub fn GetLanguage(&self, pbstrLanguage: *mut BSTR) -> HRESULT;
    pub fn GetBranch(&self, pbstrBranch: *mut BSTR) -> HRESULT;
    pub fn GetType(&self, pbstrType: *mut BSTR) -> HRESULT;
    pub fn GetUniqueId(&self, pbstrUniqueId: *mut BSTR) -> HRESULT;
    pub fn GetIsExtension(&self, pfIsExtension: *mut VARIANT_BOOL) -> HRESULT;
}

#[interface("46dccd94-a287-476a-851e-dfbc2ffdbc20")]
pub unsafe trait ISetupErrorState: IUnknown {
    pub fn GetFailedPackages(&self, ppsaFailedPackages: *mut *mut SAFEARRAY) -> HRESULT;
    pub fn GetSkippedPackages(&self, ppsaSkippedPackages: *mut *mut SAFEARRAY) -> HRESULT;
}

#[interface("9871385b-ca69-48f2-bc1f-7a37cbf0b1ef")]
pub unsafe trait ISetupErrorState2: ISetupErrorState {
    pub fn GetErrorLogFilePath(&self, pbstrErrorLogFilePath: *mut BSTR) -> HRESULT;
    pub fn GetLogFilePath(&self, pbstrLogFilePath: *mut BSTR) -> HRESULT;
}

#[interface("290019ad-28e2-46d5-9de5-da4b6bcf8057")]
pub unsafe trait ISetupErrorState3: ISetupErrorState2 {
    pub fn GetRuntimeError(&self, ppErrorInfo: *mut Option<ISetupErrorInfo>) -> HRESULT;
}

#[interface("e73559cd-7003-4022-b134-27dc650b280f")]
pub unsafe trait ISetupFailedPackageReference: ISetupPackageReference {}

#[interface("0fad873e-e874-42e3-b268-4fe2f096b9ca")]
pub unsafe trait ISetupFailedPackageReference2: ISetupFailedPackageReference {
    pub fn GetLogFilePath(&self, pbstrLogFilePath: *mut BSTR) -> HRESULT;
    pub fn GetDescription(&self, pbstrDescription: *mut BSTR) -> HRESULT;
    pub fn GetSignature(&self, pbstrSignature: *mut BSTR) -> HRESULT;
    pub fn GetDetails(&self, ppsaDetails: *mut *mut SAFEARRAY) -> HRESULT;
    pub fn GetAffectedPackages(&self, ppsaAffectedPackages: *mut *mut SAFEARRAY) -> HRESULT;
}

#[interface("ebc3ae68-ad15-44e8-8377-39dbf0316f6c")]
pub unsafe trait ISetupFailedPackageReference3: ISetupFailedPackageReference2 {
    pub fn GetAction(&self, pbstrAction: *mut BSTR) -> HRESULT;
    pub fn GetReturnCode(&self, pbstrReturnCode: *mut BSTR) -> HRESULT;
}

#[interface("a170b5ef-223d-492b-b2d4-945032980685")]
pub unsafe trait ISetupProductReference: ISetupPackageReference {
    pub fn GetIsInstalled(&self, pfIsInstalled: *mut VARIANT_BOOL) -> HRESULT;
}

#[interface("279a5db3-7503-444b-b34d-308f961b9a06")]
pub unsafe trait ISetupProductReference2: ISetupProductReference {
    pub fn GetSupportsExtensions(&self, pfSupportsExtensions: *mut VARIANT_BOOL) -> HRESULT;
}

#[interface("6380bcff-41d3-4b2e-8b2e-bf8a6810c848")]
pub unsafe trait IEnumSetupInstances: IUnknown {
    pub fn Next(
        &self,
        celt: u32,
        rgelt: *mut Option<ISetupInstance>,
        pceltFetched: *mut u32,
    ) -> HRESULT;
    pub fn Skip(&self, celt: u32) -> HRESULT;
    pub fn Reset(&self) -> HRESULT;
    pub fn Clone(&self, ppenum: *mut Option<IEnumSetupInstances>) -> HRESULT;
}

#[interface("c601c175-a3be-44bc-91f6-4568d230fc83")]
pub unsafe trait ISetupPropertyStore: IUnknown {
    pub fn GetNames(&self, ppsaNames: *mut *mut SAFEARRAY) -> HRESULT;
    pub fn GetValue(&self, pwszName: LPCOLESTR, pvtValue: *mut VARIANT) -> HRESULT;
}

#[interface("9ad8e40f-39a2-40f1-bf64-0a6c50dd9eeb")]
pub unsafe trait ISetupInstanceCatalog: IUnknown {
    pub fn GetCatalogInfo(&self, ppCatalogInfo: *mut Option<ISetupPropertyStore>) -> HRESULT;
    pub fn IsPrerelease(&self, pfIsPrerelease: *mut VARIANT_BOOL) -> HRESULT;
}

#[interface("f4bd7382-fe27-4ab4-b974-9905b2a148b0")]
pub unsafe trait ISetupLocalizedProperties: IUnknown {
    pub fn GetLocalizedProperties(
        &self,
        ppLocalizedProperties: *mut Option<ISetupLocalizedPropertyStore>,
    ) -> HRESULT;
    pub fn GetLocalizedChannelProperties(
        &self,
        ppLocalizedChannelProperties: *mut Option<ISetupLocalizedPropertyStore>,
    ) -> HRESULT;
}

#[interface("5bb53126-e0d5-43df-80f1-6b161e5c6f6c")]
pub unsafe trait ISetupLocalizedPropertyStore: IUnknown {
    pub fn GetNames(&self, lcid: LCID, ppsaNames: *mut *mut SAFEARRAY) -> HRESULT;
    pub fn GetValue(&self, pwszName: LPCOLESTR, lcid: LCID, pvtValue: *mut VARIANT) -> HRESULT;
}

#[interface("42843719-db4c-46c2-8e7c-64f1816efd5b")]
pub unsafe trait ISetupConfiguration: IUnknown {
    pub fn EnumInstances(&self, ppEnumInstances: *mut Option<IEnumSetupInstances>) -> HRESULT;
    pub fn GetInstanceForCurrentProcess(&self, ppInstance: *mut Option<ISetupInstance>) -> HRESULT;
    pub fn GetInstanceForPath(
        &self,
        wzPath: *const u16,
        ppInstance: *mut Option<ISetupInstance>,
    ) -> HRESULT;
}

#[interface("26aab78c-4a60-49d6-af3b-3c35bc93365d")]
pub unsafe trait ISetupConfiguration2: ISetupConfiguration {
    pub fn EnumAllInstances(&self, ppEnumInstances: *mut Option<IEnumSetupInstances>) -> HRESULT;
}

#[interface("e1da4cbd-64c4-4c44-821d-98fab64c4da7")]
pub unsafe trait ISetupPolicy: IUnknown {
    pub fn GetSharedInstallationPath(&self, pbstrSharedInstallationPath: *mut BSTR) -> HRESULT;
    pub fn GetValue(&self, pwszName: LPCOLESTR, pvtValue: *mut VARIANT) -> HRESULT;
}

#[interface("2a2f3292-958e-4905-b36e-013be84e27ab")]
pub unsafe trait ISetupErrorInfo: IUnknown {
    pub fn GetErrorHResult(&self, plHResult: *mut HRESULT) -> HRESULT;
    pub fn GetErrorClassName(&self, pbstrClassName: *mut BSTR) -> HRESULT;
    pub fn GetErrorMessage(&self, pbstrMessage: *mut BSTR) -> HRESULT;
}

#[interface("42b21b78-6192-463e-87bf-d577838f1d5c")]
pub unsafe trait ISetupHelper: IUnknown {
    pub fn ParseVersion(&self, pwszVersion: LPCOLESTR, pullVersion: *mut u64) -> HRESULT;
    pub fn ParseVersionRange(
        &self,
        pwszVersionRange: LPCOLESTR,
        pullMinVersion: *mut u64,
        pullMaxVersion: *mut u64,
    ) -> HRESULT;
}

pub const SETUP_CONFIGURATION_CLSID: GUID = GUID::from_u128(0x177F0C4A_1CD3_4DE7_A32C_71DBBB9FA36D);
