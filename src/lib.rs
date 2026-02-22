//! Query Visual Studio setup for information on installed instances of Visual Studio.
//!
//! This is a thin wrapper around the COM interface.
//! Consult the [`Microsoft.VisualStudio.Setup.Configuration`] documentation for more information on the API.
//!
//! To use this library you must first initialize COM.
//! The helper, [`com::initialize`] will do this for you.
//!
//! ## Example
//!
//! ```rust
//! use vssetup::{com, HRESULT, SetupConfiguration};
//!
//! fn main() -> Result<(), HRESULT> {
//!     com::initialize();
//!     let setup = SetupConfiguration::new()?;
//!     let instances = setup.EnumAllInstances()?;
//!     for instance in instances {
//!         let name = instance.GetDisplayName(0x400)?.to_string();
//!         println!("{name}");
//!     }
//!     Ok(())
//! }
//! ```
//!
//! [`com::initialize`]: [crate::com::initialize]
//! [`Microsoft.VisualStudio.Setup.Configuration`]: https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.setup.configuration

// We should use the same style as the official documentation
#![allow(nonstandard_style)]
#![allow(clippy::upper_case_acronyms)]

mod defs;
use defs::*;

mod raw;
use raw::*;

pub mod com;

pub use windows_result::HRESULT;
pub use windows_strings::{BSTR, PCWSTR};

use core::marker::PhantomData;
use core::ops::Deref;
use core::ptr::NonNull;
use core::ptr::null_mut as null;
use raw::Interface;

#[doc(hidden)]
pub use windows_strings::w;

#[macro_export]
macro_rules! wide_str {
    ($str:literal) => {
        unsafe { $crate::WideStr::from_ptr($crate::w!($str).as_ptr()).unwrap() }
    };
}

#[derive(Clone, Copy, Eq)]
pub struct WideStr<'a> {
    wide: NonNull<u16>,
    lifetime: PhantomData<&'a [u16]>,
}

impl<'a> WideStr<'a> {
    pub fn from_slice_with_nul(u16s: &[u16]) -> Result<Self, HRESULT> {
        let pos = u16s.iter().copied().position(|n| n == 0);
        if pos == Some(u16s.len() - 1) {
            // SAFETY: We've checked there is a null.
            Ok(unsafe { Self::from_slice_with_nul_unchecked(u16s) })
        } else {
            Err(E_INVALIDARG)
        }
    }

    pub fn from_slice_until_nul(u16s: &[u16]) -> Result<Self, HRESULT> {
        if u16s.last() == Some(&0) {
            // SAFETY: We've checked there is a null.
            Ok(unsafe { Self::from_slice_with_nul_unchecked(u16s) })
        } else {
            Err(E_INVALIDARG)
        }
    }

    /// Create a `WideStr` without doing any runtime checks.
    /// The `WideStr` will be truncated to the first null.
    ///
    /// # Safety
    ///
    /// The array must contain at least one null.
    pub const unsafe fn from_slice_with_nul_unchecked(u16s: &[u16]) -> Self {
        // SAFETY: It's up to the caller to ensure this is safe.
        Self {
            wide: unsafe { NonNull::new_unchecked(u16s.as_ptr().cast_mut()) },
            lifetime: PhantomData,
        }
    }

    /// Create a `WideStr` from a pointer without checking for null termination.
    ///
    /// If the pointer is null then `None` will be returned.
    ///
    /// # Safety
    ///
    /// The pointer must either be null or a pointer to a null-terminated string.
    pub const unsafe fn from_ptr(ptr: *const u16) -> Option<Self> {
        if let Some(ptr) = NonNull::new(ptr.cast_mut()) {
            Some(Self {
                wide: ptr,
                lifetime: PhantomData,
            })
        } else {
            None
        }
    }

    pub fn count_units(self) -> usize {
        // SAFETY: This type is guaranteed non-null and null-terminated.
        unsafe { PCWSTR(self.as_ptr()).len() }
    }

    pub fn to_slice(self) -> &'a [u16] {
        unsafe { core::slice::from_raw_parts(self.as_ptr(), self.count_units()) }
    }

    pub const fn as_ptr(self) -> *const u16 {
        self.wide.as_ptr()
    }
}

impl TryFrom<&[u16]> for WideStr<'_> {
    type Error = HRESULT;
    fn try_from(value: &[u16]) -> Result<Self, Self::Error> {
        Self::from_slice_with_nul(value)
    }
}

impl From<&BSTR> for WideStr<'_> {
    fn from(value: &BSTR) -> Self {
        // SAFETY: A BSTR is always null-terminated.
        unsafe { WideStr::from_slice_with_nul_unchecked(value) }
    }
}

impl PartialEq for WideStr<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.to_slice() == other.to_slice()
    }
}

impl PartialEq<BSTR> for WideStr<'_> {
    fn eq(&self, other: &BSTR) -> bool {
        self.to_slice() == other.deref()
    }
}

/// The entry point for these APIs.
///
/// # Example
///
/// ```rust
/// # fn main() -> Result<(), vssetup::HRESULT> {
/// # vssetup::com::initialize();
/// let setup = vssetup::SetupConfiguration::new()?;
/// # Ok(()) }
/// ```
pub struct SetupConfiguration {
    raw: ISetupConfiguration,
}

impl SetupConfiguration {
    /// Create a new instance of `SetupConfiguration`.
    ///
    /// This will fail if COM is not already initalized.
    pub fn new() -> Result<Self, HRESULT> {
        unsafe {
            let mut interface = null();
            CoCreateInstance(
                &SETUP_CONFIGURATION_CLSID,
                null(),
                CLSCTX_ALL,
                &ISetupConfiguration::IID,
                &mut interface,
            )
            .ok_hresult()?;
            let interface = NonNull::new(interface).assert_ok()?;
            Ok(Self::from_raw(interface))
        }
    }

    pub fn EnumInstances(&self) -> Result<EnumSetupInstances, HRESULT> {
        unsafe {
            let mut instances = None;
            self.com_ptr().EnumInstances(&mut instances).ok_hresult()?;
            let instances = instances.assert_ok()?;
            Ok(EnumSetupInstances::from_raw(instances))
        }
    }

    pub fn EnumAllInstances(&self) -> Result<EnumSetupInstances, HRESULT> {
        unsafe {
            let mut instances = None;
            let setup = self.com_ptr().cast::<ISetupConfiguration2>()?;
            setup.EnumAllInstances(&mut instances).ok_hresult()?;
            let instances = instances.assert_ok()?;
            Ok(EnumSetupInstances::from_raw(instances))
        }
    }

    pub fn GetInstanceForCurrentProcess(&self) -> Result<SetupInstance, HRESULT> {
        unsafe {
            let mut instance = None;
            self.com_ptr()
                .GetInstanceForCurrentProcess(&mut instance)
                .ok_hresult()?;
            let instance = instance.assert_ok()?;
            Ok(SetupInstance::from_raw(instance))
        }
    }

    pub fn GetInstanceForPath<'w, W: TryInto<WideStr<'w>>>(
        &self,
        path: W,
    ) -> Result<SetupInstance, HRESULT> {
        let Ok(path) = path.try_into() else {
            return Err(E_INVALIDARG);
        };
        unsafe {
            let mut instance = None;
            self.com_ptr()
                .GetInstanceForPath(path.as_ptr(), &mut instance)
                .ok_hresult()?;
            let instance = instance.assert_ok()?;
            Ok(SetupInstance::from_raw(instance))
        }
    }

    /// # Safety
    ///
    /// The pointer must be a valid ISetupConfiguration COM pointer.
    unsafe fn from_raw(raw: NonNull<core::ffi::c_void>) -> Self {
        Self {
            // SAFETY: the caller must make sure this is safe.
            raw: unsafe { ISetupConfiguration::from_raw(raw.as_ptr()) },
        }
    }

    fn com_ptr(&self) -> &ISetupConfiguration {
        &self.raw
    }
}

pub struct EnumSetupInstances {
    pub raw: IEnumSetupInstances,
}

impl EnumSetupInstances {
    /// Fill the buffer with the next set of instances.
    ///
    /// # Errors
    ///
    /// Can fail with `E_OUTOFMEMORY` if a `SetupInstance` couldn't be allocated.
    pub fn Next(
        &self,
        instances: &mut [Option<SetupInstance>],
    ) -> Result<Option<&[SetupInstance]>, HRESULT> {
        unsafe {
            let len: u32 = instances.len().try_into().unwrap_or(u32::MAX);
            let mut fetched = 0;
            let hresult = self
                .com_ptr()
                .Next(len, instances.as_mut_ptr().cast(), &mut fetched);
            if hresult == S_FALSE {
                Ok(None)
            } else if hresult.is_err() {
                Err(hresult)
            } else if fetched <= len {
                Ok(Some(core::slice::from_raw_parts(
                    instances.as_ptr().cast(),
                    fetched as usize,
                )))
            } else {
                // If this happens then something has gone very wrong with the other side of the API.
                Err(E_UNEXPECTED)
            }
        }
    }

    pub fn Skip(&self, count: u32) -> Result<bool, HRESULT> {
        let hresult = unsafe { self.com_ptr().Skip(count) };
        if hresult == S_FALSE {
            Ok(false)
        } else {
            hresult.ok_hresult().map(|_| true)
        }
    }

    pub fn Reset(&self) {
        unsafe {
            // Thie will always return S_OK
            let result = self.com_ptr().Reset();
            debug_assert_eq!(result, S_OK);
        }
    }

    pub fn Clone(&self) -> Result<EnumSetupInstances, HRESULT> {
        unsafe {
            let mut new = None;
            self.com_ptr().Clone(&mut new).ok_hresult()?;
            let new = new.assert_ok()?;
            Ok(EnumSetupInstances::from_raw(new))
        }
    }

    fn com_ptr(&self) -> &IEnumSetupInstances {
        &self.raw
    }

    unsafe fn from_raw(raw: IEnumSetupInstances) -> EnumSetupInstances {
        EnumSetupInstances { raw }
    }
}

impl Iterator for EnumSetupInstances {
    type Item = SetupInstance;

    /// Convinence method for calling [`Next`](Self::Next) in a loop.
    ///
    /// If `Next` returns an error, this will return `None` instead.
    fn next(&mut self) -> Option<Self::Item> {
        let mut instance = None;
        unsafe {
            let hresult = self.com_ptr().Next(1, &mut instance, null());
            match hresult {
                S_OK => instance.map(|raw| SetupInstance::from_raw(raw)),
                _ => None,
            }
        }
    }
}

pub struct SetupInstance {
    raw: ISetupInstance,
}

impl SetupInstance {
    pub fn GetInstanceId(&self) -> Result<BSTR, HRESULT> {
        let mut id = BSTR::new();
        unsafe {
            self.com_ptr()
                .GetInstanceId(&mut id)
                .ok_hresult()
                .map(|_| id)
        }
    }

    pub fn GetInstallDate(&self) -> Result<FILETIME, HRESULT> {
        unsafe {
            let mut time = FILETIME::default();
            self.com_ptr()
                .GetInstallDate(&mut time)
                .ok_hresult()
                .map(|_| time)
        }
    }

    pub fn GetInstallationName(&self) -> Result<BSTR, HRESULT> {
        let mut name = BSTR::new();
        unsafe {
            self.com_ptr()
                .GetInstallationName(&mut name)
                .ok_hresult()
                .map(|_| name)
        }
    }

    pub fn GetInstallationPath(&self) -> Result<BSTR, HRESULT> {
        let mut path = BSTR::new();
        unsafe {
            self.com_ptr()
                .GetInstallationPath(&mut path)
                .ok_hresult()
                .map(|_| path)
        }
    }

    pub fn GetInstallationVersion(&self) -> Result<BSTR, HRESULT> {
        let mut version = BSTR::new();
        unsafe {
            self.com_ptr()
                .GetInstallationVersion(&mut version)
                .ok_hresult()
                .map(|_| version)
        }
    }

    pub fn GetDisplayName(&self, lcid: u32) -> Result<BSTR, HRESULT> {
        let mut name = BSTR::new();
        unsafe {
            self.raw
                .GetDisplayName(lcid, &mut name)
                .ok_hresult()
                .map(|_| name)
        }
    }

    pub fn GetDescription(&self, lcid: LCID) -> Result<BSTR, HRESULT> {
        let mut description = BSTR::new();
        unsafe {
            self.raw
                .GetDescription(lcid, &mut description)
                .ok_hresult()
                .map(|_| description)
        }
    }

    pub fn ResolvePath<'w, W: TryInto<WideStr<'w>>>(
        &self,
        relative_path: W,
    ) -> Result<BSTR, HRESULT> {
        let Ok(relative_path) = relative_path.try_into() else {
            return Err(E_INVALIDARG);
        };
        unsafe {
            let mut absolute_path = BSTR::new();
            self.com_ptr()
                .ResolvePath(relative_path.as_ptr(), &mut absolute_path)
                .ok_hresult()
                .map(|_| absolute_path)
        }
    }

    pub fn GetProductPath(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let instance: ISetupInstance2 = self.com_ptr().cast()?;
            let mut path = BSTR::new();
            instance
                .GetProductPath(&mut path)
                .ok_hresult()
                .map(|_| path)
        }
    }

    pub fn GetEnginePath(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let instance: ISetupInstance2 = self.com_ptr().cast()?;
            let mut path = BSTR::new();
            instance.GetEnginePath(&mut path).ok_hresult().map(|_| path)
        }
    }

    pub fn IsLaunchable(&self) -> Result<bool, HRESULT> {
        unsafe {
            let instance: ISetupInstance2 = self.com_ptr().cast()?;
            let mut bool = 0;
            instance
                .IsLaunchable(&mut bool)
                .ok_hresult()
                .map(|_| bool != 0)
        }
    }

    pub fn IsComplete(&self) -> Result<bool, HRESULT> {
        unsafe {
            let instance: ISetupInstance2 = self.com_ptr().cast()?;
            let mut bool = 0;
            instance
                .IsComplete(&mut bool)
                .ok_hresult()
                .map(|_| bool != 0)
        }
    }

    pub fn GetProduct(&self) -> Result<Option<SetupProductReference>, HRESULT> {
        unsafe {
            let instance: ISetupInstance2 = self.com_ptr().cast()?;
            let mut product = None;
            instance.GetProduct(&mut product).ok_hresult()?;
            Ok(product.map(|raw| SetupProductReference::from_raw(raw)))
        }
    }

    pub fn GetState(&self) -> Result<InstanceState, HRESULT> {
        unsafe {
            let instance: ISetupInstance2 = self.com_ptr().cast()?;
            let mut state = InstanceState::eNone;
            instance.GetState(&mut state).ok_hresult().map(|_| state)
        }
    }

    pub fn GetPackages(&self) -> Result<SafeArray<SetupPackageReference>, HRESULT> {
        unsafe {
            let instance: ISetupInstance2 = self.com_ptr().cast()?;
            let mut packages = core::ptr::null_mut();
            instance.GetPackages(&mut packages).ok_hresult()?;
            if packages.is_null() {
                debug_assert!(!packages.is_null());
                Err(E_POINTER)
            } else {
                SafeArray::from_raw(packages.cast())
            }
        }
    }

    pub fn GetProperties(&self) -> Result<Option<SetupPropertyStore>, HRESULT> {
        unsafe {
            let instance: ISetupInstance2 = self.com_ptr().cast()?;
            let mut properties = None;
            instance.GetProperties(&mut properties).ok_hresult()?;
            Ok(properties.map(|raw| SetupPropertyStore::from_raw(raw)))
        }
    }

    pub fn GetErrors(&self) -> Result<Option<SetupErrorState>, HRESULT> {
        unsafe {
            let instance: ISetupInstance2 = self.com_ptr().cast()?;
            let mut errors = None;
            instance.GetErrors(&mut errors).ok_hresult()?;
            Ok(errors.map(|raw| SetupErrorState::from_raw(raw)))
        }
    }

    pub fn to_catalog(&self) -> Result<SetupInstanceCatalog, HRESULT> {
        unsafe {
            self.com_ptr()
                .cast()
                .map(|raw| SetupInstanceCatalog::from_raw(raw))
                .map_err(Into::into)
        }
    }

    pub fn to_property_store(&self) -> Result<SetupPropertyStore, HRESULT> {
        unsafe {
            self.com_ptr()
                .cast()
                .map(|raw| SetupPropertyStore::from_raw(raw))
                .map_err(Into::into)
        }
    }

    fn com_ptr(&self) -> &ISetupInstance {
        &self.raw
    }

    unsafe fn from_raw(raw: ISetupInstance) -> SetupInstance {
        SetupInstance { raw }
    }
}

pub struct SetupProductReference {
    // This is not a typo. `GetProduct` returns a package reference for some reason.
    raw: ISetupPackageReference,
}

impl SetupProductReference {
    pub fn GetIsInstalled(&self) -> Result<bool, HRESULT> {
        unsafe {
            let product: ISetupProductReference = self.com_ptr().cast()?;
            let mut is_installed = 0;
            product
                .GetIsInstalled(&mut is_installed)
                .ok_hresult()
                .map(|_| is_installed != 0)
        }
    }

    pub fn GetSupportsExtensions(&self) -> Result<bool, HRESULT> {
        unsafe {
            let product: ISetupProductReference2 = self.com_ptr().cast()?;
            let mut supports_extensions = 0;
            product
                .GetSupportsExtensions(&mut supports_extensions)
                .ok_hresult()
                .map(|_| supports_extensions != 0)
        }
    }

    fn com_ptr(&self) -> &ISetupPackageReference {
        &self.raw
    }

    unsafe fn from_raw(raw: ISetupPackageReference) -> SetupProductReference {
        SetupProductReference { raw }
    }
}

impl Deref for SetupProductReference {
    type Target = SetupPackageReference;
    fn deref(&self) -> &Self::Target {
        unsafe { &*(self as *const Self as *const SetupPackageReference) }
    }
}

pub struct SetupErrorState {
    raw: ISetupErrorState,
}

impl SetupErrorState {
    pub fn GetFailedPackages(
        &self,
    ) -> Result<Option<SafeArray<SetupFailedPackageReference>>, HRESULT> {
        unsafe {
            let mut packages = core::ptr::null_mut();
            self.com_ptr()
                .GetFailedPackages(&mut packages)
                .ok_hresult()?;

            if packages.is_null() {
                Ok(None)
            } else {
                SafeArray::from_raw(packages.cast()).map(Some)
            }
        }
    }

    pub fn GetSkippedPackages(&self) -> Result<Option<SafeArray<SetupPackageReference>>, HRESULT> {
        unsafe {
            let mut packages = core::ptr::null_mut();
            self.com_ptr()
                .GetSkippedPackages(&mut packages)
                .ok_hresult()?;
            if packages.is_null() {
                Ok(None)
            } else {
                SafeArray::from_raw(packages.cast()).map(Some)
            }
        }
    }

    pub fn GetErrorLogFilePath(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut path = BSTR::new();
            let state: ISetupErrorState2 = self.com_ptr().cast()?;
            state
                .GetErrorLogFilePath(&mut path)
                .ok_hresult()
                .map(|_| path)
        }
    }

    pub fn GetLogFilePath(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut path = BSTR::new();
            let state: ISetupErrorState2 = self.com_ptr().cast()?;
            state.GetLogFilePath(&mut path).ok_hresult().map(|_| path)
        }
    }

    pub fn GetRuntimeError(&self) -> Result<Option<SetupErrorInfo>, HRESULT> {
        unsafe {
            let mut info = None;
            let state: ISetupErrorState3 = self.com_ptr().cast()?;
            state.GetRuntimeError(&mut info).ok_hresult()?;
            Ok(info.map(|raw| SetupErrorInfo::from_raw(raw)))
        }
    }

    fn com_ptr(&self) -> &ISetupErrorState {
        &self.raw
    }

    unsafe fn from_raw(raw: ISetupErrorState) -> SetupErrorState {
        SetupErrorState { raw }
    }
}

pub struct SetupErrorInfo {
    raw: ISetupErrorInfo,
}

impl SetupErrorInfo {
    pub fn GetErrorHResult(&self) -> Result<HRESULT, HRESULT> {
        unsafe {
            let mut hresult = HRESULT::default();
            self.com_ptr()
                .GetErrorHResult(&mut hresult)
                .ok_hresult()
                .map(|_| hresult)
        }
    }

    pub fn GetErrorClassName(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut name = BSTR::new();
            self.com_ptr()
                .GetErrorClassName(&mut name)
                .ok_hresult()
                .map(|_| name)
        }
    }

    pub fn GetErrorMessage(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut name = BSTR::new();
            self.com_ptr()
                .GetErrorMessage(&mut name)
                .ok_hresult()
                .map(|_| name)
        }
    }

    fn com_ptr(&self) -> &ISetupErrorInfo {
        &self.raw
    }

    unsafe fn from_raw(raw: ISetupErrorInfo) -> SetupErrorInfo {
        SetupErrorInfo { raw }
    }
}

pub struct SetupFailedPackageReference {
    raw: ISetupFailedPackageReference,
}

impl SetupFailedPackageReference {
    pub fn GetLogFilePath(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut path = BSTR::new();
            let package: ISetupFailedPackageReference2 = self.com_ptr().cast()?;
            package.GetLogFilePath(&mut path).ok_hresult().map(|_| path)
        }
    }

    pub fn GetDescription(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut path = BSTR::new();
            let package: ISetupFailedPackageReference2 = self.com_ptr().cast()?;
            package.GetDescription(&mut path).ok_hresult().map(|_| path)
        }
    }

    pub fn GetSignature(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut path = BSTR::new();
            let package: ISetupFailedPackageReference2 = self.com_ptr().cast()?;
            package.GetSignature(&mut path).ok_hresult().map(|_| path)
        }
    }

    pub fn GetDetails(&self) -> Result<SafeArray<BSTR>, HRESULT> {
        unsafe {
            let mut details = null();
            let package: ISetupFailedPackageReference2 = self.com_ptr().cast()?;
            package.GetDetails(&mut details).ok_hresult()?;
            if details.is_null() {
                debug_assert!(!details.is_null());
                Err(E_POINTER)
            } else {
                SafeArray::from_raw(details.cast())
            }
        }
    }

    pub fn GetAffectedPackages(&self) -> Result<Option<SafeArray<SetupPackageReference>>, HRESULT> {
        unsafe {
            let mut packages = null();
            let package: ISetupFailedPackageReference2 = self.com_ptr().cast()?;
            package.GetAffectedPackages(&mut packages).ok_hresult()?;
            if packages.is_null() {
                Ok(None)
            } else {
                SafeArray::from_raw(packages.cast()).map(Some)
            }
        }
    }

    pub fn GetAction(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut path = BSTR::new();
            let package: ISetupFailedPackageReference3 = self.com_ptr().cast()?;
            package.GetAction(&mut path).ok_hresult().map(|_| path)
        }
    }

    pub fn GetReturnCode(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut path = BSTR::new();
            let package: ISetupFailedPackageReference3 = self.com_ptr().cast()?;
            package.GetReturnCode(&mut path).ok_hresult().map(|_| path)
        }
    }

    fn com_ptr(&self) -> &ISetupFailedPackageReference {
        &self.raw
    }
}

impl Deref for SetupFailedPackageReference {
    type Target = SetupPackageReference;
    fn deref(&self) -> &Self::Target {
        unsafe { &*(self as *const Self as *const SetupPackageReference) }
    }
}

pub struct SetupPropertyStore {
    raw: ISetupPropertyStore,
}

impl SetupPropertyStore {
    pub fn GetNames(&self) -> Result<SafeArray<BSTR>, HRESULT> {
        unsafe {
            let mut names = core::ptr::null_mut();
            self.com_ptr().GetNames(&mut names).ok_hresult()?;
            if names.is_null() {
                debug_assert!(!names.is_null());
                Err(E_POINTER)
            } else {
                SafeArray::from_raw(names.cast())
            }
        }
    }

    pub fn GetValue<'w, W: TryInto<WideStr<'w>>>(&self, name: W) -> Result<Variant, HRESULT> {
        let Ok(name) = name.try_into() else {
            return Err(E_INVALIDARG);
        };
        unsafe {
            let mut value = core::mem::zeroed();
            self.com_ptr()
                .GetValue(name.as_ptr(), &mut value)
                .ok_hresult()?;
            Ok(value.into_variant())
        }
    }

    fn com_ptr(&self) -> &ISetupPropertyStore {
        &self.raw
    }

    unsafe fn from_raw(raw: ISetupPropertyStore) -> SetupPropertyStore {
        SetupPropertyStore { raw }
    }
}

pub struct SetupPackageReference {
    raw: ISetupPackageReference,
}

impl SetupPackageReference {
    pub fn GetId(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut id = BSTR::new();
            self.com_ptr().GetId(&mut id).ok_hresult().map(|_| id)
        }
    }

    pub fn GetVersion(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut version = BSTR::new();
            self.com_ptr()
                .GetVersion(&mut version)
                .ok_hresult()
                .map(|_| version)
        }
    }

    pub fn GetChip(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut chip = BSTR::new();
            self.com_ptr().GetChip(&mut chip).ok_hresult().map(|_| chip)
        }
    }

    pub fn GetLanguage(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut lang = BSTR::new();
            self.com_ptr()
                .GetLanguage(&mut lang)
                .ok_hresult()
                .map(|_| lang)
        }
    }

    pub fn GetBranch(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut branch = BSTR::new();
            self.com_ptr()
                .GetBranch(&mut branch)
                .ok_hresult()
                .map(|_| branch)
        }
    }

    pub fn GetType(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut kind = BSTR::new();
            self.com_ptr().GetType(&mut kind).ok_hresult().map(|_| kind)
        }
    }

    pub fn GetUniqueId(&self) -> Result<BSTR, HRESULT> {
        unsafe {
            let mut id = BSTR::new();
            self.com_ptr().GetUniqueId(&mut id).ok_hresult().map(|_| id)
        }
    }

    pub fn GetIsExtension(&self) -> Result<bool, HRESULT> {
        unsafe {
            let mut id = 0;
            self.com_ptr()
                .GetIsExtension(&mut id)
                .ok_hresult()
                .map(|_| id != 0)
        }
    }

    pub fn to_property_store(&self) -> Result<SetupPropertyStore, HRESULT> {
        unsafe {
            self.com_ptr()
                .cast()
                .map(|raw| SetupPropertyStore::from_raw(raw))
                .map_err(Into::into)
        }
    }

    fn com_ptr(&self) -> &ISetupPackageReference {
        &self.raw
    }
}

pub struct SetupInstanceCatalog {
    raw: ISetupInstanceCatalog,
}
impl SetupInstanceCatalog {
    pub fn GetCatalogInfo(&self) -> Result<Option<SetupPropertyStore>, HRESULT> {
        unsafe {
            let mut catalog = None;
            self.com_ptr().GetCatalogInfo(&mut catalog).ok_hresult()?;
            Ok(catalog.map(|raw| SetupPropertyStore::from_raw(raw)))
        }
    }

    pub fn IsPrerelease(&self) -> Result<bool, HRESULT> {
        unsafe {
            let mut is_prerelease = 0;
            self.com_ptr()
                .IsPrerelease(&mut is_prerelease)
                .ok_hresult()?;
            Ok(is_prerelease != 0)
        }
    }

    fn com_ptr(&self) -> &ISetupInstanceCatalog {
        &self.raw
    }

    unsafe fn from_raw(raw: ISetupInstanceCatalog) -> SetupInstanceCatalog {
        SetupInstanceCatalog { raw }
    }
}

/// An owned slice.
///
/// This is roughly equivalent to a `Box<T>`.
/// It will deref to a slice of `T` and be freed on drop.
pub struct SafeArray<T> {
    raw: *mut SAFEARRAY,
    _item: PhantomData<*mut T>,
}

impl<T> SafeArray<T> {
    pub fn iter(&self) -> core::slice::Iter<'_, T> {
        self.as_slice().iter()
    }

    pub fn as_slice(&self) -> &[T] {
        unsafe {
            core::slice::from_raw_parts(
                (*self.raw).pvData.cast::<T>(),
                (*self.raw).rgsabound[0].cElements as usize,
            )
        }
    }

    unsafe fn from_raw(raw: *mut SAFEARRAY) -> Result<Self, HRESULT> {
        unsafe {
            SafeArrayLock(raw).ok_hresult()?;
            if (*raw).cDims != 1 {
                debug_assert_eq!((*raw).cDims, 1);
                // This cannot happen but when it does return an error in release.
                Err(E_UNEXPECTED)
            } else {
                Ok(Self {
                    raw,
                    _item: PhantomData,
                })
            }
        }
    }
}

impl<'a, T> IntoIterator for &'a SafeArray<T> {
    type Item = &'a T;
    type IntoIter = core::slice::Iter<'a, T>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

impl<T> core::ops::Deref for SafeArray<T> {
    type Target = [T];
    fn deref(&self) -> &Self::Target {
        self.as_slice()
    }
}

impl<T> Drop for SafeArray<T> {
    fn drop(&mut self) {
        unsafe {
            let _ = SafeArrayUnlock(self.raw);
            let _ = SafeArrayDestroy(self.raw);
        }
    }
}

trait AssertOk {
    type T;
    fn assert_ok(self) -> Result<Self::T, HRESULT>;
}
impl<T> AssertOk for Option<T> {
    type T = T;

    /// Use this for cases where an API that returns success must also have initialized a COM ptr.
    ///
    /// Panics in debug mode, returns `Err(E_POINTER)` in release mode.
    #[inline(always)]
    fn assert_ok(self) -> Result<T, HRESULT> {
        // If calling this method then this should really really really never happen.
        debug_assert!(self.is_some());
        self.ok_or(E_POINTER)
    }
}

trait OkHresult {
    fn ok_hresult(self) -> Result<(), HRESULT>;
}
impl OkHresult for HRESULT {
    fn ok_hresult(self) -> Result<(), HRESULT> {
        if self.is_ok() { Ok(()) } else { Err(self) }
    }
}

mod api {
    use super::*;
    // Use CoIncrementMTA on win8+?
    #[cfg(not(target_vendor = "win7"))]
    windows_link::link!("combase.dll" "system" fn CoCreateInstance(
    rclsid: *const GUID,
    pUnkOuter: *mut core::ffi::c_void,
    dwClsContext: u32,
    riid: *const GUID,
    ppv: *mut *mut core::ffi::c_void,
) -> HRESULT);
    #[cfg(target_vendor = "win7")]
    windows_link::link!("ole32.dll" "system" fn fn CoCreateInstance(
    rclsid: *const GUID,
    pUnkOuter: *mut core::ffi::c_void,
    dwClsContext: u32,
    riid: *const GUID,
    ppv: *mut *mut core::ffi::c_void,
) -> HRESULT);
    windows_link::link!("oleaut32.dll" "system" fn SafeArrayLock(psa: *const SAFEARRAY) -> HRESULT);
    windows_link::link!("oleaut32.dll" "system" fn SafeArrayUnlock(psa: *const SAFEARRAY) -> HRESULT);
    windows_link::link!("oleaut32.dll" "system" fn SafeArrayDestroy(psa: *const SAFEARRAY) -> HRESULT);
}
use api::*;
