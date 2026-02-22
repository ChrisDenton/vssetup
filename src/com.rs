//! Helpers for initailizing and uninitalizing COM.
//!
//! The API in this crate relies on COM being initialized for the duration of its use.
//! [`SetupConfiguration::new`](crate::SetupConfiguration::new) will error if COM isn't initialized.
//!
//! **WARNGING**: Using any API in this crate after COM is unitilized is Undefined Behaviour (UB).
//! If in doubt it is safer to simply not call [`uninitialize`].

use windows_result::HRESULT;

/// Runs the given function with COM initalized and uninitalizes COM afterward.
///
/// # Safety
///
/// See [`uninitialize`].
pub unsafe fn with_com<R, F: FnOnce() -> R>(f: F) -> Result<R, HRESULT> {
    initialize()?;
    let result = f();
    // SAFETY: the caller must ensure this is safe.
    unsafe { uninitialize() };
    Ok(result)
}

/// Initialize COM.
///
/// This needs to be called before any COM objects are created or used.
pub fn initialize() -> Result<(), HRESULT> {
    let result = unsafe { CoInitializeEx(core::ptr::null(), 0) };
    if result.is_ok() { Ok(()) } else { Err(result) }
}

/// Unitialize COM.
///
/// # Safety
///
/// - This must be called on the same thread that called [`initialize`].
/// - You must ensure there are no COM objects still in use before calling this.
///
/// **WARNING**: Beware of `drop` implementations that may use COM objects.
/// Calling this directly will run before any drops that are in scope.
///
/// ## Safe example
///
/// ```rust
/// use vssetup::{com, HRESULT};
///
/// fn main() -> Result<(), HRESULT> {
/// com::initialize()?;
/// {
///     // do COM stuff
/// }
///
/// // SAFETY: All uses of COM are contained and dropped by the scope above.
/// # if false { // Doing these here may interfere with other tests.
/// unsafe { com::uninitialize() };
/// # }
///
/// Ok(())
/// }
/// ```
pub unsafe fn uninitialize() {
    unsafe {
        CoUnInitialize();
    }
}

mod api {
    use super::HRESULT;
    #[cfg(not(target_vendor = "win7"))]
    windows_link::link!("combase.dll" "system" fn CoInitializeEx(pvReserved: *const (), dwCoInit: u32) -> HRESULT);
    #[cfg(target_vendor = "win7")]
    windows_link::link!("ole32.dll" "system" fn CoInitializeEx(pvReserved: *const (), dwCoInit: u32) -> HRESULT);
    #[cfg(not(target_vendor = "win7"))]
    windows_link::link!("combase.dll" "system" fn CoUnInitialize());
    #[cfg(target_vendor = "win7")]
    windows_link::link!("ole32.dll" "system" fn CoUnInitialize());
}
use api::*;
