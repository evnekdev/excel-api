//! Internal M18.1 RTD compatibility prototype.
//!
//! This package is unpublished and intentionally separate from the XLL crates.

#![cfg_attr(not(windows), allow(dead_code))]
// MSVC warns that the two conventional COM exports could be PRIVATE in an
// import-library definition; the DLL intentionally exports them publicly.
#![allow(linker_messages)]

mod model;

#[cfg(windows)]
mod windows_server;

#[cfg(not(windows))]
mod non_windows {
    // Keeps ordinary non-Windows workspace builds possible without COM.
    pub(crate) const PLATFORM_STATUS: &str = "RTD prototype requires Windows";
}
