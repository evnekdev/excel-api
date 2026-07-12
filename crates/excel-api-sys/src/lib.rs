#![no_std]
#![doc = "Raw, ABI-focused definitions for the Microsoft Excel C API."]
#![allow(non_camel_case_types, non_snake_case, non_upper_case_globals)]

pub mod constants;
pub mod errors;
pub mod functions;
pub mod types;

pub use constants::*;
pub use errors::*;
pub use functions::*;
pub use types::*;
