#![no_std]
#![warn(missing_docs)]
#![doc = include_str!("../README.md")]
//!
//! # Raw Excel 12 ABI layer
//!
//! This crate mirrors the checked-in Microsoft `XLCALL.H` Excel 12 C API
//! contract: fixed-width aliases, constants, `XLOPER12`-family layouts, and
//! callback function-pointer signatures. It is intended for audited FFI
//! boundaries on 64-bit Windows Excel, not for ordinary application code.
//!
//! The types and ownership bits describe an ABI; they do **not** make an
//! allocation safe to retain, mutate, free, or send across threads. In
//! particular, an `xlbitDLLFree` or `xlbitXLFree` tag is Excel protocol data,
//! not a Rust ownership type. Prefer [`excel_api`](https://docs.rs/excel-api)
//! for borrowed callback views, owned values, return allocation, and typed
//! callback contexts.
//!
//! # Direct-call safety
//!
//! A direct caller must verify the Excel callback context, `extern "system"`
//! calling convention, pointer alignment and accessible lengths, union tag,
//! allocation origin, and exact cleanup protocol. The initial support target is
//! 64-bit Excel; 32-bit layouts and non-Windows Excel hosts are unsupported.
//!
//! ```no_run
//! use excel_api_sys::{Xloper12, xltypeNum};
//!
//! let value = Xloper12 { xltype: xltypeNum, val: excel_api_sys::Xloper12Val { num: 42.0 } };
//! assert_eq!(value.xltype, xltypeNum);
//! ```
#![allow(non_camel_case_types, non_snake_case, non_upper_case_globals)]

/// Raw SDK constants and function IDs.
///
/// The individual identifiers intentionally preserve `XLCALL.H` spelling and
/// are a mechanically audited compatibility surface; see the module-level
/// ownership and ABI rules above rather than treating a constant as a safe API.
#[allow(missing_docs)]
pub mod constants;
/// Excel worksheet error codes.
pub mod errors;
/// Raw Excel callback and entry-point function signatures.
pub mod functions;
/// Raw C-layout Excel structures, unions, and primitive aliases.
pub mod types;

pub use constants::*;
pub use errors::*;
pub use functions::*;
pub use types::*;
