//! Explicit argument families used by worksheet-function metadata.
//!
//! These wrappers remove the otherwise ambiguous choice between Excel's `Q`
//! value-only and `U` reference-preserving registration forms. They carry no
//! ABI or ownership behavior; the M9B thunk milestone will construct them from
//! callback-scoped values.

use crate::{ExcelStr, ExcelValueRef};

/// A callback value registered with Excel's value-only `Q` type.
#[derive(Debug)]
pub struct ExcelValueArg<'call>(ExcelValueRef<'call>);

impl<'call> ExcelValueArg<'call> {
    pub const fn new(value: ExcelValueRef<'call>) -> Self {
        Self(value)
    }

    pub fn into_inner(self) -> ExcelValueRef<'call> {
        self.0
    }
}

/// A callback value registered with Excel's reference-preserving `U` type.
#[derive(Debug)]
pub struct ExcelReferenceArg<'call>(ExcelValueRef<'call>);

impl<'call> ExcelReferenceArg<'call> {
    pub const fn new(value: ExcelValueRef<'call>) -> Self {
        Self(value)
    }

    pub fn into_inner(self) -> ExcelValueRef<'call> {
        self.0
    }
}

/// A direct UTF-16 argument registered with Excel's counted `D%` type.
#[derive(Debug)]
pub struct CountedUtf16Arg<'call>(ExcelStr<'call>);

impl<'call> CountedUtf16Arg<'call> {
    pub const fn new(value: ExcelStr<'call>) -> Self {
        Self(value)
    }

    pub const fn as_excel_str(&self) -> &ExcelStr<'call> {
        &self.0
    }

    pub fn into_inner(self) -> ExcelStr<'call> {
        self.0
    }
}

/// A direct UTF-16 argument registered with Excel's NUL-terminated `C%` type.
#[derive(Debug)]
pub struct NullTerminatedUtf16Arg<'call>(ExcelStr<'call>);

impl<'call> NullTerminatedUtf16Arg<'call> {
    pub const fn new(value: ExcelStr<'call>) -> Self {
        Self(value)
    }

    pub const fn as_excel_str(&self) -> &ExcelStr<'call> {
        &self.0
    }

    pub fn into_inner(self) -> ExcelStr<'call> {
        self.0
    }
}
