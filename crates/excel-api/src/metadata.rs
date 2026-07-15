//! Explicit argument families used by worksheet-function metadata.
//!
//! These wrappers remove the otherwise ambiguous choice between Excel's `Q`
//! value-only and `U` reference-preserving registration forms. They carry no
//! independent ownership authority; M9B thunks construct them only from values
//! tied to the generated callback scope.

use crate::{ConversionError, ExcelStr, ExcelValueRef, FromExcel};

/// A callback value registered with Excel's value-only `Q` type.
#[derive(Debug)]
pub struct ExcelValueArg<'call>(ExcelValueRef<'call>);

impl<'call> ExcelValueArg<'call> {
    /// Wraps a callback value for a `Q`-registered parameter.
    pub const fn new(value: ExcelValueRef<'call>) -> Self {
        Self(value)
    }

    /// Returns the callback-borrowed value, consuming this metadata wrapper.
    pub fn into_inner(self) -> ExcelValueRef<'call> {
        self.0
    }
}

impl<'call> FromExcel<'call> for ExcelValueArg<'call> {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        Ok(Self::new(value))
    }
}

/// A callback value registered with Excel's reference-preserving `U` type.
#[derive(Debug)]
pub struct ExcelReferenceArg<'call>(ExcelValueRef<'call>);

impl<'call> ExcelReferenceArg<'call> {
    /// Wraps a callback value for a `U`-registered parameter.
    pub const fn new(value: ExcelValueRef<'call>) -> Self {
        Self(value)
    }

    /// Returns the callback-borrowed value, consuming this metadata wrapper.
    pub fn into_inner(self) -> ExcelValueRef<'call> {
        self.0
    }
}

impl<'call> FromExcel<'call> for ExcelReferenceArg<'call> {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        Ok(Self::new(value))
    }
}

/// A direct UTF-16 argument registered with Excel's counted `D%` type.
#[derive(Debug)]
pub struct CountedUtf16Arg<'call>(ExcelStr<'call>);

impl<'call> CountedUtf16Arg<'call> {
    /// Wraps a validated counted direct UTF-16 callback argument.
    pub const fn new(value: ExcelStr<'call>) -> Self {
        Self(value)
    }

    /// Borrows the underlying UTF-16 view for no longer than this callback.
    pub const fn as_excel_str(&self) -> &ExcelStr<'call> {
        &self.0
    }

    /// Returns the callback-borrowed UTF-16 view, consuming this wrapper.
    pub fn into_inner(self) -> ExcelStr<'call> {
        self.0
    }
}

/// A direct UTF-16 argument registered with Excel's NUL-terminated `C%` type.
#[derive(Debug)]
pub struct NullTerminatedUtf16Arg<'call>(ExcelStr<'call>);

impl<'call> NullTerminatedUtf16Arg<'call> {
    /// Wraps a validated NUL-terminated direct UTF-16 callback argument.
    pub const fn new(value: ExcelStr<'call>) -> Self {
        Self(value)
    }

    /// Borrows the underlying UTF-16 view for no longer than this callback.
    pub const fn as_excel_str(&self) -> &ExcelStr<'call> {
        &self.0
    }

    /// Returns the callback-borrowed UTF-16 view, consuming this wrapper.
    pub fn into_inner(self) -> ExcelStr<'call> {
        self.0
    }
}
