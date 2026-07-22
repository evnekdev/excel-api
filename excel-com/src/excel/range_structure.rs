//! Structural Range editing, de-duplication, and clipboard-backed operations.

use std::collections::BTreeSet;

use windows_sys::Win32::System::Com::SAFEARRAYBOUND;

use crate::ExcelComError;
use crate::automation::{OwnedVariant, PositionalArguments, SafeArray, invoke};
use crate::excel::formatting::{mixed_bool, property_mixed_get};
use crate::excel::table::{one_based, range_get};
use crate::excel::{MixedValue, Range};
use crate::object_model::{MemberId, member};

macro_rules! raw_structure_type {
    ($(#[$docs:meta])* $name:ident { $($constant:ident = $value:expr => $constant_docs:literal;)* }) => {
        $(#[$docs])*
        #[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
        #[repr(transparent)]
        pub struct $name(i32);
        impl $name {
            $(#[doc = $constant_docs] pub const $constant: Self = Self($value);)*
            /// Creates this value from an Excel type-library integer.
            pub const fn from_raw(value: i32) -> Self { Self(value) }
            /// Returns the raw Excel integer, preserving unknown values.
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}

raw_structure_type! {
    /// A forward-compatible `XlInsertShiftDirection` value.
    InsertShiftDirection {
        DOWN = -4121 => "`xlShiftDown`.";
        RIGHT = -4161 => "`xlShiftToRight`.";
    }
}
raw_structure_type! {
    /// A forward-compatible `XlDeleteShiftDirection` value.
    DeleteShiftDirection {
        LEFT = -4159 => "`xlShiftToLeft`.";
        UP = -4162 => "`xlShiftUp`.";
    }
}
raw_structure_type! {
    /// A forward-compatible `XlInsertFormatOrigin` value.
    InsertFormatOrigin {
        LEFT_OR_ABOVE = 0 => "`xlFormatFromLeftOrAbove`.";
        RIGHT_OR_BELOW = 1 => "`xlFormatFromRightOrBelow`.";
    }
}
raw_structure_type! {
    /// A forward-compatible `XlPasteType` value.
    PasteType {
        ALL = -4104 => "`xlPasteAll`.";
        FORMULAS = -4123 => "`xlPasteFormulas`.";
        VALUES = -4163 => "`xlPasteValues`.";
        FORMATS = -4122 => "`xlPasteFormats`.";
        VALIDATION = 6 => "`xlPasteValidation`.";
    }
}
raw_structure_type! {
    /// A forward-compatible `XlPasteSpecialOperation` value.
    PasteOperation {
        NONE = -4142 => "`xlPasteSpecialOperationNone`.";
        ADD = 2 => "`xlPasteSpecialOperationAdd`.";
        SUBTRACT = 3 => "`xlPasteSpecialOperationSubtract`.";
        MULTIPLY = 4 => "`xlPasteSpecialOperationMultiply`.";
        DIVIDE = 5 => "`xlPasteSpecialOperationDivide`.";
    }
}

/// Options for Excel `Range.RemoveDuplicates`.
///
/// Column indexes are one-based relative to the receiver. The operation
/// modifies and compacts worksheet rows in place.
#[derive(Clone, Debug)]
pub struct RemoveDuplicatesOptions {
    /// One-based columns that form the duplicate key.
    pub columns: Vec<usize>,
    /// Header declaration sent to Excel.
    pub header: crate::TableHeaderMode,
}

/// Options for Excel `Range.Insert`.
#[derive(Debug, Default)]
pub struct RangeInsertOptions {
    /// Optional cell-shift direction; Excel chooses based on shape when absent.
    pub shift: Option<InsertShiftDirection>,
    /// Optional source for copied formats.
    pub copy_origin: Option<InsertFormatOrigin>,
}

/// Explicit options for clipboard-backed Excel `Range.PasteSpecial`.
#[derive(Debug)]
pub struct PasteSpecialOptions {
    /// The portion of copied content to apply.
    pub paste: PasteType,
    /// The arithmetic operation Excel should apply during the paste.
    pub operation: PasteOperation,
    /// Whether blank source cells should be ignored.
    pub skip_blanks: bool,
    /// Whether Excel should transpose rows and columns.
    pub transpose: bool,
}

impl Range {
    /// Returns Excel's contiguous CurrentRegion around this Range.
    ///
    /// Blank rows and columns bound the region; it is not a general logical
    /// dataset detector.
    pub fn current_region(&self) -> Result<Range, ExcelComError> {
        range_get(self.dispatch_object(), "excel.range.currentregion")
    }

    /// Returns whether every selected cell has one hidden state, mixed state, or empty state.
    ///
    /// Excel generally accepts Hidden assignment only for complete rows or
    /// columns. This wrapper does not silently expand arbitrary cell ranges.
    pub fn hidden(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.hidden", mixed_bool)
    }
    /// Changes Hidden state on the exact Range supplied to Excel.
    pub fn set_hidden(&self, hidden: bool) -> Result<(), ExcelComError> {
        let _ = crate::automation::property_put(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.hidden"), true),
            OwnedVariant::bool(hidden),
        )?;
        Ok(())
    }

    /// Inserts cells and lets Excel apply its documented shifting and format-copy rules.
    ///
    /// The installed type library returns `Variant`, not a usable Range, so
    /// this wrapper deliberately returns unit rather than inventing an object.
    pub fn insert(&self, options: &RangeInsertOptions) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_optional(options.shift.map(|value| OwnedVariant::i32(value.raw())));
        args.push_optional(
            options
                .copy_origin
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.insert"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }

    /// Deletes this Range, consumes its wrapper, and lets Excel shift cells as requested.
    pub fn delete(self, shift: Option<DeleteShiftDirection>) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_optional(shift.map(|value| OwnedVariant::i32(value.raw())));
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.delete"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }

    /// Clears values, formulas, formats, comments, and other Excel cell state.
    pub fn clear(&self) -> Result<(), ExcelComError> {
        invoke_unit(self, "excel.range.clear")
    }
    /// Clears cell formatting while retaining cell values and formulas.
    pub fn clear_formats(&self) -> Result<(), ExcelComError> {
        invoke_unit(self, "excel.range.clearformats")
    }
    /// Clears legacy cell comments through Excel's `ClearComments` member.
    pub fn clear_comments(&self) -> Result<(), ExcelComError> {
        invoke_unit(self, "excel.range.clearcomments")
    }
    /// Clears hyperlinks through Excel's `ClearHyperlinks` member.
    pub fn clear_hyperlinks(&self) -> Result<(), ExcelComError> {
        invoke_unit(self, "excel.range.clearhyperlinks")
    }

    /// Copies this Range to an optional destination.
    ///
    /// With no destination Excel enters its cut/copy clipboard mode; callers
    /// should prefer an explicit destination in unattended automation.
    pub fn copy(&self, destination: Option<&Range>) -> Result<(), ExcelComError> {
        copy_or_cut(self, destination, "excel.range.copy")
    }
    /// Cuts this Range to an optional destination.
    ///
    /// With no destination Excel enters its cut/copy clipboard mode.
    pub fn cut(&self, destination: Option<&Range>) -> Result<(), ExcelComError> {
        copy_or_cut(self, destination, "excel.range.cut")
    }

    /// Pastes from Excel's own cut/copy state using explicit options.
    ///
    /// This wrapper neither reads nor writes arbitrary operating-system
    /// clipboard contents. Excel controls the existing cut/copy mode.
    pub fn paste_special(&self, options: &PasteSpecialOptions) -> Result<(), ExcelComError> {
        let args = vec![
            OwnedVariant::i32(options.paste.raw()),
            OwnedVariant::i32(options.operation.raw()),
            OwnedVariant::bool(options.skip_blanks),
            OwnedVariant::bool(options.transpose),
        ];
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.pastespecial-1928"), false),
            args,
            false,
        )?;
        Ok(())
    }

    /// Removes duplicate rows according to one or more one-based receiver-relative columns.
    pub fn remove_duplicates(
        &self,
        options: &RemoveDuplicatesOptions,
    ) -> Result<(), ExcelComError> {
        let columns = encode_duplicate_columns(&options.columns)?;
        let args = vec![columns, OwnedVariant::i32(options.header.raw())];
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.removeduplicates"), false),
            args,
            false,
        )?;
        Ok(())
    }
}

fn invoke_unit(range: &Range, id: &'static str) -> Result<(), ExcelComError> {
    let _ = invoke(
        &range.dispatch_object().dispatch,
        member(MemberId::new(id), false),
        vec![],
        false,
    )?;
    Ok(())
}

fn copy_or_cut(
    range: &Range,
    destination: Option<&Range>,
    id: &'static str,
) -> Result<(), ExcelComError> {
    let mut args = PositionalArguments::new();
    args.push_optional_object(destination.map(Range::dispatch_object));
    let _ = invoke(
        &range.dispatch_object().dispatch,
        member(MemberId::new(id), false),
        args.into_inner(),
        false,
    )?;
    Ok(())
}

fn encode_duplicate_columns(columns: &[usize]) -> Result<OwnedVariant, ExcelComError> {
    if columns.is_empty() {
        return Err(ExcelComError::Unsupported {
            detail: "Range.RemoveDuplicates requires at least one column",
        });
    }
    let mut normalized = BTreeSet::new();
    for column in columns {
        let _ = one_based(*column, "Range.RemoveDuplicates column")?;
        if !normalized.insert(*column) {
            return Err(ExcelComError::Unsupported {
                detail: "Range.RemoveDuplicates columns must be unique",
            });
        }
    }
    let count = u32::try_from(columns.len()).map_err(|_| ExcelComError::Unsupported {
        detail: "too many Range.RemoveDuplicates columns",
    })?;
    let array = SafeArray::create_variant(&[SAFEARRAYBOUND {
        cElements: count,
        lLbound: 0,
    }])
    .ok_or(ExcelComError::Unsupported {
        detail: "could not allocate RemoveDuplicates SAFEARRAY",
    })?;
    for (index, column) in columns.iter().copied().enumerate() {
        let value = OwnedVariant::i32(one_based(column, "Range.RemoveDuplicates column")?);
        let index = i32::try_from(index).map_err(|_| ExcelComError::Unsupported {
            detail: "too many Range.RemoveDuplicates columns",
        })?;
        if !array.put_variant(&[index], &value) {
            return Err(ExcelComError::Unsupported {
                detail: "could not populate RemoveDuplicates SAFEARRAY",
            });
        }
    }
    Ok(OwnedVariant::array(array))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn duplicate_columns_are_rejected_before_com() {
        assert!(encode_duplicate_columns(&[]).is_err());
        assert!(encode_duplicate_columns(&[1, 1]).is_err());
        assert!(encode_duplicate_columns(&[0]).is_err());
        assert!(encode_duplicate_columns(&[1, 2]).is_ok());
    }
}
