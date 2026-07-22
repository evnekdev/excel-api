//! Excel Range data-validation wrappers.

use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{OwnedVariant, PositionalArguments, invoke};
use crate::excel::table::{bool_get, bool_put, i32_get, object_get, string_get, text_put};
use crate::excel::{DispatchObject, Range};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

macro_rules! raw_validation_type {
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

raw_validation_type! {
    /// A forward-compatible `XlDVType` value.
    ValidationType {
        INPUT_ONLY = 0 => "`xlValidateInputOnly`.";
        WHOLE_NUMBER = 1 => "`xlValidateWholeNumber`.";
        DECIMAL = 2 => "`xlValidateDecimal`.";
        LIST = 3 => "`xlValidateList`.";
        DATE = 4 => "`xlValidateDate`.";
        TIME = 5 => "`xlValidateTime`.";
        TEXT_LENGTH = 6 => "`xlValidateTextLength`.";
        CUSTOM = 7 => "`xlValidateCustom`.";
    }
}
raw_validation_type! {
    /// A forward-compatible `XlDVAlertStyle` value.
    ValidationAlertStyle {
        STOP = 1 => "`xlValidAlertStop`.";
        WARNING = 2 => "`xlValidAlertWarning`.";
        INFORMATION = 3 => "`xlValidAlertInformation`.";
    }
}
raw_validation_type! {
    /// A forward-compatible `XlFormatConditionOperator` value used by validation.
    ValidationOperator {
        BETWEEN = 1 => "`xlBetween`.";
        NOT_BETWEEN = 2 => "`xlNotBetween`.";
        EQUAL = 3 => "`xlEqual`.";
        NOT_EQUAL = 4 => "`xlNotEqual`.";
        GREATER = 5 => "`xlGreater`.";
        LESS = 6 => "`xlLess`.";
        GREATER_EQUAL = 7 => "`xlGreaterEqual`.";
        LESS_EQUAL = 8 => "`xlLessEqual`.";
    }
}

/// Typed positional input for Excel `Validation.Add`.
///
/// Validation formulas remain Excel syntax: list validation may use a
/// comma-separated literal or a worksheet/name reference according to Excel's
/// locale and formula rules. This wrapper does not parse, translate, or impose
/// UI text-length limits on them.
#[derive(Debug)]
pub struct ValidationAddOptions<'a> {
    /// The validation kind.
    pub validation_type: ValidationType,
    /// Optional alert style.
    pub alert_style: Option<ValidationAlertStyle>,
    /// Optional comparison operator.
    pub operator: Option<ValidationOperator>,
    /// Optional first Excel validation formula.
    pub formula1: Option<&'a str>,
    /// Optional second Excel validation formula.
    pub formula2: Option<&'a str>,
}

/// Apartment-bound wrapper for a Range's Excel Validation object.
pub struct Validation {
    inner: DispatchObject,
}
impl Debug for Validation {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Validation").field(&self.inner).finish()
    }
}
impl Clone for Validation {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Validation {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Validation",
            },
        }
    }
    /// Adds validation through Excel's five-position `Validation.Add` call.
    pub fn add(&self, options: &ValidationAddOptions<'_>) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(options.validation_type.raw()));
        args.push_optional(
            options
                .alert_style
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(options.operator.map(|value| OwnedVariant::i32(value.raw())));
        push_optional_text(&mut args, options.formula1)?;
        push_optional_text(&mut args, options.formula2)?;
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.validation.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }
    /// Removes all validation from the underlying Range.
    pub fn delete(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.validation.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Returns the validation type.
    pub fn validation_type(&self) -> Result<ValidationType, ExcelComError> {
        Ok(ValidationType::from_raw(i32_get(
            &self.inner,
            "excel.validation.type",
            "Validation.Type",
        )?))
    }
    /// Returns the configured alert style.
    pub fn alert_style(&self) -> Result<ValidationAlertStyle, ExcelComError> {
        Ok(ValidationAlertStyle::from_raw(i32_get(
            &self.inner,
            "excel.validation.alertstyle",
            "Validation.AlertStyle",
        )?))
    }
    /// Returns the configured comparison operator.
    pub fn operator(&self) -> Result<ValidationOperator, ExcelComError> {
        Ok(ValidationOperator::from_raw(i32_get(
            &self.inner,
            "excel.validation.operator",
            "Validation.Operator",
        )?))
    }
    /// Returns Formula1 in Excel's formula syntax.
    pub fn formula1(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.validation.formula1")
    }
    /// Returns Formula2 in Excel's formula syntax.
    pub fn formula2(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.validation.formula2")
    }
    /// Returns Excel's IgnoreBlank setting.
    pub fn ignore_blank(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.validation.ignoreblank")
    }
    /// Changes Excel's IgnoreBlank setting.
    pub fn set_ignore_blank(&self, enabled: bool) -> Result<(), ExcelComError> {
        bool_put(&self.inner, "excel.validation.ignoreblank", enabled)
    }
    /// Returns Excel's InCellDropdown setting.
    pub fn in_cell_dropdown(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.validation.incelldropdown")
    }
    /// Changes Excel's InCellDropdown setting.
    pub fn set_in_cell_dropdown(&self, enabled: bool) -> Result<(), ExcelComError> {
        bool_put(&self.inner, "excel.validation.incelldropdown", enabled)
    }
    /// Returns whether Excel displays an input message.
    pub fn show_input(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.validation.showinput")
    }
    /// Changes whether Excel displays an input message.
    pub fn set_show_input(&self, enabled: bool) -> Result<(), ExcelComError> {
        bool_put(&self.inner, "excel.validation.showinput", enabled)
    }
    /// Returns whether Excel displays an error alert.
    pub fn show_error(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.validation.showerror")
    }
    /// Changes whether Excel displays an error alert.
    pub fn set_show_error(&self, enabled: bool) -> Result<(), ExcelComError> {
        bool_put(&self.inner, "excel.validation.showerror", enabled)
    }
    /// Returns Excel's input-message title.
    pub fn input_title(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.validation.inputtitle")
    }
    /// Changes Excel's input-message title after rejecting embedded NUL.
    pub fn set_input_title(&self, title: &str) -> Result<(), ExcelComError> {
        text_put(&self.inner, "excel.validation.inputtitle", title)
    }
    /// Returns Excel's input-message text.
    pub fn input_message(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.validation.inputmessage")
    }
    /// Changes Excel's input-message text after rejecting embedded NUL.
    pub fn set_input_message(&self, message: &str) -> Result<(), ExcelComError> {
        text_put(&self.inner, "excel.validation.inputmessage", message)
    }
    /// Returns Excel's validation error title.
    pub fn error_title(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.validation.errortitle")
    }
    /// Changes Excel's validation error title after rejecting embedded NUL.
    pub fn set_error_title(&self, title: &str) -> Result<(), ExcelComError> {
        text_put(&self.inner, "excel.validation.errortitle", title)
    }
    /// Returns Excel's validation error message.
    pub fn error_message(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.validation.errormessage")
    }
    /// Changes Excel's validation error message after rejecting embedded NUL.
    pub fn set_error_message(&self, message: &str) -> Result<(), ExcelComError> {
        text_put(&self.inner, "excel.validation.errormessage", message)
    }
    /// Returns whether Excel considers the current value valid for this Validation object.
    pub fn is_value_valid(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.validation.value")
    }
}

impl Range {
    /// Returns the apartment-bound Excel Validation object for this Range.
    pub fn validation(&self) -> Result<Validation, ExcelComError> {
        object_get(
            self.dispatch_object(),
            "excel.range.validation",
            Validation::from_dispatch,
        )
    }
}

fn push_optional_text(
    args: &mut PositionalArguments,
    value: Option<&str>,
) -> Result<(), ExcelComError> {
    match value {
        Some(value) => args.push_result(crate::excel::text::text_bstr(value))?,
        None => args.push_optional(None),
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn validation_add_keeps_all_optional_positions() {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(ValidationType::LIST.raw()));
        args.push_optional(None);
        args.push_optional(None);
        push_optional_text(&mut args, Some("A,B")).expect("text");
        push_optional_text(&mut args, None).expect("text");
        let values = args.into_inner();
        assert_eq!(values.len(), 5);
        assert_eq!(values[1].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[4].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }
}
