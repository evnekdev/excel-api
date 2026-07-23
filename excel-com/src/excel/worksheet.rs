use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{
    ConversionPolicy, OwnedVariant, PositionalArguments, decode_variant, invoke, property_get,
    property_put,
};
use crate::excel::text::text_bstr;
use crate::excel::{
    Application, DispatchObject, FormulaConversionOptions, Names, Range, ReferenceStyle,
};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

/// Excel worksheet visibility value.
///
/// This is intentionally transparent rather than a closed Rust enum: Excel
/// may add values, and a wrapper must preserve a value it reads even if it has
/// not yet assigned it a convenience constant.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct SheetVisibility(i32);

impl SheetVisibility {
    /// A visible sheet (`-1`).
    pub const VISIBLE: Self = Self(-1);
    /// A normally hidden sheet (`0`).
    pub const HIDDEN: Self = Self(0);
    /// A sheet hidden from Excel's normal Unhide UI (`2`).
    pub const VERY_HIDDEN: Self = Self(2);
    /// Preserves a raw value supplied by Excel.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }
    /// Returns the raw Excel value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Experimental wrapper for an Excel `Worksheet`.
pub struct Worksheet {
    inner: DispatchObject,
}

impl Debug for Worksheet {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Worksheet")
            .field(&self.inner)
            .finish()
    }
}

impl Clone for Worksheet {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl Worksheet {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Worksheet",
            },
        }
    }
    pub(crate) fn dispatch_object(&self) -> &DispatchObject {
        &self.inner
    }
    /// Returns whether two Worksheet wrappers denote the same COM object identity.
    pub fn is_same_object(&self, other: &Self) -> Result<bool, ExcelComError> {
        self.inner.same_object(&other.inner)
    }

    /// Returns the worksheet name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.name"), false),
            vec![],
        )?
        .as_string()
    }

    /// Changes the worksheet name.
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.name"), true),
            text_bstr(value)?,
        )?;
        Ok(())
    }

    /// Returns the one-based worksheet index in its workbook.
    pub fn index(&self) -> Result<i32, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.index"), false),
            vec![],
        )?
        .as_i32()
        .ok_or(ExcelComError::Unsupported {
            detail: "Worksheet.Index did not return VT_I4",
        })
    }

    /// Returns the worksheet's visibility state.
    pub fn visible(&self) -> Result<SheetVisibility, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.visible"), false),
            vec![],
        )?
        .as_i32()
        .ok_or(ExcelComError::Unsupported {
            detail: "Worksheet.Visible did not return VT_I4",
        })?;
        Ok(SheetVisibility::from_raw(value))
    }

    /// Changes the worksheet visibility state.
    pub fn set_visible(&self, value: SheetVisibility) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.visible"), true),
            OwnedVariant::i32(value.raw()),
        )?;
        Ok(())
    }

    /// Returns a Range selected by one A1-style reference.
    ///
    /// Excel validates the A1 syntax, including sheet-qualified and full-row
    /// or full-column forms. Embedded NUL is rejected before COM.
    ///
    /// ```no_run
    /// # fn example(worksheet: &excel_com::Worksheet) -> Result<(), excel_com::ExcelComError> {
    /// let range = worksheet.range("A1:X10")?;
    /// # drop(range);
    /// # Ok(())
    /// # }
    /// ```
    pub fn range(&self, a1_reference: &str) -> Result<Range, ExcelComError> {
        self.range_with_arguments(vec![text_bstr(a1_reference)?])
    }

    /// Returns the rectangle delimited by two A1-style corner references.
    ///
    /// Excel validates the corner syntax. This form is useful when callers
    /// already hold each corner separately and do not want to assemble an A1
    /// string themselves.
    pub fn range_between(&self, first_a1: &str, last_a1: &str) -> Result<Range, ExcelComError> {
        self.range_with_arguments(vec![text_bstr(first_a1)?, text_bstr(last_a1)?])
    }

    /// Returns a Range selected from an explicit R1C1 reference.
    ///
    /// The method converts through Excel's `ConvertFormula` engine before
    /// selecting the resulting A1 text. It therefore does not rely on, or
    /// persistently mutate, Excel's global [`ReferenceStyle`] setting.
    ///
    /// ```no_run
    /// # fn example(worksheet: &excel_com::Worksheet) -> Result<(), excel_com::ExcelComError> {
    /// let range = worksheet.range_r1c1("R1C1:R10C24")?;
    /// # drop(range);
    /// # Ok(())
    /// # }
    /// ```
    pub fn range_r1c1(&self, r1c1_reference: &str) -> Result<Range, ExcelComError> {
        if r1c1_reference.contains('\0') {
            return Err(ExcelComError::Unsupported {
                detail: "embedded NUL is not supported by Excel Automation text input",
            });
        }
        let application = self.application()?;
        let converted = application.convert_formula(
            r1c1_reference,
            ReferenceStyle::R1C1,
            ReferenceStyle::A1,
            &FormulaConversionOptions::default(),
        )?;
        self.range(converted.strip_prefix('=').unwrap_or(&converted))
    }

    /// Returns the one-based worksheet cell at `row`, `column`.
    ///
    /// Indices are one-based; zero and values above Excel's `i32` COM boundary
    /// are rejected before any COM invocation.
    pub fn cell(&self, row: usize, column: usize) -> Result<Range, ExcelComError> {
        validate_cell_index(row, "Worksheet.cell row")?;
        validate_cell_index(column, "Worksheet.cell column")?;
        self.cells()?.cell(row, column)
    }

    /// Returns a rectangle bounded by two one-based worksheet coordinates.
    ///
    /// The method obtains Excel's `Worksheet.Cells` Range and supplies the two
    /// resulting Range objects to `Worksheet.Range`; it never generates A1
    /// column letters in Rust.
    ///
    /// ```no_run
    /// # fn example(worksheet: &excel_com::Worksheet) -> Result<(), excel_com::ExcelComError> {
    /// let range = worksheet.range_from_cells(1, 1, 10, 24)?;
    /// # drop(range);
    /// # Ok(())
    /// # }
    /// ```
    pub fn range_from_cells(
        &self,
        first_row: usize,
        first_column: usize,
        last_row: usize,
        last_column: usize,
    ) -> Result<Range, ExcelComError> {
        for (value, detail) in [
            (first_row, "Worksheet.range_from_cells first row"),
            (first_column, "Worksheet.range_from_cells first column"),
            (last_row, "Worksheet.range_from_cells last row"),
            (last_column, "Worksheet.range_from_cells last column"),
        ] {
            validate_cell_index(value, detail)?;
        }
        let cells = self.cells()?;
        let first = cells.cell(first_row, first_column)?;
        let last = cells.cell(last_row, last_column)?;
        let mut arguments = PositionalArguments::new();
        arguments.push_object(first.dispatch_object());
        arguments.push_object(last.dispatch_object());
        self.range_with_arguments(arguments.into_inner())
    }

    /// Returns this worksheet's Name collection for worksheet-local names.
    pub fn names(&self) -> Result<Names, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.names"), false),
            vec![],
        )?;
        Ok(Names::from_dispatch(result.take_dispatch()?))
    }

    /// Evaluates a worksheet-scoped expression as a scalar or array value.
    ///
    /// Worksheet evaluation is useful for local-name resolution. If Excel
    /// returns an object, use [`Self::evaluate_range`] instead.
    pub fn evaluate_value(
        &self,
        expression: &str,
    ) -> Result<crate::AutomationValue, ExcelComError> {
        let result = self.evaluate(expression)?;
        if result.vt() == windows_sys::Win32::System::Variant::VT_DISPATCH {
            return Err(ExcelComError::Unsupported {
                detail: "Worksheet.evaluate_value returned an object; use Worksheet.evaluate_range",
            });
        }
        decode_variant(&result, ConversionPolicy::default())
    }

    /// Evaluates a worksheet-scoped expression as an Excel Range.
    ///
    /// If Excel returns a scalar or array value, use [`Self::evaluate_value`]
    /// instead.
    pub fn evaluate_range(&self, expression: &str) -> Result<Range, ExcelComError> {
        let mut result = self.evaluate(expression)?;
        if result.vt() != windows_sys::Win32::System::Variant::VT_DISPATCH {
            return Err(ExcelComError::Unsupported {
                detail: "Worksheet.evaluate_range returned a value; use Worksheet.evaluate_value",
            });
        }
        Ok(Range::from_dispatch(result.take_dispatch()?))
    }

    /// Calculates this worksheet through Excel.
    ///
    /// Excel defines the dependency scope and completion behavior. Use
    /// [`Application::calculation_state`] when a state snapshot is needed.
    pub fn calculate(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.calculate"), false),
            vec![],
            false,
        )?;
        Ok(())
    }

    fn range_with_arguments(&self, arguments: Vec<OwnedVariant>) -> Result<Range, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.range"), false),
            arguments,
        )?;
        Ok(Range::from_dispatch(result.take_dispatch()?))
    }

    fn cells(&self) -> Result<Range, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.cells"), false),
            vec![],
        )?;
        Ok(Range::from_dispatch(result.take_dispatch()?))
    }

    fn application(&self) -> Result<Application, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.application"), false),
            vec![],
        )?;
        Ok(Application::from_dispatch(result.take_dispatch()?))
    }

    fn evaluate(&self, expression: &str) -> Result<OwnedVariant, ExcelComError> {
        invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.evaluate-1"), false),
            vec![text_bstr(expression)?],
            false,
        )
    }

    /// Returns the worksheet's current used Range.
    pub fn used_range(&self) -> Result<Range, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.usedrange"), false),
            vec![],
        )?;
        Ok(Range::from_dispatch(result.take_dispatch()?))
    }
}

fn validate_cell_index(value: usize, detail: &'static str) -> Result<(), ExcelComError> {
    if value == 0 {
        return Err(ExcelComError::Unsupported {
            detail: "Worksheet row and column indices are one-based and nonzero",
        });
    }
    let _ = i32::try_from(value).map_err(|_| ExcelComError::Unsupported { detail })?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numeric_indices_reject_zero_and_overflow_before_com() {
        assert!(validate_cell_index(0, "test").is_err());
        assert!(validate_cell_index(i32::MAX as usize + 1, "test").is_err());
        assert!(validate_cell_index(1, "test").is_ok());
    }
}
