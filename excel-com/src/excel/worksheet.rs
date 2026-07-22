use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{AutomationArgument, OwnedVariant, property_get, property_put};
use crate::excel::{DispatchObject, Range};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

/// Excel worksheet visibility values accepted by the core wrapper.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[repr(i32)]
pub enum XlSheetVisibility {
    /// Visible worksheet (`-1`).
    Visible = -1,
    /// Hidden worksheet (`0`).
    Hidden = 0,
    /// Very hidden worksheet (`2`).
    VeryHidden = 2,
}

impl TryFrom<i32> for XlSheetVisibility {
    type Error = ExcelComError;

    /// Converts Excel's numeric visibility result into the supported enum.
    fn try_from(value: i32) -> Result<Self, Self::Error> {
        match value {
            -1 => Ok(Self::Visible),
            0 => Ok(Self::Hidden),
            2 => Ok(Self::VeryHidden),
            _ => Err(ExcelComError::Unsupported {
                detail: "unknown worksheet visibility value",
            }),
        }
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
            OwnedVariant::bstr(value)?,
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
    pub fn visible(&self) -> Result<XlSheetVisibility, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.visible"), false),
            vec![],
        )?
        .as_i32()
        .ok_or(ExcelComError::Unsupported {
            detail: "Worksheet.Visible did not return VT_I4",
        })?;
        XlSheetVisibility::try_from(value)
    }

    /// Changes the worksheet visibility state.
    pub fn set_visible(&self, value: XlSheetVisibility) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.visible"), true),
            OwnedVariant::i32(value as i32),
        )?;
        Ok(())
    }

    /// Returns a Range selected by `cell1` and an optional `cell2`.
    ///
    /// Supply a text value such as `"A1"` or `"A1:B2"` for ordinary A1
    /// notation. Omitting `cell2` omits the trailing Automation argument.
    pub fn range(
        &self,
        cell1: AutomationArgument,
        cell2: Option<AutomationArgument>,
    ) -> Result<Range, ExcelComError> {
        let policy = crate::ConversionPolicy::default();
        let mut arguments = vec![cell1.encode(policy)?];
        if let Some(cell2) = cell2 {
            arguments.push(cell2.encode(policy)?);
        }
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheet.range"), false),
            arguments,
        )?;
        Ok(Range::from_dispatch(result.take_dispatch()?))
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
