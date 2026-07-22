use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{invoke, property_get};
use crate::excel::{DispatchObject, Range};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

/// Experimental wrapper for one Excel-defined Name.
///
/// A valid Excel Name need not resolve to a Range: it can denote a constant,
/// a formula, an invalid reference, or an external target. Consequently,
/// [`Self::range`] is deliberately fallible.
pub struct Name {
    inner: DispatchObject,
}

impl Debug for Name {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Name").field(&self.inner).finish()
    }
}

impl Clone for Name {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl Name {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Name",
            },
        }
    }

    /// Returns the name text Excel reports, including local-scope qualification when Excel uses it.
    pub fn name(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.name.name"), false),
            vec![],
        )?
        .as_string()
    }

    /// Returns Excel's A1-oriented `RefersTo` expression.
    ///
    /// Excel commonly returns a leading `=`. The exact text, including
    /// workbook and worksheet qualification, remains Excel-owned.
    pub fn refers_to(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.name.refersto"), false),
            vec![],
        )?
        .as_string()
    }

    /// Returns Excel's R1C1-oriented `RefersToR1C1` expression.
    pub fn refers_to_r1c1(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.name.referstor1c1"), false),
            vec![],
        )?
        .as_string()
    }

    /// Returns whether the Name is visible in Excel's user interface.
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.name.visible"), false),
            vec![],
        )?;
        result.as_bool().ok_or(ExcelComError::Conversion(
            crate::ConversionError::UnsupportedVariantType {
                vartype: result.vt(),
            },
        ))
    }

    /// Resolves the Name using Excel's `RefersToRange` property.
    ///
    /// This can fail for non-Range names even when [`Self::refers_to`] is a
    /// valid expression. Excel's structured invocation error is preserved.
    ///
    /// ```no_run
    /// # fn example(name: &excel_com::Name) -> Result<(), excel_com::ExcelComError> {
    /// let resolved = name.range()?;
    /// # drop(resolved);
    /// # Ok(())
    /// # }
    /// ```
    pub fn range(&self) -> Result<Range, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.name.referstorange"), false),
            vec![],
        )?;
        Ok(Range::from_dispatch(result.take_dispatch()?))
    }

    /// Deletes this Name in Excel.
    ///
    /// The wrapper is consumed regardless of whether Excel accepts the
    /// deletion request, preventing further use of a potentially stale Name.
    pub fn delete(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.name.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}
