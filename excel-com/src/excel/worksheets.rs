use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{AutomationArgument, OwnedVariant, invoke, property_get};
use crate::excel::{DispatchObject, Worksheet};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

/// Options for [`Worksheets::add`].
///
/// The `Type` Automation parameter is deliberately always `Missing`: this
/// bounded wrapper does not implement Excel's alternate sheet types.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct WorksheetsAddOptions {
    before: Option<AutomationArgument>,
    after: Option<AutomationArgument>,
    count: Option<u32>,
}

impl WorksheetsAddOptions {
    /// Starts a default `Worksheets.Add` request.
    pub const fn new() -> Self {
        Self {
            before: None,
            after: None,
            count: None,
        }
    }

    /// Supplies Excel's optional `Before` parameter.
    pub fn before(mut self, value: AutomationArgument) -> Self {
        self.before = Some(value);
        self
    }

    /// Supplies Excel's optional `After` parameter.
    pub fn after(mut self, value: AutomationArgument) -> Self {
        self.after = Some(value);
        self
    }

    /// Supplies a positive optional `Count` parameter.
    pub fn count(mut self, value: u32) -> Self {
        self.count = Some(value);
        self
    }
}

/// Experimental wrapper for an Excel `Worksheets` collection.
pub struct Worksheets {
    inner: DispatchObject,
}

impl Debug for Worksheets {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Worksheets")
            .field(&self.inner)
            .finish()
    }
}

impl Clone for Worksheets {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl Worksheets {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Worksheets",
            },
        }
    }

    /// Returns the number of worksheets in the collection.
    pub fn count(&self) -> Result<i32, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheets.count"), false),
            vec![],
        )?
        .as_i32()
        .ok_or(ExcelComError::Unsupported {
            detail: "Worksheets.Count did not return VT_I4",
        })
    }

    /// Returns the one-based worksheet at `index`.
    pub fn item_by_index(&self, index: u32) -> Result<Worksheet, ExcelComError> {
        if index == 0 {
            return Err(ExcelComError::Unsupported {
                detail: "worksheet index is one-based",
            });
        }
        let index = i32::try_from(index).map_err(|_| ExcelComError::Unsupported {
            detail: "worksheet index exceeds i32",
        })?;
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheets.item"), false),
            vec![OwnedVariant::i32(index)],
        )?;
        Ok(Worksheet::from_dispatch(result.take_dispatch()?))
    }

    /// Returns the worksheet selected by its current name.
    pub fn item_by_name(&self, name: &str) -> Result<Worksheet, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheets.item"), false),
            vec![OwnedVariant::bstr(name)?],
        )?;
        Ok(Worksheet::from_dispatch(result.take_dispatch()?))
    }

    /// Adds a worksheet using optional arguments in their logical Excel order.
    ///
    /// `Before` and `After` are mutually exclusive. Each missing optional
    /// position is encoded as `VT_ERROR` / `DISP_E_PARAMNOTFOUND`; the dispatch
    /// layer reverses the four logical arguments exactly once for COM.
    pub fn add(&self, options: WorksheetsAddOptions) -> Result<Worksheet, ExcelComError> {
        if options.before.is_some() && options.after.is_some() {
            return Err(ExcelComError::Unsupported {
                detail: "Worksheets.Add does not permit both Before and After",
            });
        }
        if options.count == Some(0) {
            return Err(ExcelComError::Unsupported {
                detail: "Worksheets.Add Count must be positive",
            });
        }
        let policy = crate::ConversionPolicy::default();
        let before = options
            .before
            .unwrap_or(AutomationArgument::Missing)
            .encode(policy)?;
        let after = options
            .after
            .unwrap_or(AutomationArgument::Missing)
            .encode(policy)?;
        let count =
            match options.count {
                Some(value) => OwnedVariant::i32(i32::try_from(value).map_err(|_| {
                    ExcelComError::Unsupported {
                        detail: "Worksheets.Add Count exceeds i32",
                    }
                })?),
                None => AutomationArgument::Missing.encode(policy)?,
            };
        let sheet_type = AutomationArgument::Missing.encode(policy)?;
        let mut result = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheets.add"), false),
            vec![before, after, count, sheet_type],
            false,
        )?;
        Ok(Worksheet::from_dispatch(result.take_dispatch()?))
    }
}
