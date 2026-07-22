use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{AutomationArgument, OwnedVariant, PositionalArguments, invoke};
use crate::automation::{EnumVariant, enumerated_dispatch};
use crate::excel::collection::{
    CollectionDescriptor, count as collection_count, enumerator, item_by_index, item_by_name,
};
use crate::excel::{DispatchObject, Worksheet};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use std::iter::FusedIterator;

const DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Worksheets",
    count: MemberId::new("excel.worksheets.count"),
    item: MemberId::new("excel.worksheets.item"),
    new_enum: MemberId::new("excel.worksheets.newenum"),
};

/// Options for [`Worksheets::add`].
///
/// The `Type` Automation parameter is deliberately always `Missing`: this
/// bounded wrapper does not implement Excel's alternate sheet types.
#[derive(Clone, Debug, Default)]
pub struct WorksheetsAddOptions<'a> {
    before: Option<AutomationArgument>,
    after: Option<AutomationArgument>,
    before_worksheet: Option<&'a Worksheet>,
    after_worksheet: Option<&'a Worksheet>,
    count: Option<u32>,
}

impl<'a> WorksheetsAddOptions<'a> {
    /// Starts a default `Worksheets.Add` request.
    pub const fn new() -> Self {
        Self {
            before: None,
            after: None,
            before_worksheet: None,
            after_worksheet: None,
            count: None,
        }
    }

    /// Supplies Excel's optional `Before` parameter.
    pub fn before(mut self, value: AutomationArgument) -> Self {
        self.before = Some(value);
        self
    }

    /// Supplies a worksheet object as Excel's optional `Before` parameter.
    pub fn before_worksheet(mut self, value: &'a Worksheet) -> Self {
        self.before_worksheet = Some(value);
        self
    }

    /// Supplies Excel's optional `After` parameter.
    pub fn after(mut self, value: AutomationArgument) -> Self {
        self.after = Some(value);
        self
    }

    /// Supplies a worksheet object as Excel's optional `After` parameter.
    pub fn after_worksheet(mut self, value: &'a Worksheet) -> Self {
        self.after_worksheet = Some(value);
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
        i32::try_from(collection_count(&self.inner, DESCRIPTOR)?).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "Worksheets.Count exceeds i32",
            }
        })
    }

    /// Returns the one-based worksheet at `index`.
    pub fn item_by_index(&self, index: usize) -> Result<Worksheet, ExcelComError> {
        Ok(Worksheet::from_dispatch(item_by_index(
            &self.inner,
            DESCRIPTOR,
            index,
        )?))
    }

    /// Returns the worksheet selected by its current name.
    pub fn item_by_name(&self, name: &str) -> Result<Worksheet, ExcelComError> {
        Ok(Worksheet::from_dispatch(item_by_name(
            &self.inner,
            DESCRIPTOR,
            name,
        )?))
    }

    /// Adds a worksheet using optional arguments in their logical Excel order.
    ///
    /// `Before` and `After` are mutually exclusive. Each missing optional
    /// position is encoded as `VT_ERROR` / `DISP_E_PARAMNOTFOUND`; the dispatch
    /// layer reverses the four logical arguments exactly once for COM.
    pub fn add(&self, options: WorksheetsAddOptions<'_>) -> Result<Worksheet, ExcelComError> {
        if (options.before.is_some() || options.before_worksheet.is_some())
            && (options.after.is_some() || options.after_worksheet.is_some())
        {
            return Err(ExcelComError::Unsupported {
                detail: "Worksheets.Add does not permit both Before and After",
            });
        }
        if options.count == Some(0) {
            return Err(ExcelComError::Unsupported {
                detail: "Worksheets.Add Count must be positive",
            });
        }
        if options.count.is_some_and(|value| value > i32::MAX as u32) {
            return Err(ExcelComError::Unsupported {
                detail: "Worksheets.Add Count exceeds i32",
            });
        }
        let policy = crate::ConversionPolicy::default();
        let mut arguments = PositionalArguments::new();
        match (options.before, options.before_worksheet) {
            (Some(value), None) => arguments.push_argument(value, policy)?,
            (None, Some(value)) => arguments.push_object(value.dispatch_object()),
            (None, None) => arguments.push_optional(None),
            (Some(_), Some(_)) => {
                return Err(ExcelComError::Unsupported {
                    detail: "Worksheets.Add has duplicate Before values",
                });
            }
        }
        match (options.after, options.after_worksheet) {
            (Some(value), None) => arguments.push_argument(value, policy)?,
            (None, Some(value)) => arguments.push_object(value.dispatch_object()),
            (None, None) => arguments.push_optional(None),
            (Some(_), Some(_)) => {
                return Err(ExcelComError::Unsupported {
                    detail: "Worksheets.Add has duplicate After values",
                });
            }
        }
        arguments.push_optional(options.count.map(|value| OwnedVariant::i32(value as i32)));
        arguments.push_optional(None);
        let mut result = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.worksheets.add"), false),
            arguments.into_inner(),
            false,
        )?;
        Ok(Worksheet::from_dispatch(result.take_dispatch()?))
    }
    /// Iterates worksheet objects in Excel's `_NewEnum` order.
    pub fn iter(&self) -> Result<WorksheetsIter, ExcelComError> {
        Ok(WorksheetsIter {
            enumerator: enumerator(&self.inner, DESCRIPTOR)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Typed, single-pass worksheet collection iterator.
pub struct WorksheetsIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for WorksheetsIter {
    type Item = Result<Worksheet, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(
                    enumerated_dispatch(&mut value, "Worksheets", index)
                        .map(Worksheet::from_dispatch),
                )
            }
            Ok(None) => {
                self.terminal = true;
                None
            }
            Err(error) => {
                self.terminal = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for WorksheetsIter {}
