use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{EnumVariant, enumerated_dispatch};
use crate::automation::{OwnedVariant, PositionalArguments, invoke};
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

/// A concrete worksheet kind accepted by `Worksheets.Add`.
///
/// The transparent representation preserves future Excel values rather than
/// narrowing this forward-compatible Automation enum to today's variants.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct SheetType(i32);

impl SheetType {
    /// A standard worksheet.
    pub const WORKSHEET: Self = Self(-4167);
    /// An Excel 4 macro sheet.
    pub const EXCEL4_MACRO: Self = Self(3);
    /// An international Excel 4 macro sheet.
    pub const EXCEL4_INTL_MACRO: Self = Self(4);
    /// A chart sheet.
    pub const CHART: Self = Self(-4109);
    /// A dialog sheet.
    pub const DIALOG: Self = Self(-4116);
    /// Retains an Excel value not yet named by this wrapper.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }
    /// Returns Excel's numeric sheet-type value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Optional positions for [`Worksheets::add`].
///
/// `before` and `after` are mutually exclusive. `sheet_type` intentionally
/// remains explicit: requesting a non-worksheet type is rejected because this
/// typed collection cannot truthfully return that heterogeneous sheet object.
#[derive(Clone, Copy, Debug, Default)]
pub struct WorksheetAddOptions<'a> {
    /// Inserts before this worksheet.
    pub before: Option<&'a Worksheet>,
    /// Inserts after this worksheet.
    pub after: Option<&'a Worksheet>,
    /// Adds this many worksheets; omitted lets Excel choose one.
    pub count: Option<usize>,
    /// Requests an Excel sheet type.
    pub sheet_type: Option<SheetType>,
}

/// Compatibility name for the earlier worksheet-only options type.
pub type WorksheetsAddOptions<'a> = WorksheetAddOptions<'a>;

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
    pub fn add(&self, options: &WorksheetAddOptions<'_>) -> Result<Worksheet, ExcelComError> {
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
        if options.count.is_some_and(|value| value > i32::MAX as usize) {
            return Err(ExcelComError::Unsupported {
                detail: "Worksheets.Add Count exceeds i32",
            });
        }
        if options
            .sheet_type
            .is_some_and(|value| value != SheetType::WORKSHEET)
        {
            return Err(ExcelComError::Unsupported {
                detail: "Worksheets::add returns Worksheet and cannot create a heterogeneous sheet type",
            });
        }
        let mut arguments = PositionalArguments::new();
        arguments.push_optional_object(options.before.map(Worksheet::dispatch_object));
        arguments.push_optional_object(options.after.map(Worksheet::dispatch_object));
        arguments.push_optional(options.count.map(|value| OwnedVariant::i32(value as i32)));
        arguments.push_optional(
            options
                .sheet_type
                .map(|value| OwnedVariant::i32(value.raw())),
        );
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
