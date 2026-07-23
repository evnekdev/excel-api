//! Heterogeneous worksheet and sheet collection wrappers.

use super::*;

/// A non-worksheet object contained by the heterogeneous Excel `Sheets` collection.
pub struct SheetObject {
    inner: DispatchObject,
}

impl Debug for SheetObject {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("SheetObject")
            .field(&self.inner)
            .finish()
    }
}

impl Clone for SheetObject {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl SheetObject {
    pub(super) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "SheetObject",
            },
        }
    }

    /// Returns the sheet name supplied by Excel.
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.worksheet.name")
    }

    /// Returns the Excel sheet-type value.
    pub fn sheet_type(&self) -> Result<i32, ExcelComError> {
        get_i32(&self.inner, "excel.worksheet.type", "Sheet.Type")
    }
}

/// A safe heterogeneous member of Excel's `Sheets` collection.
#[derive(Clone, Debug)]
pub enum Sheet {
    /// A normal worksheet with the full typed worksheet API.
    Worksheet(Worksheet),
    /// A chart, macro sheet, dialog sheet, or other non-worksheet object.
    Other(SheetObject),
}

impl Sheet {
    /// Returns the sheet name supplied by Excel.
    pub fn name(&self) -> Result<String, ExcelComError> {
        match self {
            Self::Worksheet(value) => value.name(),
            Self::Other(value) => value.name(),
        }
    }

    /// Returns the numeric Excel sheet type.
    pub fn sheet_type(&self) -> Result<i32, ExcelComError> {
        match self {
            Self::Worksheet(_) => Ok(-4167),
            Self::Other(value) => value.sheet_type(),
        }
    }
}

const SHEETS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Sheets",
    count: MemberId::new("excel.sheets.count"),
    item: MemberId::new("excel.sheets.item"),
    new_enum: MemberId::new("excel.sheets.newenum"),
};

/// Safe wrapper for Excel's heterogeneous `Sheets` collection.
pub struct Sheets {
    inner: DispatchObject,
}

impl Debug for Sheets {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Sheets").field(&self.inner).finish()
    }
}
impl Clone for Sheets {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Sheets {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Sheets",
            },
        }
    }
    /// Returns the number of sheets, including non-worksheets.
    pub fn count(&self) -> Result<i32, ExcelComError> {
        count_i32(&self.inner, SHEETS_DESCRIPTOR)
    }
    /// Returns a one-based heterogeneous sheet member.
    pub fn item_by_index(&self, index: usize) -> Result<Sheet, ExcelComError> {
        sheet_from_dispatch(item_by_index(&self.inner, SHEETS_DESCRIPTOR, index)?)
    }
    /// Iterates every safe heterogeneous sheet member in Excel enum order.
    pub fn iter(&self) -> Result<SheetsIter, ExcelComError> {
        Ok(SheetsIter {
            enumerator: enumerator(&self.inner, SHEETS_DESCRIPTOR)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Single-pass iterator over [`Sheets`].
pub struct SheetsIter {
    enumerator: crate::automation::EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for SheetsIter {
    type Item = Result<Sheet, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(
                    crate::automation::enumerated_dispatch(&mut value, "Sheets", index)
                        .and_then(sheet_from_dispatch),
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
impl FusedIterator for SheetsIter {}
