//! Rust-owned coordinates, not ownership of a workbook or its cell contents.

use excel_api_sys::{XLOPER12, XLOPER12Value, XLREF12};

/// Invalid reference geometry or a request outside the supported operation scope.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ReferenceError {
    /// Coordinates are reversed, negative, or outside the Excel 12 grid.
    InvalidGeometry,
    /// Explicit sheet identifiers must be nonzero.
    InvalidSheet,
    /// A reference needs between one and 65,535 areas.
    InvalidAreaCount,
    /// This operation requires a single rectangle.
    MultipleAreasUnsupported,
    /// Missing is an omitted argument, not a cell value to write.
    MissingValue,
}

impl core::fmt::Display for ReferenceError {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        write!(f, "invalid Excel reference operation: {self:?}")
    }
}
impl std::error::Error for ReferenceError {}

/// Validated inclusive, zero-based rectangular coordinates in the Excel 12 grid.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ReferenceArea {
    first_row: i32,
    last_row: i32,
    first_column: i32,
    last_column: i32,
}

impl ReferenceArea {
    /// Creates a rectangle without clipping, swapping, or normalizing coordinates.
    pub fn new(
        first_row: i32,
        last_row: i32,
        first_column: i32,
        last_column: i32,
    ) -> Result<Self, ReferenceError> {
        if first_row < 0
            || last_row < first_row
            || last_row >= excel_api_sys::EXCEL12_MAX_ROWS
            || first_column < 0
            || last_column < first_column
            || last_column >= excel_api_sys::EXCEL12_MAX_COLUMNS
        {
            return Err(ReferenceError::InvalidGeometry);
        }
        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
        })
    }
    /// Creates one cell using zero-based coordinates.
    pub fn cell(row: i32, column: i32) -> Result<Self, ReferenceError> {
        Self::new(row, row, column, column)
    }
    /// First included row.
    pub const fn first_row(self) -> i32 {
        self.first_row
    }
    /// Last included row.
    pub const fn last_row(self) -> i32 {
        self.last_row
    }
    /// First included column.
    pub const fn first_column(self) -> i32 {
        self.first_column
    }
    /// Last included column.
    pub const fn last_column(self) -> i32 {
        self.last_column
    }
    /// Number of included rows.
    pub const fn rows(self) -> usize {
        (self.last_row - self.first_row + 1) as usize
    }
    /// Number of included columns.
    pub const fn columns(self) -> usize {
        (self.last_column - self.first_column + 1) as usize
    }
    pub(crate) const fn raw(self) -> XLREF12 {
        XLREF12 {
            rwFirst: self.first_row,
            rwLast: self.last_row,
            colFirst: self.first_column,
            colLast: self.last_column,
        }
    }
}

impl TryFrom<XLREF12> for ReferenceArea {
    type Error = ReferenceError;
    fn try_from(raw: XLREF12) -> Result<Self, Self::Error> {
        Self::new(raw.rwFirst, raw.rwLast, raw.colFirst, raw.colLast)
    }
}

/// Immutable coordinates owned by Rust. Does not keep a sheet open or its ID valid.
///
/// Current-sheet references are resolved by Excel at call time. Explicit IDs must
/// come from Excel (for example `active_sheet_id` or a borrowed reference); their
/// continued existence cannot be checked locally. Multiple areas share one sheet.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OwnedExcelReference {
    sheet: Option<usize>,
    areas: Box<[ReferenceArea]>,
}

impl OwnedExcelReference {
    /// Describes a rectangle on the current callback sheet, not necessarily the active sheet.
    pub fn current_sheet(area: ReferenceArea) -> Self {
        Self {
            sheet: None,
            areas: Box::new([area]),
        }
    }
    /// Describes one or more areas on a single explicit Excel sheet ID.
    pub fn on_sheet(
        sheet: usize,
        areas: impl Into<Box<[ReferenceArea]>>,
    ) -> Result<Self, ReferenceError> {
        if sheet == 0 {
            return Err(ReferenceError::InvalidSheet);
        }
        let areas = areas.into();
        if areas.is_empty() || areas.len() > u16::MAX as usize {
            return Err(ReferenceError::InvalidAreaCount);
        }
        Ok(Self {
            sheet: Some(sheet),
            areas,
        })
    }
    /// Copies geometry only; borrowed Excel backing storage is never retained.
    pub fn from_borrowed(reference: &crate::ExcelReference<'_>) -> Result<Self, ReferenceError> {
        fn area(a: &crate::ExcelReferenceArea<'_>) -> Result<ReferenceArea, ReferenceError> {
            ReferenceArea::new(
                a.first_row(),
                a.last_row(),
                a.first_column(),
                a.last_column(),
            )
        }
        match reference {
            crate::ExcelReference::Single(r) => Ok(Self::current_sheet(area(r.area())?)),
            crate::ExcelReference::Multiple(r) => Self::on_sheet(
                r.sheet_id(),
                r.areas().map(|a| area(&a)).collect::<Result<Vec<_>, _>>()?,
            ),
        }
    }
    /// Explicit sheet ID, or `None` for current-sheet coordinates.
    pub const fn sheet_id(&self) -> Option<usize> {
        self.sheet
    }
    /// Areas in supplied order; overlaps are preserved, not merged.
    pub fn areas(&self) -> &[ReferenceArea] {
        &self.areas
    }
    pub(crate) fn require_rectangle(&self) -> Result<(), ReferenceError> {
        if self.areas.len() != 1 {
            Err(ReferenceError::MultipleAreasUnsupported)
        } else {
            Ok(())
        }
    }
    pub(crate) fn argument(&self) -> ReferenceArgument {
        if let Some(sheet) = self.sheet {
            // XLMREF12: WORD count, padding to offset 4, then 16-byte XLREF12s.
            // A u32 allocation provides the required alignment and initialized padding.
            let mut storage = vec![0_u32; 1 + 4 * self.areas.len()].into_boxed_slice();
            let pointer = storage.as_mut_ptr().cast::<excel_api_sys::XLMREF12>();
            // SAFETY: allocation is aligned, large enough for count and all trailing
            // areas, uniquely owned here, and remains boxed through the synchronous call.
            unsafe {
                pointer.cast::<u16>().write(self.areas.len() as u16);
                let table = storage.as_mut_ptr().add(1).cast::<XLREF12>();
                for (i, area) in self.areas.iter().enumerate() {
                    table.add(i).write(area.raw());
                }
            }
            ReferenceArgument {
                root: XLOPER12 {
                    val: XLOPER12Value {
                        mref: excel_api_sys::XLOPER12MRef {
                            lpmref: pointer,
                            idSheet: sheet,
                        },
                    },
                    xltype: excel_api_sys::xltypeRef,
                },
                _storage: storage,
            }
        } else {
            ReferenceArgument {
                root: XLOPER12 {
                    val: XLOPER12Value {
                        sref: excel_api_sys::XLOPER12SRef {
                            count: 1,
                            reference: self.areas[0].raw(),
                        },
                    },
                    xltype: excel_api_sys::xltypeSRef,
                },
                _storage: Box::new([]),
            }
        }
    }
}

pub(crate) struct ReferenceArgument {
    pub(crate) root: XLOPER12,
    _storage: Box<[u32]>,
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn geometry_is_inclusive_and_never_clipped() {
        assert_eq!(ReferenceArea::cell(0, 0).unwrap().rows(), 1);
        let all = ReferenceArea::new(0, 1_048_575, 0, 16_383).unwrap();
        assert_eq!((all.rows(), all.columns()), (1_048_576, 16_384));
        for a in [
            (-1, 0, 0, 0),
            (2, 1, 0, 0),
            (0, 0, 1, 0),
            (0, i32::MAX, 0, 0),
            (0, 0, 0, 16_384),
        ] {
            assert_eq!(
                ReferenceArea::new(a.0, a.1, a.2, a.3),
                Err(ReferenceError::InvalidGeometry)
            );
        }
        assert!(
            ReferenceArea::try_from(XLREF12 {
                rwFirst: 0,
                rwLast: -1,
                colFirst: 0,
                colLast: 0
            })
            .is_err()
        );
    }
    #[test]
    fn sheet_and_area_counts_are_validated() {
        let a = ReferenceArea::cell(0, 0).unwrap();
        assert!(OwnedExcelReference::on_sheet(0, vec![a]).is_err());
        assert!(OwnedExcelReference::on_sheet(1, vec![]).is_err());
        assert!(OwnedExcelReference::on_sheet(1, vec![a; 65_536]).is_err());
        assert!(OwnedExcelReference::on_sheet(1, vec![a; 65_535]).is_ok());
    }
    #[test]
    fn raw_storage_roundtrips_through_existing_borrowed_decoder() {
        assert_eq!(core::mem::offset_of!(excel_api_sys::XLMREF12, reftbl), 4);
        assert_eq!(core::mem::size_of::<XLREF12>(), 16);
        assert_eq!(
            core::mem::align_of::<excel_api_sys::XLMREF12>(),
            core::mem::align_of::<u32>()
        );
        let a = ReferenceArea::cell(0, 0).unwrap();
        let b = ReferenceArea::new(8, 10, 3, 7).unwrap();
        for reference in [
            OwnedExcelReference::current_sheet(a),
            OwnedExcelReference::on_sheet(42, vec![a, b]).unwrap(),
        ] {
            let argument = reference.argument();
            // SAFETY: locally owned argument tree stays live and immutable throughout decoding.
            let decoded = unsafe { crate::RawExcelValue::from_callback(&argument.root) }
                .decode()
                .unwrap();
            let crate::ExcelValueRef::Reference(borrowed) = decoded else {
                panic!("reference expected")
            };
            assert_eq!(
                OwnedExcelReference::from_borrowed(&borrowed).unwrap(),
                reference
            );
        }
    }
}
