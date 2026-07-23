//! Worksheet PivotTable collection access.

use super::PivotTable;
use super::helpers::{Iter, method_iter};
use crate::ExcelComError;
use crate::automation::{OwnedVariant, invoke};
use crate::excel::collection::{CollectionDescriptor, count};
use crate::excel::{DispatchObject, Worksheet};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use std::fmt::{Debug, Formatter};

const PIVOT_TABLES: CollectionDescriptor = CollectionDescriptor {
    name: "PivotTables",
    count: MemberId::new("excel.pivottables.count"),
    item: MemberId::new("excel.pivottables.item"),
    new_enum: MemberId::new("excel.pivottables.newenum"),
};
/// Apartment-bound collection of PivotTables on a worksheet.
pub struct PivotTables {
    inner: DispatchObject,
}
impl Debug for PivotTables {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotTables").field(&self.inner).finish()
    }
}
impl Clone for PivotTables {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotTables {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotTables",
            },
        }
    }
    /// Returns the PivotTable count.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        count(&self.inner, PIVOT_TABLES)
    }
    /// Returns a PivotTable by a one-based index. Excel declares `Item` as a method.
    pub fn item_by_index(&self, index: usize) -> Result<PivotTable, ExcelComError> {
        if index == 0 {
            return Err(ExcelComError::Unsupported {
                detail: "PivotTable indexes are one-based",
            });
        }
        self.item(OwnedVariant::i32(i32::try_from(index).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "PivotTable index exceeds i32",
            }
        })?))
    }
    /// Returns a PivotTable by its name. Excel declares `Item` as a method.
    pub fn item_by_name(&self, name: &str) -> Result<PivotTable, ExcelComError> {
        self.item(OwnedVariant::bstr(name)?)
    }
    fn item(&self, value: OwnedVariant) -> Result<PivotTable, ExcelComError> {
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivottables.item"), false),
            vec![value],
            false,
        )?;
        Ok(PivotTable::from_dispatch(value.take_dispatch()?))
    }
    /// Iterates PivotTables in Excel's method-based `_NewEnum` order.
    pub fn iter(&self) -> Result<PivotTablesIter, ExcelComError> {
        Ok(PivotTablesIter {
            inner: method_iter(
                &self.inner,
                "excel.pivottables.newenum",
                "PivotTables",
                PivotTable::from_dispatch,
            )?,
        })
    }
}
/// Fallible, fused, apartment-bound iterator over [`PivotTables`].
pub struct PivotTablesIter {
    inner: Iter<PivotTable>,
}
impl Iterator for PivotTablesIter {
    type Item = Result<PivotTable, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}
impl std::iter::FusedIterator for PivotTablesIter {}
impl Worksheet {
    /// Returns this sheet's PivotTables collection. Excel declares the member as a method with an optional index.
    pub fn pivot_tables(&self) -> Result<PivotTables, ExcelComError> {
        let mut value = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.worksheet.pivottables"), false),
            vec![],
            false,
        )?;
        Ok(PivotTables::from_dispatch(value.take_dispatch()?))
    }
}
