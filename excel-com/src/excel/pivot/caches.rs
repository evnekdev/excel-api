//! PivotCache collection and structurally typed cache creation.

use super::helpers::{Iter, count_collection, item, source_reference};
use super::{PivotCache, PivotSourceType, PivotTableVersion};
use crate::ExcelComError;
use crate::WorkbookConnection;
use crate::automation::{OwnedVariant, PositionalArguments, invoke};
use crate::excel::collection::CollectionDescriptor;
use crate::excel::{DispatchObject, ListObject, Range, Workbook};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use std::fmt::{Debug, Formatter};

const PIVOT_CACHES: CollectionDescriptor = CollectionDescriptor {
    name: "PivotCaches",
    count: MemberId::new("excel.pivotcaches.count"),
    item: MemberId::new("excel.pivotcaches.item"),
    new_enum: MemberId::new("excel.pivotcaches.newenum"),
};

/// Apartment-bound workbook PivotCache collection.
pub struct PivotCaches {
    inner: DispatchObject,
}
impl Debug for PivotCaches {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotCaches").field(&self.inner).finish()
    }
}
impl Clone for PivotCaches {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotCaches {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotCaches",
            },
        }
    }
    /// Returns the number of workbook PivotCaches.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        count_collection(&self.inner, PIVOT_CACHES)
    }
    /// Returns the cache at a one-based index.
    pub fn item(&self, index: usize) -> Result<PivotCache, ExcelComError> {
        item(&self.inner, PIVOT_CACHES, index, PivotCache::from_dispatch)
    }
    /// Iterates caches in `_NewEnum` order on the owning apartment.
    pub fn iter(&self) -> Result<PivotCachesIter, ExcelComError> {
        Ok(PivotCachesIter {
            inner: super::helpers::iter(
                &self.inner,
                PIVOT_CACHES,
                "PivotCaches",
                PivotCache::from_dispatch,
            )?,
        })
    }
    /// Creates a database PivotCache from an Excel-qualified source Range reference.
    pub fn create_from_range(
        &self,
        source: &Range,
        version: Option<PivotTableVersion>,
    ) -> Result<PivotCache, ExcelComError> {
        self.create(
            PivotSourceType::DATABASE,
            OwnedVariant::bstr(&source_reference(source)?)?,
            version,
        )
    }
    /// Creates a database PivotCache from the full source range of a ListObject.
    pub fn create_from_table(
        &self,
        table: &ListObject,
        version: Option<PivotTableVersion>,
    ) -> Result<PivotCache, ExcelComError> {
        self.create_from_range(&table.range()?, version)
    }
    /// Creates an external PivotCache from an existing workbook connection.
    pub fn create_from_connection(
        &self,
        connection: &WorkbookConnection,
        version: Option<PivotTableVersion>,
    ) -> Result<PivotCache, ExcelComError> {
        self.create(
            PivotSourceType::EXTERNAL,
            OwnedVariant::dispatch_borrowed(&connection.dispatch_object().dispatch),
            version,
        )
    }
    fn create(
        &self,
        source_type: PivotSourceType,
        source: OwnedVariant,
        version: Option<PivotTableVersion>,
    ) -> Result<PivotCache, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(source_type.raw()));
        args.push_required(source);
        args.push_optional(version.map(|v| OwnedVariant::i32(v.raw())));
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotcaches.create"), false),
            args.into_inner(),
            false,
        )?;
        Ok(PivotCache::from_dispatch(value.take_dispatch()?))
    }
}
/// Fallible, fused, apartment-bound iterator over [`PivotCaches`].
pub struct PivotCachesIter {
    inner: Iter<PivotCache>,
}
impl Iterator for PivotCachesIter {
    type Item = Result<PivotCache, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}
impl std::iter::FusedIterator for PivotCachesIter {}
impl Workbook {
    /// Returns the workbook PivotCaches collection. Excel declares this member as a method.
    pub fn pivot_caches(&self) -> Result<PivotCaches, ExcelComError> {
        let mut value = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.workbook.pivotcaches"), false),
            vec![],
            false,
        )?;
        Ok(PivotCaches::from_dispatch(value.take_dispatch()?))
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn range_cache_uses_database_source() {
        assert_eq!(PivotSourceType::DATABASE.raw(), 1);
    }
}
