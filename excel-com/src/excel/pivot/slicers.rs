//! Read-only slicer-cache inspection where installed Excel exposes it.

use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::excel::collection::{CollectionDescriptor, count, item_by_index, item_by_name};
use crate::excel::{DispatchObject, Workbook};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::MemberId;

use super::PivotTables;
use super::helpers::{Iter, object, text};

const SLICER_CACHES: CollectionDescriptor = CollectionDescriptor {
    name: "SlicerCaches",
    count: MemberId::new("excel.slicercaches.count"),
    item: MemberId::new("excel.slicercaches.item"),
    new_enum: MemberId::new("excel.slicercaches.newenum"),
};
const SLICERS: CollectionDescriptor = CollectionDescriptor {
    name: "Slicers",
    count: MemberId::new("excel.slicers.count"),
    item: MemberId::new("excel.slicers.item"),
    new_enum: MemberId::new("excel.slicers.newenum"),
};

/// Apartment-bound collection of workbook slicer caches.
pub struct SlicerCaches {
    inner: DispatchObject,
}
impl Debug for SlicerCaches {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SlicerCaches").field(&self.inner).finish()
    }
}
impl Clone for SlicerCaches {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl SlicerCaches {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "SlicerCaches",
            },
        }
    }
    /// Returns slicer-cache count.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        count(&self.inner, SLICER_CACHES)
    }
    /// Returns a cache by one-based index.
    pub fn item_by_index(&self, index: usize) -> Result<SlicerCache, ExcelComError> {
        Ok(SlicerCache::from_dispatch(item_by_index(
            &self.inner,
            SLICER_CACHES,
            index,
        )?))
    }
    /// Returns a cache by Excel name.
    pub fn item_by_name(&self, name: &str) -> Result<SlicerCache, ExcelComError> {
        Ok(SlicerCache::from_dispatch(item_by_name(
            &self.inner,
            SLICER_CACHES,
            name,
        )?))
    }
    /// Iterates caches in `_NewEnum` order.
    pub fn iter(&self) -> Result<SlicerCachesIter, ExcelComError> {
        Ok(SlicerCachesIter {
            inner: super::helpers::iter(
                &self.inner,
                SLICER_CACHES,
                "SlicerCaches",
                SlicerCache::from_dispatch,
            )?,
        })
    }
}
/// Fallible, fused iterator over [`SlicerCaches`].
pub struct SlicerCachesIter {
    inner: Iter<SlicerCache>,
}
impl Iterator for SlicerCachesIter {
    type Item = Result<SlicerCache, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}
impl std::iter::FusedIterator for SlicerCachesIter {}

/// A slicer cache with its source and connected reports.
pub struct SlicerCache {
    inner: DispatchObject,
}
impl Debug for SlicerCache {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SlicerCache").field(&self.inner).finish()
    }
}
impl Clone for SlicerCache {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl SlicerCache {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "SlicerCache",
            },
        }
    }
    /// Returns the cache name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.slicercache.name")
    }
    /// Returns the source field name.
    pub fn source_name(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.slicercache.sourcename")
    }
    /// Returns PivotTables connected to the cache.
    pub fn pivot_tables(&self) -> Result<PivotTables, ExcelComError> {
        object(
            &self.inner,
            "excel.slicercache.pivottables",
            PivotTables::from_dispatch,
        )
    }
    /// Returns slicers using this cache.
    pub fn slicers(&self) -> Result<Slicers, ExcelComError> {
        object(
            &self.inner,
            "excel.slicercache.slicers",
            Slicers::from_dispatch,
        )
    }
}

/// Apartment-bound collection of slicer UI objects.
pub struct Slicers {
    inner: DispatchObject,
}
impl Debug for Slicers {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Slicers").field(&self.inner).finish()
    }
}
impl Clone for Slicers {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Slicers {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Slicers",
            },
        }
    }
    /// Returns slicer count.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        count(&self.inner, SLICERS)
    }
    /// Returns one slicer by one-based index.
    pub fn item_by_index(&self, index: usize) -> Result<Slicer, ExcelComError> {
        Ok(Slicer::from_dispatch(item_by_index(
            &self.inner,
            SLICERS,
            index,
        )?))
    }
    /// Iterates slicers in `_NewEnum` order.
    pub fn iter(&self) -> Result<SlicersIter, ExcelComError> {
        Ok(SlicersIter {
            inner: super::helpers::iter(&self.inner, SLICERS, "Slicers", Slicer::from_dispatch)?,
        })
    }
}
/// Fallible, fused iterator over [`Slicers`].
pub struct SlicersIter {
    inner: Iter<Slicer>,
}
impl Iterator for SlicersIter {
    type Item = Result<Slicer, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}
impl std::iter::FusedIterator for SlicersIter {}

/// One slicer UI object; creation is intentionally version-dependent and unavailable.
pub struct Slicer {
    inner: DispatchObject,
}
impl Debug for Slicer {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Slicer").field(&self.inner).finish()
    }
}
impl Clone for Slicer {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Slicer {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Slicer",
            },
        }
    }
    /// Returns slicer name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.slicer.name")
    }
    /// Returns the cache that supplies this slicer.
    pub fn slicer_cache(&self) -> Result<SlicerCache, ExcelComError> {
        object(
            &self.inner,
            "excel.slicer.slicercache",
            SlicerCache::from_dispatch,
        )
    }
}

impl Workbook {
    /// Returns the workbook slicer-cache collection when supported by the installed Excel version.
    pub fn slicer_caches(&self) -> Result<SlicerCaches, ExcelComError> {
        object(
            self.dispatch_object(),
            "excel.workbook.slicercaches",
            SlicerCaches::from_dispatch,
        )
    }
}
