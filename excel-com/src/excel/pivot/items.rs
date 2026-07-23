//! PivotItem collection and manual visibility operations.

use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{OwnedVariant, invoke};
use crate::excel::DispatchObject;
use crate::excel::collection::{CollectionDescriptor, count};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

use super::helpers::{Iter, boolean, integer, method_iter, put, text};

const PIVOT_ITEMS: CollectionDescriptor = CollectionDescriptor {
    name: "PivotItems",
    count: MemberId::new("excel.pivotitems.count"),
    item: MemberId::new("excel.pivotitems.item"),
    new_enum: MemberId::new("excel.pivotitems.newenum"),
};

/// Apartment-bound collection of items in one PivotField.
pub struct PivotItems {
    inner: DispatchObject,
}
impl Debug for PivotItems {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotItems").field(&self.inner).finish()
    }
}
impl Clone for PivotItems {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotItems {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotItems",
            },
        }
    }
    /// Returns the item count.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        count(&self.inner, PIVOT_ITEMS)
    }
    /// Returns one item by one-based index.
    pub fn item_by_index(&self, index: usize) -> Result<PivotItem, ExcelComError> {
        if index == 0 {
            return Err(ExcelComError::Unsupported {
                detail: "PivotItem indexes are one-based",
            });
        }
        self.item(OwnedVariant::i32(i32::try_from(index).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "PivotItem index exceeds i32",
            }
        })?))
    }
    /// Returns an item by Excel name.
    pub fn item_by_name(&self, name: &str) -> Result<PivotItem, ExcelComError> {
        self.item(OwnedVariant::bstr(name)?)
    }
    fn item(&self, value: OwnedVariant) -> Result<PivotItem, ExcelComError> {
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotitems.item"), false),
            vec![value],
            false,
        )?;
        Ok(PivotItem::from_dispatch(value.take_dispatch()?))
    }
    /// Iterates items on the owning apartment.
    pub fn iter(&self) -> Result<PivotItemsIter, ExcelComError> {
        Ok(PivotItemsIter {
            inner: method_iter(
                &self.inner,
                "excel.pivotitems.newenum",
                "PivotItems",
                PivotItem::from_dispatch,
            )?,
        })
    }
}

/// Fallible, fused iterator over [`PivotItems`].
pub struct PivotItemsIter {
    inner: Iter<PivotItem>,
}
impl Iterator for PivotItemsIter {
    type Item = Result<PivotItem, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}
impl std::iter::FusedIterator for PivotItemsIter {}

/// One PivotField item. Excel may refuse hiding the final visible item.
pub struct PivotItem {
    inner: DispatchObject,
}
impl Debug for PivotItem {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotItem").field(&self.inner).finish()
    }
}
impl Clone for PivotItem {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotItem {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotItem",
            },
        }
    }
    /// Returns the item name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivotitem.name")
    }
    /// Returns the display caption.
    pub fn caption(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivotitem.caption")
    }
    /// Returns Excel's current visibility.
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        boolean(&self.inner, "excel.pivotitem.visible")
    }
    /// Requests item visibility; Excel enforces the last-visible-item constraint.
    pub fn set_visible(&self, visible: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivotitem.visible",
            OwnedVariant::bool(visible),
        )
    }
    /// Returns one-based display position.
    pub fn position(&self) -> Result<usize, ExcelComError> {
        usize::try_from(integer(&self.inner, "excel.pivotitem.position")?).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "PivotItem.Position was negative",
            }
        })
    }
    /// Sets a one-based item position.
    pub fn set_position(&self, position: usize) -> Result<(), ExcelComError> {
        if position == 0 {
            return Err(ExcelComError::Unsupported {
                detail: "PivotItem positions are one-based",
            });
        }
        put(
            &self.inner,
            "excel.pivotitem.position",
            OwnedVariant::i32(
                i32::try_from(position).map_err(|_| ExcelComError::Unsupported {
                    detail: "PivotItem position exceeds i32",
                })?,
            ),
        )
    }
}
