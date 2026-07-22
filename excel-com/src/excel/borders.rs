//! Excel Borders collection wrapper.

use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;

use crate::ExcelComError;
use crate::automation::{EnumVariant, OwnedVariant, enumerated_dispatch, property_get};
use crate::excel::collection::{CollectionDescriptor, count as collection_count, enumerator};
use crate::excel::formatting::{
    BorderIndex, BorderLineStyle, BorderWeight, ExcelColor, ExcelColorIndex, MixedValue, map_mixed,
    mixed_i32, property_mixed_get, property_put_value,
};
use crate::excel::{Border, DispatchObject};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

const DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Borders",
    count: MemberId::new("excel.borders.count"),
    item: MemberId::new("excel.borders.item"),
    new_enum: MemberId::new("excel.borders.newenum"),
};

/// An apartment-bound Excel Borders collection returned by [`crate::Range::borders`].
///
/// `Item` accepts an Excel [`BorderIndex`] enum key, not a one-based
/// collection index. This wrapper is neither `Send` nor `Sync`.
pub struct Borders {
    inner: DispatchObject,
}

impl Debug for Borders {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Borders").field(&self.inner).finish()
    }
}

impl Clone for Borders {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl Borders {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Borders",
            },
        }
    }

    /// Returns the Excel Borders collection count.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, DESCRIPTOR)
    }

    /// Returns the border selected by its Excel [`BorderIndex`] enum key.
    pub fn item(&self, index: BorderIndex) -> Result<Border, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.borders.item"), false),
            vec![OwnedVariant::i32(index.raw())],
        )?;
        Ok(Border::from_dispatch(result.take_dispatch()?))
    }

    /// Creates a fallible, single-pass iterator in Excel's `_NewEnum` order.
    ///
    /// The order is controlled by the installed Excel runtime and is not a
    /// portable ordering guarantee.
    pub fn iter(&self) -> Result<BordersIter, ExcelComError> {
        Ok(BordersIter {
            enumerator: enumerator(&self.inner, DESCRIPTOR)?,
            next_index: 0,
            terminal: false,
        })
    }

    /// Returns the aggregate border color or a mixed result.
    pub fn color(&self) -> Result<MixedValue<ExcelColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.borders.color", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColor::from_raw))
        })
    }

    /// Sets Excel's aggregate border color.
    ///
    /// Excel controls which present borders the aggregate setter affects.
    pub fn set_color(&self, color: ExcelColor) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.borders.color",
            OwnedVariant::i32(color.raw()),
        )
    }

    /// Returns the aggregate border color index or a mixed result.
    pub fn color_index(&self) -> Result<MixedValue<ExcelColorIndex>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.borders.colorindex", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColorIndex::from_raw))
        })
    }

    /// Sets Excel's aggregate border color index.
    pub fn set_color_index(&self, index: ExcelColorIndex) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.borders.colorindex",
            OwnedVariant::i32(index.raw()),
        )
    }

    /// Returns the aggregate border line style or a mixed result.
    pub fn line_style(&self) -> Result<MixedValue<BorderLineStyle>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.borders.linestyle", |value| {
            mixed_i32(value).map(|result| map_mixed(result, BorderLineStyle::from_raw))
        })
    }

    /// Sets Excel's aggregate border line style.
    pub fn set_line_style(&self, style: BorderLineStyle) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.borders.linestyle",
            OwnedVariant::i32(style.raw()),
        )
    }

    /// Returns the aggregate border weight or a mixed result.
    pub fn weight(&self) -> Result<MixedValue<BorderWeight>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.borders.weight", |value| {
            mixed_i32(value).map(|result| map_mixed(result, BorderWeight::from_raw))
        })
    }

    /// Sets Excel's aggregate border weight.
    pub fn set_weight(&self, weight: BorderWeight) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.borders.weight",
            OwnedVariant::i32(weight.raw()),
        )
    }
}

/// A fallible, single-pass iterator over Excel Border objects.
///
/// A terminal COM or conversion error fuses the iterator, and dropping it
/// early releases the owned enumerator. It is apartment-bound and neither
/// `Send` nor `Sync`.
pub struct BordersIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}

impl Iterator for BordersIter {
    type Item = Result<Border, ExcelComError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(enumerated_dispatch(&mut value, "Borders", index).map(Border::from_dispatch))
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

impl FusedIterator for BordersIter {}
