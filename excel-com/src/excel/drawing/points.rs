//! Typed access to Series points and individual point formatting.
#![allow(missing_docs)]

use super::helpers::*;
use super::*;

/// One Series' collection of individual chart points.
pub struct Points {
    inner: DispatchObject,
}

impl Debug for Points {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Points").field(&self.inner).finish()
    }
}

impl Points {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Points", value),
        }
    }

    /// Returns the number of points exposed by Excel.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, POINTS)
    }

    /// Returns a one-based point.
    pub fn item(&self, index: usize) -> Result<Point, ExcelComError> {
        collection_item(&self.inner, POINTS, index, Point::from_dispatch)
    }

    /// Creates a fallible, single-pass point iterator.
    pub fn iter(&self) -> Result<PointsIter, ExcelComError> {
        Ok(PointsIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(POINTS.new_enum),
                POINTS.name,
            )?,
            index: 0,
            done: false,
        })
    }
}

/// Fallible, single-pass iterator over [`Points`].
pub struct PointsIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}

impl Iterator for PointsIter {
    type Item = Result<Point, ExcelComError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(enumerated_dispatch(&mut value, POINTS.name, index).map(Point::from_dispatch))
            }
            Ok(None) => {
                self.done = true;
                None
            }
            Err(error) => {
                self.done = true;
                Some(Err(error))
            }
        }
    }
}

impl FusedIterator for PointsIter {}

/// One formatable chart data point.
///
/// Point properties are chart-family-specific. Excel remains authoritative
/// about whether a particular property applies to the selected chart type.
pub struct Point {
    inner: DispatchObject,
}

impl Debug for Point {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Point").field(&self.inner).finish()
    }
}

impl Point {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Point", value),
        }
    }

    /// Returns whether this point has a data label.
    pub fn has_data_label(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.point.hasdatalabel")
    }

    /// Shows or hides this point's data label.
    pub fn set_has_data_label(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.point.hasdatalabel",
            OwnedVariant::bool(value),
        )
    }

    /// Returns this point's label when it is enabled.
    pub fn data_label(&self) -> Result<Option<DataLabel>, ExcelComError> {
        if !self.has_data_label()? {
            return Ok(None);
        }
        optional_dispatch(
            &self.inner,
            "excel.point.datalabel",
            DataLabel::from_dispatch,
        )
    }

    /// Returns Office drawing formatting for this point.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.point.format",
            ChartFormat::from_dispatch,
        )
    }

    /// Returns this point's legacy marker style.
    pub fn marker_style(&self) -> Result<MarkerStyle, ExcelComError> {
        Ok(MarkerStyle::from_raw(get_i32(
            &self.inner,
            "excel.point.markerstyle",
            "Point.MarkerStyle was not an integer",
        )?))
    }

    /// Sets this point's legacy marker style.
    pub fn set_marker_style(&self, value: MarkerStyle) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.point.markerstyle",
            OwnedVariant::i32(value.raw()),
        )
    }

    /// Returns this point's marker size in points.
    pub fn marker_size(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.point.markersize",
            "Point.MarkerSize was not an integer",
        )
    }

    /// Sets this point's marker size from 2 through 72 points.
    pub fn set_marker_size(&self, value: i32) -> Result<(), ExcelComError> {
        if !(2..=72).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "Point.MarkerSize must be between 2 and 72",
            });
        }
        put(
            &self.inner,
            "excel.point.markersize",
            OwnedVariant::i32(value),
        )
    }

    /// Returns the explosion percentage for an applicable pie or doughnut point.
    pub fn explosion(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.point.explosion",
            "Point.Explosion was not an integer",
        )
    }

    /// Sets an applicable point explosion percentage from 0 through 400.
    pub fn set_explosion(&self, value: i32) -> Result<(), ExcelComError> {
        if !(0..=400).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "Point.Explosion must be between 0 and 400",
            });
        }
        put(
            &self.inner,
            "excel.point.explosion",
            OwnedVariant::i32(value),
        )
    }

    /// Returns whether an applicable point inverts its formatting when negative.
    pub fn invert_if_negative(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.point.invertifnegative")
    }

    /// Changes negative-value inversion for an applicable point.
    pub fn set_invert_if_negative(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.point.invertifnegative",
            OwnedVariant::bool(value),
        )
    }
}
