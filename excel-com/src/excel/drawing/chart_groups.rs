//! Typed chart-group access and chart-family-specific common properties.
#![allow(missing_docs)]

use super::helpers::*;
use super::*;

/// Collection of chart-type-specific groups owned by a chart.
pub struct ChartGroups {
    inner: DispatchObject,
}

impl Debug for ChartGroups {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("ChartGroups")
            .field(&self.inner)
            .finish()
    }
}

impl ChartGroups {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ChartGroups", value),
        }
    }

    /// Returns the number of Excel chart groups.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, CHART_GROUPS)
    }

    /// Returns a one-based chart group.
    pub fn item(&self, index: usize) -> Result<ChartGroup, ExcelComError> {
        collection_item(&self.inner, CHART_GROUPS, index, ChartGroup::from_dispatch)
    }

    /// Creates a fallible, single-pass chart-group iterator.
    pub fn iter(&self) -> Result<ChartGroupsIter, ExcelComError> {
        Ok(ChartGroupsIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(CHART_GROUPS.new_enum),
                CHART_GROUPS.name,
            )?,
            index: 0,
            done: false,
        })
    }
}

/// Fallible, single-pass iterator over [`ChartGroups`].
pub struct ChartGroupsIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}

impl Iterator for ChartGroupsIter {
    type Item = Result<ChartGroup, ExcelComError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, CHART_GROUPS.name, index)
                        .map(ChartGroup::from_dispatch),
                )
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

impl FusedIterator for ChartGroupsIter {}

/// Chart-family-specific group settings for an Excel chart.
///
/// Excel reports an Automation error when a member does not apply to the
/// selected chart family; the wrapper preserves that error rather than
/// emulating an inapplicable setting.
pub struct ChartGroup {
    inner: DispatchObject,
}

impl Debug for ChartGroup {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("ChartGroup")
            .field(&self.inner)
            .finish()
    }
}

impl ChartGroup {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ChartGroup", value),
        }
    }

    /// Returns the primary or secondary axis group.
    pub fn axis_group(&self) -> Result<AxisGroup, ExcelComError> {
        Ok(AxisGroup::from_raw(get_i32(
            &self.inner,
            "excel.chartgroup.axisgroup",
            "ChartGroup.AxisGroup was not an integer",
        )?))
    }

    /// Returns the applicable cluster gap width percentage.
    pub fn gap_width(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chartgroup.gapwidth",
            "ChartGroup.GapWidth was not an integer",
        )
    }

    /// Sets a bounded applicable cluster gap width percentage.
    pub fn set_gap_width(&self, value: i32) -> Result<(), ExcelComError> {
        if !(0..=500).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "ChartGroup.GapWidth must be between 0 and 500",
            });
        }
        put(
            &self.inner,
            "excel.chartgroup.gapwidth",
            OwnedVariant::i32(value),
        )
    }

    /// Returns the applicable series overlap percentage.
    pub fn overlap(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chartgroup.overlap",
            "ChartGroup.Overlap was not an integer",
        )
    }

    /// Sets a bounded applicable series overlap percentage.
    pub fn set_overlap(&self, value: i32) -> Result<(), ExcelComError> {
        if !(-100..=100).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "ChartGroup.Overlap must be between -100 and 100",
            });
        }
        put(
            &self.inner,
            "excel.chartgroup.overlap",
            OwnedVariant::i32(value),
        )
    }

    /// Returns whether Excel varies formatting by category.
    pub fn vary_by_categories(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.chartgroup.varybycategories")
    }

    /// Changes applicable category-varying formatting.
    pub fn set_vary_by_categories(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chartgroup.varybycategories",
            OwnedVariant::bool(value),
        )
    }

    /// Returns the first-slice angle for an applicable pie or doughnut group.
    pub fn first_slice_angle(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chartgroup.firstsliceangle",
            "ChartGroup.FirstSliceAngle was not an integer",
        )
    }

    /// Sets an applicable pie or doughnut first-slice angle from 0 through 360.
    pub fn set_first_slice_angle(&self, value: i32) -> Result<(), ExcelComError> {
        if !(0..=360).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "ChartGroup.FirstSliceAngle must be between 0 and 360",
            });
        }
        put(
            &self.inner,
            "excel.chartgroup.firstsliceangle",
            OwnedVariant::i32(value),
        )
    }

    /// Returns the doughnut-hole percentage for an applicable doughnut group.
    pub fn doughnut_hole_size(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chartgroup.doughnutholesize",
            "ChartGroup.DoughnutHoleSize was not an integer",
        )
    }

    /// Sets an applicable doughnut-hole percentage from 10 through 90.
    pub fn set_doughnut_hole_size(&self, value: i32) -> Result<(), ExcelComError> {
        if !(10..=90).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "ChartGroup.DoughnutHoleSize must be between 10 and 90",
            });
        }
        put(
            &self.inner,
            "excel.chartgroup.doughnutholesize",
            OwnedVariant::i32(value),
        )
    }

    /// Returns the bubble-scale percentage for an applicable bubble group.
    pub fn bubble_scale(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chartgroup.bubblescale",
            "ChartGroup.BubbleScale was not an integer",
        )
    }

    /// Sets an applicable bubble-scale percentage from 0 through 300.
    pub fn set_bubble_scale(&self, value: i32) -> Result<(), ExcelComError> {
        if !(0..=300).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "ChartGroup.BubbleScale must be between 0 and 300",
            });
        }
        put(
            &self.inner,
            "excel.chartgroup.bubblescale",
            OwnedVariant::i32(value),
        )
    }

    /// Returns whether an applicable bubble group displays negative bubbles.
    pub fn show_negative_bubbles(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.chartgroup.shownegativebubbles")
    }

    /// Shows or hides negative bubbles for an applicable bubble group.
    pub fn set_show_negative_bubbles(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chartgroup.shownegativebubbles",
            OwnedVariant::bool(value),
        )
    }
}
