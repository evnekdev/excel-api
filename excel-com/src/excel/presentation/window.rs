//! Window, tab, and outline presentation wrappers.

use super::*;

const WINDOWS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Windows",
    count: MemberId::new("excel.windows.count"),
    item: MemberId::new("excel.windows.item"),
    new_enum: MemberId::new("excel.windows.newenum"),
};

/// A workbook or application window controlled through Excel Automation.
pub struct Window {
    inner: DispatchObject,
}
impl Debug for Window {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Window").field(&self.inner).finish()
    }
}
impl Clone for Window {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Window {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Window",
            },
        }
    }
    /// Activates this Excel window.
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.window.activate", vec![])
    }
    /// Returns the window's one-based Excel index.
    pub fn index(&self) -> Result<i32, ExcelComError> {
        get_i32(&self.inner, "excel.window.index", "Window.Index")
    }
    /// Returns whether gridlines are displayed.
    pub fn display_gridlines(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.window.displaygridlines")
    }
    /// Changes whether gridlines are displayed.
    pub fn set_display_gridlines(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.displaygridlines",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether headings are displayed.
    pub fn display_headings(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.window.displayheadings")
    }
    /// Changes whether headings are displayed.
    pub fn set_display_headings(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.displayheadings",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether zero values are displayed.
    pub fn display_zeros(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.window.displayzeros")
    }
    /// Changes whether zero values are displayed.
    pub fn set_display_zeros(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.displayzeros",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether panes are frozen.
    pub fn freeze_panes(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.window.freezepanes")
    }
    /// Changes the window's frozen-panes state.
    pub fn set_freeze_panes(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.freezepanes",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the first visible column.
    pub fn scroll_column(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.window.scrollcolumn",
            "Window.ScrollColumn",
        )
    }
    /// Sets the first visible column.
    pub fn set_scroll_column(&self, value: usize) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.scrollcolumn",
            one_based(value, "Window.ScrollColumn")?,
        )
    }
    /// Returns the first visible row.
    pub fn scroll_row(&self) -> Result<i32, ExcelComError> {
        get_i32(&self.inner, "excel.window.scrollrow", "Window.ScrollRow")
    }
    /// Sets the first visible row.
    pub fn set_scroll_row(&self, value: usize) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.scrollrow",
            one_based(value, "Window.ScrollRow")?,
        )
    }
    /// Returns the horizontal split position in points.
    pub fn split_column(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.window.splitcolumn",
            "Window.SplitColumn",
        )
    }
    /// Sets the horizontal split position in points.
    pub fn set_split_column(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value)?;
        put(
            &self.inner,
            "excel.window.splitcolumn",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the vertical split position in points.
    pub fn split_row(&self) -> Result<f64, ExcelComError> {
        get_f64(&self.inner, "excel.window.splitrow", "Window.SplitRow")
    }
    /// Sets the vertical split position in points.
    pub fn set_split_row(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value)?;
        put(
            &self.inner,
            "excel.window.splitrow",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the split column count.
    pub fn split_horizontal(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.window.splithorizontal",
            "Window.SplitHorizontal",
        )
    }
    /// Returns the split row count.
    pub fn split_vertical(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.window.splitvertical",
            "Window.SplitVertical",
        )
    }
    /// Returns the zoom percentage or Excel's automatic setting.
    pub fn zoom(&self) -> Result<PageZoom, ExcelComError> {
        page_zoom_get(&self.inner, "excel.window.zoom")
    }
    /// Sets a numeric zoom percentage or automatic zoom.
    pub fn set_zoom(&self, value: PageZoom) -> Result<(), ExcelComError> {
        page_zoom_put(&self.inner, "excel.window.zoom", value)
    }
    /// Returns the window view.
    pub fn view(&self) -> Result<WindowView, ExcelComError> {
        Ok(WindowView::from_raw(get_i32(
            &self.inner,
            "excel.window.view",
            "Window.View",
        )?))
    }
    /// Changes the window view.
    pub fn set_view(&self, value: WindowView) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.view",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Selects `cell` and freezes panes at that selection.
    pub fn freeze_at(&self, cell: &Range) -> Result<(), ExcelComError> {
        cell.select()?;
        self.set_freeze_panes(false)?;
        self.set_freeze_panes(true)
    }
}

/// Safe collection of Excel windows.
pub struct Windows {
    inner: DispatchObject,
}
impl Debug for Windows {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Windows").field(&self.inner).finish()
    }
}
impl Clone for Windows {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Windows {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Windows",
            },
        }
    }
    /// Returns the number of Excel windows.
    pub fn count(&self) -> Result<i32, ExcelComError> {
        count_i32(&self.inner, WINDOWS_DESCRIPTOR)
    }
    /// Returns a one-based window.
    pub fn item_by_index(&self, index: usize) -> Result<Window, ExcelComError> {
        Ok(Window::from_dispatch(item_by_index(
            &self.inner,
            WINDOWS_DESCRIPTOR,
            index,
        )?))
    }
}

/// The worksheet tab formatting object.
pub struct Tab {
    pub(crate) inner: DispatchObject,
}
impl Debug for Tab {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Tab").field(&self.inner).finish()
    }
}
impl Clone for Tab {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Tab {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Tab",
            },
        }
    }
    /// Returns the tab RGB color, or `Mixed`/`Empty` as returned by Excel.
    pub fn color(&self) -> Result<MixedValue<ExcelColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.tab.color", |value| {
            mixed_i32(value).map(|v| map_mixed(v, ExcelColor::from_raw))
        })
    }
    /// Sets the tab RGB color in Excel's low-byte-red COLORREF order.
    pub fn set_color(&self, value: ExcelColor) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.tab.color",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Clears the explicit tab color so Excel selects its normal tab color.
    pub fn clear_color(&self) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.tab.colorindex",
            OwnedVariant::i32(-4142),
        )
    }
}

/// Worksheet outline configuration.
pub struct Outline {
    inner: DispatchObject,
}
impl Debug for Outline {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Outline").field(&self.inner).finish()
    }
}
impl Clone for Outline {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Outline {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Outline",
            },
        }
    }
    /// Returns whether Excel applies automatic outline styles.
    pub fn automatic_styles(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.outline.automaticstyles")
    }
    /// Enables or disables Excel's automatic outline styles.
    pub fn set_automatic_styles(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.outline.automaticstyles",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the summary-row position.
    pub fn summary_row(&self) -> Result<SummaryRow, ExcelComError> {
        Ok(SummaryRow::from_raw(get_i32(
            &self.inner,
            "excel.outline.summaryrow",
            "Outline.SummaryRow",
        )?))
    }
    /// Changes the summary-row position.
    pub fn set_summary_row(&self, value: SummaryRow) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.outline.summaryrow",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the summary-column position.
    pub fn summary_column(&self) -> Result<SummaryColumn, ExcelComError> {
        Ok(SummaryColumn::from_raw(get_i32(
            &self.inner,
            "excel.outline.summarycolumn",
            "Outline.SummaryColumn",
        )?))
    }
    /// Changes the summary-column position.
    pub fn set_summary_column(&self, value: SummaryColumn) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.outline.summarycolumn",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Shows the requested row and column outline levels.
    pub fn show_levels(
        &self,
        row_level: Option<usize>,
        column_level: Option<usize>,
    ) -> Result<(), ExcelComError> {
        let mut arguments = PositionalArguments::new();
        arguments.push_optional(
            row_level
                .map(|v| one_based(v, "Outline.ShowLevels row"))
                .transpose()?,
        );
        arguments.push_optional(
            column_level
                .map(|v| one_based(v, "Outline.ShowLevels column"))
                .transpose()?,
        );
        call(
            &self.inner,
            "excel.outline.showlevels",
            arguments.into_inner(),
        )
    }
}
