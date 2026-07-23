//! Focused drawing implementation component.
#![allow(missing_docs)]
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
/// A chart's typed Series collection.
pub struct SeriesCollection {
    inner: DispatchObject,
}
impl Debug for SeriesCollection {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SeriesCollection")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for SeriesCollection {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl SeriesCollection {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("SeriesCollection", value),
        }
    }
    /// Returns the number of series in this collection.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, SERIES)
    }
    pub fn item(&self, index: usize) -> Result<Series, ExcelComError> {
        collection_item(&self.inner, SERIES, index, Series::from_dispatch)
    }
    pub fn iter(&self) -> Result<SeriesCollectionIter, ExcelComError> {
        Ok(SeriesCollectionIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(SERIES.new_enum),
                SERIES.name,
            )?,
            index: 0,
            done: false,
        })
    }
    pub fn new_series(&self) -> Result<Series, ExcelComError> {
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.seriescollection.newseries"), false),
            vec![],
            false,
        )?;
        Ok(Series::from_dispatch(value.take_dispatch()?))
    }
    pub fn add_from_range(
        &self,
        source: &Range,
        options: &SeriesAddOptions,
    ) -> Result<Series, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_object(source.dispatch_object());
        args.push_optional(options.row_col.map(|value| OwnedVariant::i32(value.raw())));
        args.push_optional(options.series_labels.map(OwnedVariant::bool));
        args.push_optional(options.category_labels.map(OwnedVariant::bool));
        args.push_optional(options.replace.map(OwnedVariant::bool));
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.seriescollection.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Series::from_dispatch(value.take_dispatch()?))
    }
}
pub struct SeriesCollectionIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for SeriesCollectionIter {
    type Item = Result<Series, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "SeriesCollection", index)
                        .map(Series::from_dispatch),
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
impl FusedIterator for SeriesCollectionIter {}

/// One Excel-native chart Series.
pub struct Series {
    inner: DispatchObject,
}
impl Debug for Series {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Series").field(&self.inner).finish()
    }
}
impl Clone for Series {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Series {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Series", value),
        }
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.series.name")
    }
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.series.name", text_bstr(value)?)
    }
    pub fn formula(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.series.formula")
    }
    pub fn set_formula(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.series.formula", text_bstr(value)?)
    }
    pub fn values(&self) -> Result<AutomationValue, ExcelComError> {
        decode_variant(
            &property_get(
                &self.inner.dispatch,
                member(MemberId::new("excel.series.values"), false),
                vec![],
            )?,
            ConversionPolicy::default(),
        )
    }
    pub fn set_values(&self, value: SeriesData<'_>) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.series.values", series_data(value)?)
    }
    pub fn x_values(&self) -> Result<AutomationValue, ExcelComError> {
        decode_variant(
            &property_get(
                &self.inner.dispatch,
                member(MemberId::new("excel.series.xvalues"), false),
                vec![],
            )?,
            ConversionPolicy::default(),
        )
    }
    pub fn set_x_values(&self, value: SeriesData<'_>) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.series.xvalues", series_data(value)?)
    }
    pub fn axis_group(&self) -> Result<AxisGroup, ExcelComError> {
        Ok(AxisGroup::from_raw(get_i32(
            &self.inner,
            "excel.series.axisgroup",
            "Series.AxisGroup was not an integer",
        )?))
    }
    pub fn set_axis_group(&self, value: AxisGroup) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.axisgroup",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn chart_type(&self) -> Result<ChartType, ExcelComError> {
        Ok(ChartType::from_raw(get_i32(
            &self.inner,
            "excel.series.charttype",
            "Series.ChartType was not an integer",
        )?))
    }
    pub fn set_chart_type(&self, value: ChartType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.charttype",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn marker_style(&self) -> Result<MarkerStyle, ExcelComError> {
        Ok(MarkerStyle::from_raw(get_i32(
            &self.inner,
            "excel.series.markerstyle",
            "Series.MarkerStyle was not an integer",
        )?))
    }
    pub fn set_marker_style(&self, value: MarkerStyle) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.markerstyle",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn marker_size(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.series.markersize",
            "Series.MarkerSize was not an integer",
        )
    }
    pub fn set_marker_size(&self, value: i32) -> Result<(), ExcelComError> {
        if !(2..=72).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "Series.MarkerSize must be between 2 and 72",
            });
        }
        put(
            &self.inner,
            "excel.series.markersize",
            OwnedVariant::i32(value),
        )
    }
    /// Returns Office drawing formatting for this series.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.series.format",
            ChartFormat::from_dispatch,
        )
    }
    /// Returns the point collection for chart families that expose individual points.
    pub fn points(&self) -> Result<Points, ExcelComError> {
        get_dispatch(&self.inner, "excel.series.points", Points::from_dispatch)
    }
    /// Returns whether Excel renders an applicable line or scatter series smoothly.
    pub fn smooth(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.series.smooth")
    }
    /// Changes smoothing for an applicable line or scatter series.
    pub fn set_smooth(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.smooth",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether Excel inverts formatting for negative values.
    pub fn invert_if_negative(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.series.invertifnegative")
    }
    /// Changes negative-value formatting inversion.
    pub fn set_invert_if_negative(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.invertifnegative",
            OwnedVariant::bool(value),
        )
    }
    /// Returns this series' one-based plot order.
    pub fn plot_order(&self) -> Result<usize, ExcelComError> {
        usize::try_from(get_i32(
            &self.inner,
            "excel.series.plotorder",
            "Series.PlotOrder was not an integer",
        )?)
        .map_err(|_| ExcelComError::Unsupported {
            detail: "Series.PlotOrder was negative",
        })
    }
    /// Sets this series' one-based plot order.
    pub fn set_plot_order(&self, value: usize) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.plotorder",
            one_based(value, "Series.PlotOrder is one-based")?,
        )
    }
    /// Returns the bubble-size source for an applicable bubble series.
    pub fn bubble_sizes(&self) -> Result<AutomationValue, ExcelComError> {
        decode_variant(
            &property_get(
                &self.inner.dispatch,
                member(MemberId::new("excel.series.bubblesizes"), false),
                vec![],
            )?,
            ConversionPolicy::default(),
        )
    }
    /// Sets the bubble-size source for an applicable bubble series.
    pub fn set_bubble_sizes(&self, values: SeriesData<'_>) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.bubblesizes",
            series_data(values)?,
        )
    }
    /// Applies the supplied legacy marker fields without touching omitted fields.
    ///
    /// Theme colours are rejected because Excel's legacy marker-colour members
    /// do not accept Office theme selectors. Use [`ChartFormat`] for Office
    /// fill and line transparency instead.
    pub fn apply_marker_format(&self, format: &MarkerFormat) -> Result<(), ExcelComError> {
        if let Some(style) = format.style {
            self.set_marker_style(style)?;
        }
        if let Some(size) = format.size {
            self.set_marker_size(size)?;
        }
        if let Some(color) = format.foreground_color {
            put(
                &self.inner,
                "excel.series.markerforegroundcolor",
                legacy_chart_color(color)?,
            )?;
        }
        if let Some(color) = format.background_color {
            put(
                &self.inner,
                "excel.series.markerbackgroundcolor",
                legacy_chart_color(color)?,
            )?;
        }
        Ok(())
    }
    /// Applies Excel's Series-level data-label options.
    pub fn apply_data_labels(&self, options: &DataLabelOptions) -> Result<(), ExcelComError> {
        if let Some(value) = &options.separator {
            let _ = text_bstr(value)?;
        }
        let mut args = PositionalArguments::new();
        args.push_optional(
            options
                .label_type
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(options.show_legend_key.map(OwnedVariant::bool));
        args.push_optional(None);
        args.push_optional(None);
        args.push_optional(options.show_series_name.map(OwnedVariant::bool));
        args.push_optional(options.show_category_name.map(OwnedVariant::bool));
        args.push_optional(options.show_value.map(OwnedVariant::bool));
        args.push_optional(options.show_percentage.map(OwnedVariant::bool));
        args.push_optional(None);
        match &options.separator {
            Some(value) => args.push_result(text_bstr(value))?,
            None => args.push_optional(None),
        };
        call(
            &self.inner,
            "excel.series.applydatalabels",
            args.into_inner(),
        )
    }
    /// Returns the Series data-label collection when Excel exposes one.
    pub fn data_labels(&self) -> Result<Option<DataLabels>, ExcelComError> {
        optional_dispatch(
            &self.inner,
            "excel.series.datalabels",
            DataLabels::from_dispatch,
        )
    }
    /// Returns the trendline collection for this series.
    pub fn trendlines(&self) -> Result<Trendlines, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.series.trendlines",
            Trendlines::from_dispatch,
        )
    }
    /// Configures Excel error bars for this series.
    pub fn set_error_bars(&self, options: &ErrorBarOptions<'_>) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(options.direction.raw()));
        args.push_required(OwnedVariant::i32(options.include.raw()));
        args.push_required(OwnedVariant::i32(options.error_type.raw()));
        args.push_optional(
            options
                .amount
                .as_ref()
                .map(|value| series_data_ref(value))
                .transpose()?,
        );
        args.push_optional(
            options
                .minus_values
                .as_ref()
                .map(|value| series_data_ref(value))
                .transpose()?,
        );
        call(&self.inner, "excel.series.errorbar", args.into_inner())
    }
    /// Removes error bars from this series through Excel's supported `ErrorBar` call.
    pub fn clear_error_bars(&self) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.series.errorbar",
            vec![
                OwnedVariant::i32(ErrorBarDirection::Y.raw()),
                OwnedVariant::i32(ErrorBarInclude::NONE.raw()),
                OwnedVariant::i32(ErrorBarType::FIXED_VALUE.raw()),
            ],
        )
    }
    /// Returns whether Excel reports error bars for this series.
    pub fn has_error_bars(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.series.haserrorbars")
    }
    /// Returns whether Excel reports leader lines for this series.
    pub fn has_leader_lines(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.series.hasleaderlines")
    }
    /// Shows or hides leader lines for this series where the chart type supports them.
    pub fn set_has_leader_lines(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.hasleaderlines",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the error-bar formatting surface when error bars are present.
    pub fn error_bars(&self) -> Result<Option<ErrorBars>, ExcelComError> {
        if !self.has_error_bars()? {
            return Ok(None);
        }
        optional_dispatch(
            &self.inner,
            "excel.series.errorbars",
            ErrorBars::from_dispatch,
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.series.delete", vec![])
    }
}
fn legacy_chart_color(color: ChartColor) -> Result<OwnedVariant, ExcelComError> {
    match color {
        ChartColor::Rgb(value) => Ok(OwnedVariant::i32(value.raw())),
        ChartColor::Automatic => Ok(OwnedVariant::i32(-4105)),
        ChartColor::Theme { .. } => Err(ExcelComError::Unsupported {
            detail: "legacy marker colours do not accept Office theme colours",
        }),
    }
}
fn series_data_ref(value: &SeriesData<'_>) -> Result<OwnedVariant, ExcelComError> {
    match value {
        SeriesData::Range(range) => Ok(OwnedVariant::dispatch_borrowed(
            &range.dispatch_object().dispatch,
        )),
        SeriesData::Array(array) => encode_variant(
            &AutomationValue::Array((*array).clone()),
            ConversionPolicy::default(),
        ),
        SeriesData::Formula(value) => text_bstr(value),
    }
}

/// A collection of data labels returned by Excel for a Series.
pub struct DataLabels {
    inner: DispatchObject,
}
impl Debug for DataLabels {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("DataLabels").field(&self.inner).finish()
    }
}
impl DataLabels {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("DataLabels", value),
        }
    }
    /// Returns the number of labels in this collection.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(
            &self.inner,
            Collection {
                name: "DataLabels",
                count: "excel.datalabels.count",
                item: "excel.datalabels.item",
                new_enum: "excel.datalabels.newenum",
            },
        )
    }
    /// Returns a one-based data label.
    pub fn item(&self, index: usize) -> Result<DataLabel, ExcelComError> {
        collection_item(
            &self.inner,
            Collection {
                name: "DataLabels",
                count: "excel.datalabels.count",
                item: "excel.datalabels.item",
                new_enum: "excel.datalabels.newenum",
            },
            index,
            DataLabel::from_dispatch,
        )
    }
    /// Creates a fallible, single-pass data-label iterator.
    pub fn iter(&self) -> Result<DataLabelsIter, ExcelComError> {
        Ok(DataLabelsIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new("excel.datalabels.newenum"),
                "DataLabels",
            )?,
            index: 0,
            done: false,
        })
    }
    /// Returns Excel's raw collection label-position value.
    pub fn position(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.datalabels.position",
            "DataLabels.Position was not an integer",
        )
    }
    /// Sets Excel's raw collection label-position value.
    pub fn set_position(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabels.position",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the separator used between collection label components.
    pub fn separator(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.datalabels.separator")
    }
    /// Sets the separator used between collection label components.
    pub fn set_separator(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.datalabels.separator", text_bstr(value)?)
    }
    /// Returns whether labels display the series name.
    pub fn show_series_name(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabels.showseriesname")
    }
    /// Shows or hides series names in labels.
    pub fn set_show_series_name(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabels.showseriesname",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether labels display category names.
    pub fn show_category_name(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabels.showcategoryname")
    }
    /// Shows or hides category names in labels.
    pub fn set_show_category_name(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabels.showcategoryname",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether labels display values.
    pub fn show_value(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabels.showvalue")
    }
    /// Shows or hides values in labels.
    pub fn set_show_value(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabels.showvalue",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether labels display percentages.
    pub fn show_percentage(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabels.showpercentage")
    }
    /// Shows or hides percentages in labels.
    pub fn set_show_percentage(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabels.showpercentage",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether labels display bubble sizes.
    pub fn show_bubble_size(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabels.showbubblesize")
    }
    /// Shows or hides bubble sizes in labels.
    pub fn set_show_bubble_size(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabels.showbubblesize",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether labels display legend keys.
    pub fn show_legend_key(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabels.showlegendkey")
    }
    /// Shows or hides legend keys in labels.
    pub fn set_show_legend_key(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabels.showlegendkey",
            OwnedVariant::bool(value),
        )
    }
    /// Returns Excel's number format for this label collection.
    pub fn number_format(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.datalabels.numberformat")
    }
    /// Sets Excel's number format for this label collection.
    pub fn set_number_format(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabels.numberformat",
            text_bstr(value)?,
        )
    }
    /// Returns the legacy Excel font for this label collection.
    pub fn font(&self) -> Result<Font, ExcelComError> {
        get_dispatch(&self.inner, "excel.datalabels.font", Font::from_dispatch)
    }
    /// Returns Office drawing formatting for this label collection.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.datalabels.format",
            ChartFormat::from_dispatch,
        )
    }
}

/// Fallible, single-pass iterator over [`DataLabels`].
pub struct DataLabelsIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for DataLabelsIter {
    type Item = Result<DataLabel, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "DataLabels", index)
                        .map(DataLabel::from_dispatch),
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
impl FusedIterator for DataLabelsIter {}
/// One Excel data label.
pub struct DataLabel {
    inner: DispatchObject,
}

/// Formatting and state of a Series error-bar collection.
pub struct ErrorBars {
    inner: DispatchObject,
}
impl Debug for ErrorBars {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("ErrorBars")
            .field(&self.inner)
            .finish()
    }
}
impl ErrorBars {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ErrorBars", value),
        }
    }
    /// Returns Office drawing formatting for the error bars.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.errorbars.format",
            ChartFormat::from_dispatch,
        )
    }
    /// Returns the legacy Excel border for the error bars.
    pub fn border(&self) -> Result<Border, ExcelComError> {
        get_dispatch(&self.inner, "excel.errorbars.border", Border::from_dispatch)
    }
}
impl Debug for DataLabel {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("DataLabel").field(&self.inner).finish()
    }
}
impl DataLabel {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("DataLabel", value),
        }
    }
    /// Returns Excel's raw position value for this label.
    pub fn position(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.datalabel.position",
            "DataLabel.Position was not an integer",
        )
    }
    /// Sets Excel's raw position value for this label.
    pub fn set_position(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabel.position",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the separator used between this label's components.
    pub fn separator(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.datalabel.separator")
    }
    /// Sets the separator used between this label's components.
    pub fn set_separator(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.datalabel.separator", text_bstr(value)?)
    }
    /// Returns whether this label displays its series name.
    pub fn show_series_name(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabel.showseriesname")
    }
    /// Shows or hides the series name in this label.
    pub fn set_show_series_name(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabel.showseriesname",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether this label displays its category name.
    pub fn show_category_name(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabel.showcategoryname")
    }
    /// Shows or hides the category name in this label.
    pub fn set_show_category_name(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabel.showcategoryname",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether this label displays its value.
    pub fn show_value(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabel.showvalue")
    }
    /// Shows or hides the value in this label.
    pub fn set_show_value(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabel.showvalue",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether this label displays a percentage.
    pub fn show_percentage(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabel.showpercentage")
    }
    /// Shows or hides the percentage in this label.
    pub fn set_show_percentage(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabel.showpercentage",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether this label displays a bubble size.
    pub fn show_bubble_size(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabel.showbubblesize")
    }
    /// Shows or hides the bubble size in this label.
    pub fn set_show_bubble_size(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabel.showbubblesize",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether this label displays a legend key.
    pub fn show_legend_key(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.datalabel.showlegendkey")
    }
    /// Shows or hides the legend key in this label.
    pub fn set_show_legend_key(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabel.showlegendkey",
            OwnedVariant::bool(value),
        )
    }
    /// Returns Excel's number format for this label.
    pub fn number_format(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.datalabel.numberformat")
    }
    /// Sets Excel's number format for this label.
    pub fn set_number_format(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.datalabel.numberformat",
            text_bstr(value)?,
        )
    }
    /// Returns the legacy Excel font for this label.
    pub fn font(&self) -> Result<Font, ExcelComError> {
        get_dispatch(&self.inner, "excel.datalabel.font", Font::from_dispatch)
    }
    /// Returns Office drawing formatting for this label.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.datalabel.format",
            ChartFormat::from_dispatch,
        )
    }
    /// Returns Excel's current text for this label.
    pub fn text(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.datalabel.text")
    }
    /// Sets this label's custom text.
    pub fn set_text(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.datalabel.text", text_bstr(value)?)
    }
}
