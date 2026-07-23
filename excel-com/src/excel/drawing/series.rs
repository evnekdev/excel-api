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
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.series.format",
            ChartFormat::from_dispatch,
        )
    }
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
    pub fn data_labels(&self) -> Result<Option<DataLabels>, ExcelComError> {
        optional_dispatch(
            &self.inner,
            "excel.series.datalabels",
            DataLabels::from_dispatch,
        )
    }
    pub fn trendlines(&self) -> Result<Trendlines, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.series.trendlines",
            Trendlines::from_dispatch,
        )
    }
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
    pub fn has_error_bars(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.series.haserrorbars")
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.series.delete", vec![])
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
}
/// One Excel data label.
pub struct DataLabel {
    inner: DispatchObject,
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
}
