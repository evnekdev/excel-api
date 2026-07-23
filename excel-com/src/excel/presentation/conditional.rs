//! Conditional-formatting and effective-display-format wrappers.
#![allow(missing_docs)]

use super::*;
use crate::automation::encode_variant;
use crate::excel::{
    Borders, ExcelColorIndex, Font, HorizontalAlignment, Interior, VerticalAlignment,
};

macro_rules! conditional_i32 {
    ($name:ident { $($constant:ident = $value:expr;)* }) => {
        #[repr(transparent)]
        #[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
        pub struct $name(i32);
        impl $name {
            $(pub const $constant: Self = Self($value);)*
            pub const fn from_raw(value: i32) -> Self { Self(value) }
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}

conditional_i32! { ConditionalFormatType {
    CELL_VALUE = 1; EXPRESSION = 2; COLOR_SCALE = 3; DATA_BAR = 4; TOP10 = 5;
    ICON_SET = 6; ABOVE_AVERAGE = 7; UNIQUE_VALUES = 8; TEXT_STRING = 9;
    BLANKS = 10; TIME_PERIOD = 11; NO_BLANKS = 13; ERRORS = 16; NO_ERRORS = 17;
} }
conditional_i32! { ConditionalOperator {
    BETWEEN = 1; NOT_BETWEEN = 2; EQUAL = 3; NOT_EQUAL = 4; GREATER = 5;
    LESS = 6; GREATER_OR_EQUAL = 7; LESS_OR_EQUAL = 8;
} }
conditional_i32! { ConditionValueType {
    NUMBER = 0; LOWEST_VALUE = 1; HIGHEST_VALUE = 2; PERCENT = 3; FORMULA = 4;
    PERCENTILE = 5; AUTOMATIC_MINIMUM = 6; AUTOMATIC_MAXIMUM = 7;
} }
conditional_i32! { DataBarAxisPosition { AUTOMATIC = 0; NONE = 1; CELL_MIDPOINT = 2; } }
conditional_i32! { DataBarDirection { CONTEXT = 0; LEFT_TO_RIGHT = 1; RIGHT_TO_LEFT = 2; } }
conditional_i32! { DataBarFillType { SOLID = 0; GRADIENT = 1; } }
conditional_i32! { IconSetKind {
    THREE_ARROWS = 1; THREE_ARROWS_GRAY = 2; THREE_FLAGS = 3; THREE_TRAFFIC_LIGHTS_1 = 4;
    THREE_TRAFFIC_LIGHTS_2 = 5; THREE_SIGNS = 6; THREE_SYMBOLS = 7; THREE_SYMBOLS_2 = 8;
    FOUR_ARROWS = 9; FOUR_ARROWS_GRAY = 10; FOUR_RED_TO_BLACK = 11; FOUR_RATING = 12;
    FOUR_TRAFFIC_LIGHTS = 13; FIVE_ARROWS = 14; FIVE_ARROWS_GRAY = 15; FIVE_RATING = 16;
    FIVE_QUARTERS = 17; FIVE_BOXES = 18;
} }
conditional_i32! { IconKind { NONE = 0; } }
conditional_i32! { TextConditionOperator { CONTAINS = 0; NOT_CONTAINS = 1; BEGINS_WITH = 2; ENDS_WITH = 3; } }
conditional_i32! { AboveBelowMode { ABOVE = 0; BELOW = 1; } }
conditional_i32! { DuplicateMode { DUPLICATE = 1; UNIQUE = 2; } }
conditional_i32! { TimePeriod { YESTERDAY = 1; TODAY = 2; TOMORROW = 3; LAST_7_DAYS = 4; LAST_WEEK = 5; THIS_WEEK = 6; NEXT_WEEK = 7; LAST_MONTH = 8; THIS_MONTH = 9; NEXT_MONTH = 10; } }

#[derive(Clone, Debug)]
pub struct CellValueRuleOptions<'a> {
    pub operator: ConditionalOperator,
    pub formula1: &'a str,
    pub formula2: Option<&'a str>,
}
#[derive(Clone, Debug)]
pub struct ExpressionRuleOptions<'a> {
    pub formula: &'a str,
}
/// Arguments for a text conditional-format rule.
#[derive(Clone, Debug)]
pub struct TextRuleOptions<'a> {
    /// Excel's contains/not-contains/begins/ends operator.
    pub operator: TextConditionOperator,
    /// Text Excel compares using the selected operator.
    pub text: &'a str,
}
#[derive(Clone, Debug, Default)]
pub struct Top10Options {
    pub rank: Option<usize>,
    pub percent: Option<bool>,
    pub bottom: Option<bool>,
}
/// Compatibility name for options that create a top or bottom rule.
pub type TopBottomRuleOptions = Top10Options;
#[derive(Clone, Debug, Default)]
pub struct AboveAverageOptions {
    pub above_below: Option<AboveBelowMode>,
    pub num_std_dev: Option<i32>,
}
/// Compatibility name for options that create an above/below-average rule.
pub type AboveAverageRuleOptions = AboveAverageOptions;
#[derive(Clone, Debug)]
pub struct UniqueValuesOptions {
    pub mode: DuplicateMode,
}

const FORMAT_CONDITIONS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "FormatConditions",
    count: MemberId::new("excel.formatconditions.count"),
    item: MemberId::new("excel.formatconditions.item"),
    new_enum: MemberId::new("excel.formatconditions.newenum"),
};

#[derive(Clone, Debug)]
pub struct FormatConditions {
    inner: DispatchObject,
}
impl FormatConditions {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "FormatConditions",
            },
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, FORMAT_CONDITIONS_DESCRIPTOR)
    }
    pub fn item_by_index(&self, index: usize) -> Result<ConditionalFormat, ExcelComError> {
        let mut arguments = PositionalArguments::new();
        arguments.push_required(one_based(index, "FormatConditions.Item")?);
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.formatconditions.item"), false),
            arguments.into_inner(),
            false,
        )?;
        conditional_from_dispatch(value.take_dispatch()?)
    }
    /// Returns the one-based conditional-format item at `index`.
    pub fn item(&self, index: usize) -> Result<ConditionalFormat, ExcelComError> {
        self.item_by_index(index)
    }
    pub fn iter(&self) -> Result<FormatConditionsIter, ExcelComError> {
        Ok(FormatConditionsIter {
            enumerator: enumerator(&self.inner, FORMAT_CONDITIONS_DESCRIPTOR)?,
            index: 0,
            terminal: false,
        })
    }
    pub fn delete_all(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.formatconditions.delete", vec![])
    }
    pub fn add_cell_value(
        &self,
        options: &CellValueRuleOptions<'_>,
    ) -> Result<FormatCondition, ExcelComError> {
        self.add_format_condition(
            ConditionalFormatType::CELL_VALUE,
            Some(options.operator),
            Some(options.formula1),
            options.formula2,
        )
    }
    pub fn add_expression(
        &self,
        options: &ExpressionRuleOptions<'_>,
    ) -> Result<FormatCondition, ExcelComError> {
        self.add_format_condition(
            ConditionalFormatType::EXPRESSION,
            None,
            Some(options.formula),
            None,
        )
    }
    pub fn add_color_scale(&self, criteria: usize) -> Result<ColorScale, ExcelComError> {
        if !(2..=3).contains(&criteria) {
            return Err(ExcelComError::Unsupported {
                detail: "ColorScale criteria must be two or three",
            });
        }
        let mut a = PositionalArguments::new();
        a.push_required(OwnedVariant::i32(i32::try_from(criteria).map_err(
            |_| ExcelComError::Unsupported {
                detail: "ColorScale criteria overflow",
            },
        )?));
        dispatch_result(
            &self.inner,
            "excel.formatconditions.addcolorscale",
            a.into_inner(),
            ColorScale::from_dispatch,
        )
    }
    pub fn add_data_bar(&self) -> Result<DataBar, ExcelComError> {
        dispatch_result(
            &self.inner,
            "excel.formatconditions.adddatabar",
            vec![],
            DataBar::from_dispatch,
        )
    }
    pub fn add_icon_set(&self) -> Result<IconSetCondition, ExcelComError> {
        dispatch_result(
            &self.inner,
            "excel.formatconditions.addiconsetcondition",
            vec![],
            IconSetCondition::from_dispatch,
        )
    }
    /// Creates an icon-set rule and immediately sets its icon-set family.
    pub fn add_icon_set_with_kind(
        &self,
        kind: IconSetKind,
    ) -> Result<IconSetCondition, ExcelComError> {
        let result = self.add_icon_set()?;
        result.set_icon_set(kind)?;
        Ok(result)
    }
    pub fn add_top10(&self, options: Top10Options) -> Result<Top10, ExcelComError> {
        let value = self.add_format_condition(ConditionalFormatType::TOP10, None, None, None)?;
        let result = Top10 { inner: value.inner };
        if let Some(rank) = options.rank {
            result.set_rank(rank)?;
        }
        if let Some(percent) = options.percent {
            result.set_percent(percent)?;
        }
        if let Some(bottom) = options.bottom {
            result.set_bottom(bottom)?;
        }
        Ok(result)
    }
    /// Creates a top/bottom rule using the supplied options.
    pub fn add_top_bottom(&self, options: &Top10Options) -> Result<Top10, ExcelComError> {
        self.add_top10(options.clone())
    }
    pub fn add_above_average(
        &self,
        options: &AboveAverageRuleOptions,
    ) -> Result<AboveAverage, ExcelComError> {
        let value =
            self.add_format_condition(ConditionalFormatType::ABOVE_AVERAGE, None, None, None)?;
        let result = AboveAverage { inner: value.inner };
        if let Some(mode) = options.above_below {
            result.set_above_below(mode)?;
        }
        if let Some(value) = options.num_std_dev {
            result.set_num_std_dev(value)?;
        }
        Ok(result)
    }
    /// Creates a unique or duplicate values rule.
    pub fn add_unique_or_duplicate(
        &self,
        mode: DuplicateMode,
    ) -> Result<UniqueValues, ExcelComError> {
        self.add_unique_values(UniqueValuesOptions { mode })
    }
    pub fn add_unique_values(
        &self,
        options: UniqueValuesOptions,
    ) -> Result<UniqueValues, ExcelComError> {
        let value =
            self.add_format_condition(ConditionalFormatType::UNIQUE_VALUES, None, None, None)?;
        let result = UniqueValues { inner: value.inner };
        result.set_mode(options.mode)?;
        Ok(result)
    }
    pub fn add_text_rule(
        &self,
        options: &TextRuleOptions<'_>,
    ) -> Result<TextCondition, ExcelComError> {
        let result = self.add_format_condition(
            ConditionalFormatType::TEXT_STRING,
            Some(ConditionalOperator::from_raw(options.operator.raw())),
            Some(options.text),
            None,
        )?;
        Ok(TextCondition {
            inner: result.inner,
            operator: options.operator,
        })
    }
    pub fn add_time_period(&self, period: TimePeriod) -> Result<FormatCondition, ExcelComError> {
        self.add_format_condition(
            ConditionalFormatType::TIME_PERIOD,
            Some(ConditionalOperator::from_raw(period.raw())),
            None,
            None,
        )
    }
    fn add_format_condition(
        &self,
        kind: ConditionalFormatType,
        operator: Option<ConditionalOperator>,
        formula1: Option<&str>,
        formula2: Option<&str>,
    ) -> Result<FormatCondition, ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_required(OwnedVariant::i32(kind.raw()));
        a.push_optional(operator.map(|value| OwnedVariant::i32(value.raw())));
        a.push_optional(formula1.map(text_no_nul).transpose()?);
        a.push_optional(formula2.map(text_no_nul).transpose()?);
        dispatch_result(
            &self.inner,
            "excel.formatconditions.add",
            a.into_inner(),
            FormatCondition::from_dispatch,
        )
    }
}

pub struct FormatConditionsIter {
    enumerator: crate::automation::EnumVariant,
    index: usize,
    terminal: bool,
}
impl Iterator for FormatConditionsIter {
    type Item = Result<ConditionalFormat, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    crate::automation::enumerated_dispatch(&mut value, "FormatConditions", index)
                        .and_then(conditional_from_dispatch),
                )
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
impl FusedIterator for FormatConditionsIter {}

#[derive(Clone, Debug)]
pub struct FormatCondition {
    inner: DispatchObject,
}
impl FormatCondition {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "FormatCondition",
            },
        }
    }
    pub fn condition_type(&self) -> Result<ConditionalFormatType, ExcelComError> {
        condition_type(&self.inner)
    }
    pub fn operator(&self) -> Result<ConditionalOperator, ExcelComError> {
        Ok(ConditionalOperator::from_raw(get_i32(
            &self.inner,
            "excel.formatcondition.operator",
            "FormatCondition.Operator",
        )?))
    }
    pub fn formula1(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.formatcondition.formula1")
    }
    pub fn formula2(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.formatcondition.formula2")
    }
    pub fn priority(&self) -> Result<usize, ExcelComError> {
        priority_get(
            &self.inner,
            "excel.formatcondition.priority",
            "FormatCondition.Priority",
        )
    }
    pub fn set_priority(&self, value: usize) -> Result<(), ExcelComError> {
        priority_put(&self.inner, "excel.formatcondition.priority", value)
    }
    pub fn set_first_priority(&self) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.formatcondition.setfirstpriority",
            vec![],
        )
    }
    pub fn set_last_priority(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.formatcondition.setlastpriority", vec![])
    }
    pub fn stop_if_true(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.formatcondition.stopiftrue")
    }
    pub fn set_stop_if_true(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.formatcondition.stopiftrue",
            OwnedVariant::bool(value),
        )
    }
    pub fn applies_to(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.formatcondition.appliesto")
    }
    /// Reassigns the rule's `AppliesTo` range.
    pub fn modify_applies_to_range(&self, range: &Range) -> Result<(), ExcelComError> {
        call_with_range(
            &self.inner,
            "excel.formatcondition.modifyappliestorange",
            range,
        )
    }
    pub fn font(&self) -> Result<Font, ExcelComError> {
        get_object(
            &self.inner,
            "excel.formatcondition.font",
            Font::from_dispatch,
        )
    }
    pub fn interior(&self) -> Result<Interior, ExcelComError> {
        get_object(
            &self.inner,
            "excel.formatcondition.interior",
            Interior::from_dispatch,
        )
    }
    pub fn borders(&self) -> Result<Borders, ExcelComError> {
        get_object(
            &self.inner,
            "excel.formatcondition.borders",
            Borders::from_dispatch,
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.formatcondition.delete", vec![])
    }
}

#[derive(Clone, Debug)]
pub struct ColorScale {
    inner: DispatchObject,
}
impl ColorScale {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ColorScale",
            },
        }
    }
    pub fn criteria(&self) -> Result<ColorScaleCriteria, ExcelComError> {
        get_object(
            &self.inner,
            "excel.colorscale.colorscalecriteria",
            ColorScaleCriteria::from_dispatch,
        )
    }
    pub fn priority(&self) -> Result<usize, ExcelComError> {
        priority_get(
            &self.inner,
            "excel.colorscale.priority",
            "ColorScale.Priority",
        )
    }
    pub fn set_priority(&self, value: usize) -> Result<(), ExcelComError> {
        priority_put(&self.inner, "excel.colorscale.priority", value)
    }
    pub fn applies_to(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.colorscale.appliesto")
    }
    pub fn modify_applies_to_range(&self, range: &Range) -> Result<(), ExcelComError> {
        call_with_range(&self.inner, "excel.colorscale.modifyappliestorange", range)
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.colorscale.delete", vec![])
    }
}
const COLOR_SCALE_CRITERIA_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "ColorScaleCriteria",
    count: MemberId::new("excel.colorscalecriteria.count"),
    item: MemberId::new("excel.colorscalecriteria.item"),
    new_enum: MemberId::new("excel.colorscalecriteria.newenum"),
};
#[derive(Clone, Debug)]
pub struct ColorScaleCriteria {
    inner: DispatchObject,
}
impl ColorScaleCriteria {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ColorScaleCriteria",
            },
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, COLOR_SCALE_CRITERIA_DESCRIPTOR)
    }
    pub fn item_by_index(&self, index: usize) -> Result<ColorScaleCriterion, ExcelComError> {
        get_object_item(
            &self.inner,
            "excel.colorscalecriteria.item",
            index,
            ColorScaleCriterion::from_dispatch,
        )
    }
    pub fn item(&self, index: usize) -> Result<ColorScaleCriterion, ExcelComError> {
        self.item_by_index(index)
    }
}
#[derive(Clone, Debug)]
pub struct ColorScaleCriterion {
    inner: DispatchObject,
}
impl ColorScaleCriterion {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ColorScaleCriterion",
            },
        }
    }
    pub fn value_type(&self) -> Result<ConditionValueType, ExcelComError> {
        Ok(ConditionValueType::from_raw(get_i32(
            &self.inner,
            "excel.colorscalecriterion.type",
            "ColorScaleCriterion.Type",
        )?))
    }
    pub fn criterion_type(&self) -> Result<ConditionValueType, ExcelComError> {
        self.value_type()
    }
    pub fn set_value_type(&self, value: ConditionValueType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.colorscalecriterion.type",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn set_criterion_type(&self, value: ConditionValueType) -> Result<(), ExcelComError> {
        self.set_value_type(value)
    }
    pub fn value(&self) -> Result<AutomationValue, ExcelComError> {
        get_automation_value(&self.inner, "excel.colorscalecriterion.value")
    }
    pub fn set_value(&self, value: &AutomationValue) -> Result<(), ExcelComError> {
        put_automation_value(&self.inner, "excel.colorscalecriterion.value", value)
    }
    pub fn format_color(&self) -> Result<FormatColor, ExcelComError> {
        get_object(
            &self.inner,
            "excel.colorscalecriterion.formatcolor",
            FormatColor::from_dispatch,
        )
    }
    pub fn color(&self) -> Result<FormatColor, ExcelComError> {
        self.format_color()
    }
}

#[derive(Clone, Debug)]
pub struct DataBar {
    inner: DispatchObject,
}
impl DataBar {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "DataBar",
            },
        }
    }
    pub fn min_point(&self) -> Result<ConditionValue, ExcelComError> {
        get_object(
            &self.inner,
            "excel.databar.minpoint",
            ConditionValue::from_dispatch,
        )
    }
    pub fn max_point(&self) -> Result<ConditionValue, ExcelComError> {
        get_object(
            &self.inner,
            "excel.databar.maxpoint",
            ConditionValue::from_dispatch,
        )
    }
    pub fn bar_color(&self) -> Result<FormatColor, ExcelComError> {
        get_object(
            &self.inner,
            "excel.databar.barcolor",
            FormatColor::from_dispatch,
        )
    }
    pub fn show_value(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.databar.showvalue")
    }
    pub fn set_show_value(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.databar.showvalue",
            OwnedVariant::bool(value),
        )
    }
    pub fn direction(&self) -> Result<DataBarDirection, ExcelComError> {
        Ok(DataBarDirection::from_raw(get_i32(
            &self.inner,
            "excel.databar.direction",
            "DataBar.Direction",
        )?))
    }
    pub fn set_direction(&self, value: DataBarDirection) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.databar.direction",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn axis_position(&self) -> Result<DataBarAxisPosition, ExcelComError> {
        Ok(DataBarAxisPosition::from_raw(get_i32(
            &self.inner,
            "excel.databar.axisposition",
            "DataBar.AxisPosition",
        )?))
    }
    pub fn set_axis_position(&self, value: DataBarAxisPosition) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.databar.axisposition",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn fill_type(&self) -> Result<DataBarFillType, ExcelComError> {
        Ok(DataBarFillType::from_raw(get_i32(
            &self.inner,
            "excel.databar.barfilltype",
            "DataBar.BarFillType",
        )?))
    }
    pub fn set_fill_type(&self, value: DataBarFillType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.databar.barfilltype",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn priority(&self) -> Result<usize, ExcelComError> {
        priority_get(&self.inner, "excel.databar.priority", "DataBar.Priority")
    }
    pub fn set_priority(&self, value: usize) -> Result<(), ExcelComError> {
        priority_put(&self.inner, "excel.databar.priority", value)
    }
    pub fn applies_to(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.databar.appliesto")
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.databar.delete", vec![])
    }
}
#[derive(Clone, Debug)]
pub struct ConditionValue {
    inner: DispatchObject,
}
impl ConditionValue {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ConditionValue",
            },
        }
    }
    pub fn value_type(&self) -> Result<ConditionValueType, ExcelComError> {
        Ok(ConditionValueType::from_raw(get_i32(
            &self.inner,
            "excel.conditionvalue.type",
            "ConditionValue.Type",
        )?))
    }
    pub fn set_value_type(&self, value: ConditionValueType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.conditionvalue.type",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn value(&self) -> Result<AutomationValue, ExcelComError> {
        get_automation_value(&self.inner, "excel.conditionvalue.value")
    }
    pub fn set_value(&self, value: &AutomationValue) -> Result<(), ExcelComError> {
        put_automation_value(&self.inner, "excel.conditionvalue.value", value)
    }
}

#[derive(Clone, Debug)]
pub struct IconSetCondition {
    inner: DispatchObject,
}
impl IconSetCondition {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "IconSetCondition",
            },
        }
    }
    pub fn icon_set(&self) -> Result<IconSetKind, ExcelComError> {
        Ok(IconSetKind::from_raw(get_i32(
            &self.inner,
            "excel.iconsetcondition.iconset",
            "IconSetCondition.IconSet",
        )?))
    }
    pub fn set_icon_set(&self, value: IconSetKind) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.iconsetcondition.iconset",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn reverse_order(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.iconsetcondition.reverseorder")
    }
    pub fn set_reverse_order(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.iconsetcondition.reverseorder",
            OwnedVariant::bool(value),
        )
    }
    pub fn show_icon_only(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.iconsetcondition.showicononly")
    }
    pub fn set_show_icon_only(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.iconsetcondition.showicononly",
            OwnedVariant::bool(value),
        )
    }
    pub fn criteria(&self) -> Result<IconCriteria, ExcelComError> {
        get_object(
            &self.inner,
            "excel.iconsetcondition.iconcriteria",
            IconCriteria::from_dispatch,
        )
    }
    pub fn priority(&self) -> Result<usize, ExcelComError> {
        priority_get(
            &self.inner,
            "excel.iconsetcondition.priority",
            "IconSetCondition.Priority",
        )
    }
    pub fn set_priority(&self, value: usize) -> Result<(), ExcelComError> {
        priority_put(&self.inner, "excel.iconsetcondition.priority", value)
    }
    pub fn applies_to(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.iconsetcondition.appliesto")
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.iconsetcondition.delete", vec![])
    }
}
const ICON_CRITERIA_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "IconCriteria",
    count: MemberId::new("excel.iconcriteria.count"),
    item: MemberId::new("excel.iconcriteria.item"),
    new_enum: MemberId::new("excel.iconcriteria.newenum"),
};
#[derive(Clone, Debug)]
pub struct IconCriteria {
    inner: DispatchObject,
}
impl IconCriteria {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "IconCriteria",
            },
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, ICON_CRITERIA_DESCRIPTOR)
    }
    pub fn item_by_index(&self, index: usize) -> Result<IconCriterion, ExcelComError> {
        get_object_item(
            &self.inner,
            "excel.iconcriteria.item",
            index,
            IconCriterion::from_dispatch,
        )
    }
    pub fn item(&self, index: usize) -> Result<IconCriterion, ExcelComError> {
        self.item_by_index(index)
    }
}
#[derive(Clone, Debug)]
pub struct IconCriterion {
    inner: DispatchObject,
}
impl IconCriterion {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "IconCriterion",
            },
        }
    }
    pub fn value_type(&self) -> Result<ConditionValueType, ExcelComError> {
        Ok(ConditionValueType::from_raw(get_i32(
            &self.inner,
            "excel.iconcriterion.type",
            "IconCriterion.Type",
        )?))
    }
    pub fn set_value_type(&self, value: ConditionValueType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.iconcriterion.type",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn value(&self) -> Result<AutomationValue, ExcelComError> {
        get_automation_value(&self.inner, "excel.iconcriterion.value")
    }
    pub fn set_value(&self, value: &AutomationValue) -> Result<(), ExcelComError> {
        put_automation_value(&self.inner, "excel.iconcriterion.value", value)
    }
    pub fn icon(&self) -> Result<IconKind, ExcelComError> {
        Ok(IconKind::from_raw(get_i32(
            &self.inner,
            "excel.iconcriterion.icon",
            "IconCriterion.Icon",
        )?))
    }
}

#[derive(Clone, Debug)]
pub struct Top10 {
    inner: DispatchObject,
}
impl Top10 {
    pub fn priority(&self) -> Result<usize, ExcelComError> {
        priority_get(&self.inner, "excel.top10.priority", "Top10.Priority")
    }
    pub fn set_priority(&self, value: usize) -> Result<(), ExcelComError> {
        priority_put(&self.inner, "excel.top10.priority", value)
    }
    pub fn applies_to(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.top10.appliesto")
    }
    pub fn rank(&self) -> Result<usize, ExcelComError> {
        priority_like_get(&self.inner, "excel.top10.rank", "Top10.Rank")
    }
    pub fn set_rank(&self, value: usize) -> Result<(), ExcelComError> {
        priority_like_put(&self.inner, "excel.top10.rank", value)
    }
    pub fn percent(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.top10.percent")
    }
    pub fn set_percent(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.top10.percent",
            OwnedVariant::bool(value),
        )
    }
    pub fn bottom(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.top10.topbottom")
    }
    pub fn set_bottom(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.top10.topbottom",
            OwnedVariant::bool(value),
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.top10.delete", vec![])
    }
}
#[derive(Clone, Debug)]
pub struct AboveAverage {
    inner: DispatchObject,
}
impl AboveAverage {
    pub fn priority(&self) -> Result<usize, ExcelComError> {
        priority_get(
            &self.inner,
            "excel.aboveaverage.priority",
            "AboveAverage.Priority",
        )
    }
    pub fn set_priority(&self, value: usize) -> Result<(), ExcelComError> {
        priority_put(&self.inner, "excel.aboveaverage.priority", value)
    }
    pub fn applies_to(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.aboveaverage.appliesto")
    }
    pub fn above_below(&self) -> Result<AboveBelowMode, ExcelComError> {
        Ok(AboveBelowMode::from_raw(get_i32(
            &self.inner,
            "excel.aboveaverage.abovebelow",
            "AboveAverage.AboveBelow",
        )?))
    }
    pub fn set_above_below(&self, value: AboveBelowMode) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.aboveaverage.abovebelow",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn num_std_dev(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.aboveaverage.numstddev",
            "AboveAverage.NumStdDev",
        )
    }
    pub fn set_num_std_dev(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.aboveaverage.numstddev",
            OwnedVariant::i32(value),
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.aboveaverage.delete", vec![])
    }
}
#[derive(Clone, Debug)]
pub struct UniqueValues {
    inner: DispatchObject,
}
impl UniqueValues {
    pub fn priority(&self) -> Result<usize, ExcelComError> {
        priority_get(
            &self.inner,
            "excel.uniquevalues.priority",
            "UniqueValues.Priority",
        )
    }
    pub fn set_priority(&self, value: usize) -> Result<(), ExcelComError> {
        priority_put(&self.inner, "excel.uniquevalues.priority", value)
    }
    pub fn applies_to(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.uniquevalues.appliesto")
    }
    pub fn mode(&self) -> Result<DuplicateMode, ExcelComError> {
        Ok(DuplicateMode::from_raw(get_i32(
            &self.inner,
            "excel.uniquevalues.dupeunique",
            "UniqueValues.DupeUnique",
        )?))
    }
    pub fn set_mode(&self, value: DuplicateMode) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.uniquevalues.dupeunique",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.uniquevalues.delete", vec![])
    }
}
#[derive(Clone, Debug)]
pub struct TextCondition {
    inner: DispatchObject,
    operator: TextConditionOperator,
}
impl TextCondition {
    pub fn operator(&self) -> TextConditionOperator {
        self.operator
    }
    pub fn text(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.formatcondition.formula1")
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.formatcondition.delete", vec![])
    }
}
#[derive(Clone, Debug)]
pub struct UnsupportedConditionalFormat {
    inner: DispatchObject,
    format_type: ConditionalFormatType,
}
impl UnsupportedConditionalFormat {
    pub fn format_type(&self) -> ConditionalFormatType {
        self.format_type
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.formatcondition.delete", vec![])
    }
}

#[derive(Clone, Debug)]
pub enum ConditionalFormat {
    CellValue(FormatCondition),
    Expression(FormatCondition),
    ColorScale(ColorScale),
    DataBar(DataBar),
    IconSet(IconSetCondition),
    Top10(Top10),
    AboveAverage(AboveAverage),
    UniqueValues(UniqueValues),
    Text(TextCondition),
    Unsupported(UnsupportedConditionalFormat),
}
impl ConditionalFormat {
    pub fn format_type(&self) -> Result<ConditionalFormatType, ExcelComError> {
        Ok(match self {
            Self::CellValue(v) | Self::Expression(v) => v.condition_type()?,
            Self::ColorScale(_) => ConditionalFormatType::COLOR_SCALE,
            Self::DataBar(_) => ConditionalFormatType::DATA_BAR,
            Self::IconSet(_) => ConditionalFormatType::ICON_SET,
            Self::Top10(_) => ConditionalFormatType::TOP10,
            Self::AboveAverage(_) => ConditionalFormatType::ABOVE_AVERAGE,
            Self::UniqueValues(_) => ConditionalFormatType::UNIQUE_VALUES,
            Self::Text(_) => ConditionalFormatType::TEXT_STRING,
            Self::Unsupported(v) => v.format_type,
        })
    }
}

#[derive(Clone, Debug)]
pub struct FormatColor {
    inner: DispatchObject,
}
impl FormatColor {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "FormatColor",
            },
        }
    }
    pub fn color(&self) -> Result<ExcelColor, ExcelComError> {
        Ok(ExcelColor::from_raw(get_i32(
            &self.inner,
            "excel.formatcolor.color",
            "FormatColor.Color",
        )?))
    }
    pub fn set_color(&self, value: ExcelColor) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.formatcolor.color",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn color_index(&self) -> Result<ExcelColorIndex, ExcelComError> {
        Ok(ExcelColorIndex::from_raw(get_i32(
            &self.inner,
            "excel.formatcolor.colorindex",
            "FormatColor.ColorIndex",
        )?))
    }
    pub fn set_color_index(&self, value: ExcelColorIndex) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.formatcolor.colorindex",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn theme_color(&self) -> Result<ThemeColor, ExcelComError> {
        Ok(ThemeColor::from_raw(get_i32(
            &self.inner,
            "excel.formatcolor.themecolor",
            "FormatColor.ThemeColor",
        )?))
    }
    pub fn set_theme_color(&self, value: ThemeColor) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.formatcolor.themecolor",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn tint_and_shade(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.formatcolor.tintandshade",
            "FormatColor.TintAndShade",
        )
    }
    pub fn set_tint_and_shade(&self, value: f64) -> Result<(), ExcelComError> {
        tint_put(&self.inner, "excel.formatcolor.tintandshade", value)
    }
}

#[derive(Clone, Debug)]
pub struct DisplayFormat {
    inner: DispatchObject,
}
impl DisplayFormat {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "DisplayFormat",
            },
        }
    }
    pub fn font(&self) -> Result<Font, ExcelComError> {
        get_object(&self.inner, "excel.displayformat.font", Font::from_dispatch)
    }
    pub fn interior(&self) -> Result<Interior, ExcelComError> {
        get_object(
            &self.inner,
            "excel.displayformat.interior",
            Interior::from_dispatch,
        )
    }
    pub fn borders(&self) -> Result<Borders, ExcelComError> {
        get_object(
            &self.inner,
            "excel.displayformat.borders",
            Borders::from_dispatch,
        )
    }
    pub fn number_format(&self) -> Result<MixedValue<String>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.displayformat.numberformat", |value| {
            value.as_string().map(MixedValue::Uniform)
        })
    }
    pub fn horizontal_alignment(&self) -> Result<MixedValue<HorizontalAlignment>, ExcelComError> {
        property_mixed_get(
            &self.inner,
            "excel.displayformat.horizontalalignment",
            |value| mixed_i32(value).map(|mixed| map_mixed(mixed, HorizontalAlignment::from_raw)),
        )
    }
    pub fn vertical_alignment(&self) -> Result<MixedValue<VerticalAlignment>, ExcelComError> {
        property_mixed_get(
            &self.inner,
            "excel.displayformat.verticalalignment",
            |value| mixed_i32(value).map(|mixed| map_mixed(mixed, VerticalAlignment::from_raw)),
        )
    }
    pub fn wrap_text(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.displayformat.wraptext", mixed_bool)
    }
}

impl Range {
    pub fn format_conditions(&self) -> Result<FormatConditions, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.range.formatconditions",
            FormatConditions::from_dispatch,
        )
    }
    pub fn display_format(&self) -> Result<DisplayFormat, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.range.displayformat",
            DisplayFormat::from_dispatch,
        )
    }
}

fn conditional_from_dispatch(
    dispatch: ComPtr<Dispatch>,
) -> Result<ConditionalFormat, ExcelComError> {
    let inner = DispatchObject {
        dispatch,
        kind: "ConditionalFormat",
    };
    match condition_type(&inner)? {
        value if value == ConditionalFormatType::CELL_VALUE => {
            Ok(ConditionalFormat::CellValue(FormatCondition { inner }))
        }
        value if value == ConditionalFormatType::EXPRESSION => {
            Ok(ConditionalFormat::Expression(FormatCondition { inner }))
        }
        value if value == ConditionalFormatType::COLOR_SCALE => {
            Ok(ConditionalFormat::ColorScale(ColorScale { inner }))
        }
        value if value == ConditionalFormatType::DATA_BAR => {
            Ok(ConditionalFormat::DataBar(DataBar { inner }))
        }
        value if value == ConditionalFormatType::ICON_SET => {
            Ok(ConditionalFormat::IconSet(IconSetCondition { inner }))
        }
        value if value == ConditionalFormatType::TOP10 => {
            Ok(ConditionalFormat::Top10(Top10 { inner }))
        }
        value if value == ConditionalFormatType::ABOVE_AVERAGE => {
            Ok(ConditionalFormat::AboveAverage(AboveAverage { inner }))
        }
        value if value == ConditionalFormatType::UNIQUE_VALUES => {
            Ok(ConditionalFormat::UniqueValues(UniqueValues { inner }))
        }
        value if value == ConditionalFormatType::TEXT_STRING => {
            Ok(ConditionalFormat::Text(TextCondition {
                inner,
                operator: TextConditionOperator::CONTAINS,
            }))
        }
        value => Ok(ConditionalFormat::Unsupported(
            UnsupportedConditionalFormat {
                inner,
                format_type: value,
            },
        )),
    }
}
fn condition_type(target: &DispatchObject) -> Result<ConditionalFormatType, ExcelComError> {
    Ok(ConditionalFormatType::from_raw(get_i32(
        target,
        "excel.formatcondition.type",
        "FormatCondition.Type",
    )?))
}
fn dispatch_result<T>(
    target: &DispatchObject,
    id: &'static str,
    args: Vec<OwnedVariant>,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut result = invoke(
        &target.dispatch,
        member(MemberId::new(id), false),
        args,
        false,
    )?;
    Ok(from(result.take_dispatch()?))
}
fn get_object_item<T>(
    target: &DispatchObject,
    id: &'static str,
    index: usize,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut args = PositionalArguments::new();
    args.push_required(one_based(index, "collection index")?);
    let mut result = property_get(
        &target.dispatch,
        member(MemberId::new(id), false),
        args.into_inner(),
    )?;
    Ok(from(result.take_dispatch()?))
}
fn priority_get(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<usize, ExcelComError> {
    priority_like_get(target, id, detail)
}
fn priority_put(
    target: &DispatchObject,
    id: &'static str,
    value: usize,
) -> Result<(), ExcelComError> {
    priority_like_put(target, id, value)
}
fn priority_like_get(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<usize, ExcelComError> {
    usize::try_from(get_i32(target, id, detail)?)
        .ok()
        .filter(|v| *v > 0)
        .ok_or(ExcelComError::Unsupported { detail })
}
fn priority_like_put(
    target: &DispatchObject,
    id: &'static str,
    value: usize,
) -> Result<(), ExcelComError> {
    put(target, id, one_based(value, "priority")?)
}
fn call_with_range(
    target: &DispatchObject,
    id: &'static str,
    range: &Range,
) -> Result<(), ExcelComError> {
    let mut args = PositionalArguments::new();
    args.push_object(range.dispatch_object());
    call(target, id, args.into_inner())
}
fn text_no_nul(value: &str) -> Result<OwnedVariant, ExcelComError> {
    if value.contains('\0') {
        return Err(ExcelComError::Unsupported {
            detail: "Excel formula text cannot contain NUL",
        });
    }
    text_bstr(value)
}
fn get_automation_value(
    target: &DispatchObject,
    id: &'static str,
) -> Result<AutomationValue, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    decode_variant(&value, ConversionPolicy::default())
}
fn put_automation_value(
    target: &DispatchObject,
    id: &'static str,
    value: &AutomationValue,
) -> Result<(), ExcelComError> {
    let value = encode_variant(value, ConversionPolicy::default())?;
    put(target, id, value)
}
fn tint_put(target: &DispatchObject, id: &'static str, value: f64) -> Result<(), ExcelComError> {
    finite(value)?;
    if !(-1.0..=1.0).contains(&value) {
        return Err(ExcelComError::Unsupported {
            detail: "TintAndShade must be between -1.0 and 1.0",
        });
    }
    put(target, id, OwnedVariant::f64(value))
}
