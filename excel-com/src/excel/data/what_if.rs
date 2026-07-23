use crate::ExcelComError;
use crate::automation::{OwnedVariant, PositionalArguments, invoke};
use crate::excel::Range;
use crate::object_model::{MemberId, member};

use super::helpers::{one_based, one_dimensional_i32};
use super::{
    ConsolidateOptions, ConsolidationSource, DataTableInputs, GoalSeekOptions, SubtotalOptions,
};

impl Range {
    /// Adds Excel subtotal rows. Data should normally be sorted by the group column first.
    pub fn subtotal(&self, options: &SubtotalOptions) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(one_based(
            options.group_by,
            "Subtotal group_by is one-based",
        )?));
        args.push_required(OwnedVariant::i32(options.function.raw()));
        args.push_required(one_dimensional_i32(
            &options.total_columns,
            "Subtotal total_columns must be non-empty and one-based",
        )?);
        args.push_optional(options.replace.map(OwnedVariant::bool));
        args.push_optional(options.page_breaks.map(OwnedVariant::bool));
        args.push_optional(options.summary_below_data.map(OwnedVariant::bool));
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.subtotal"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }
    /// Removes subtotal rows created by Excel from this Range.
    pub fn remove_subtotals(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.removesubtotal"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Consolidates controlled local sources through Excel. `create_links` can create external-reference formulas.
    pub fn consolidate(&self, options: &ConsolidateOptions<'_>) -> Result<(), ExcelComError> {
        if options.sources.is_empty() {
            return Err(ExcelComError::Unsupported {
                detail: "Range.Consolidate requires at least one source",
            });
        }
        let mut source_values = Vec::with_capacity(options.sources.len());
        for source in &options.sources {
            let reference = match source {
                ConsolidationSource::Range(range) => {
                    range.external_address(crate::ReferenceStyle::A1)?
                }
                ConsolidationSource::Reference(reference) => {
                    if reference.contains('\0') {
                        return Err(ExcelComError::Unsupported {
                            detail: "Consolidation reference cannot contain embedded NUL",
                        });
                    }
                    (*reference).to_owned()
                }
            };
            source_values.push(reference);
        }
        let source_array = crate::AutomationArray::row(
            source_values
                .into_iter()
                .map(crate::AutomationValue::Text)
                .collect(),
        )?;
        let source_array = crate::automation::encode_variant(
            &crate::AutomationValue::Array(source_array),
            crate::ConversionPolicy::default(),
        )?;
        let mut args = PositionalArguments::new();
        args.push_required(source_array);
        args.push_optional(Some(OwnedVariant::i32(options.function.raw())));
        args.push_optional(options.top_row_labels.map(OwnedVariant::bool));
        args.push_optional(options.left_column_labels.map(OwnedVariant::bool));
        args.push_optional(options.create_links.map(OwnedVariant::bool));
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.consolidate"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }
    /// Invokes Excel's numerical Goal Seek solver and returns Excel's Boolean success result.
    pub fn goal_seek(&self, options: &GoalSeekOptions<'_>) -> Result<bool, ExcelComError> {
        if !options.goal.is_finite() {
            return Err(ExcelComError::Unsupported {
                detail: "GoalSeek goal must be finite",
            });
        }
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::f64(options.goal));
        args.push_object(options.changing_cell.dispatch_object());
        let result = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.goalseek"), false),
            args.into_inner(),
            false,
        )?;
        result.as_bool().ok_or(ExcelComError::Unsupported {
            detail: "Range.GoalSeek did not return VT_BOOL",
        })
    }
    /// Creates Excel's what-if Data Table. This is not a `ListObject` table and can be expensive.
    pub fn create_data_table(&self, inputs: &DataTableInputs<'_>) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        match inputs {
            DataTableInputs::Row { row_input } => {
                args.push_object(row_input.dispatch_object());
                args.push_optional(None);
            }
            DataTableInputs::Column { column_input } => {
                args.push_optional(None);
                args.push_object(column_input.dispatch_object());
            }
            DataTableInputs::TwoVariable {
                row_input,
                column_input,
            } => {
                args.push_object(row_input.dispatch_object());
                args.push_object(column_input.dispatch_object());
            }
        }
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.table"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }
    /// Returns a potentially multi-area Range of cells differing by row from Excel's comparison cell.
    pub fn row_differences(&self, comparison: &Range) -> Result<Range, ExcelComError> {
        differences(self, comparison, "excel.range.rowdifferences")
    }
    /// Returns a potentially multi-area Range of cells differing by column from Excel's comparison cell.
    pub fn column_differences(&self, comparison: &Range) -> Result<Range, ExcelComError> {
        differences(self, comparison, "excel.range.columndifferences")
    }
}

fn differences(
    range: &Range,
    comparison: &Range,
    id: &'static str,
) -> Result<Range, ExcelComError> {
    let mut args = PositionalArguments::new();
    args.push_object(comparison.dispatch_object());
    let mut value = invoke(
        &range.dispatch_object().dispatch,
        member(MemberId::new(id), false),
        args.into_inner(),
        false,
    )?;
    Ok(Range::from_dispatch(value.take_dispatch()?))
}
