use crate::ExcelComError;
use crate::automation::{OwnedVariant, PositionalArguments, invoke};
use crate::excel::Range;
use crate::object_model::{MemberId, member};

use super::{AdvancedFilterAction, AdvancedFilterOptions};

impl Range {
    /// Runs Excel Advanced Filter. Source and criteria ranges normally include header rows.
    /// Filtering in place changes worksheet row visibility; a copy action requires `copy_to_range`.
    pub fn advanced_filter(
        &self,
        options: &AdvancedFilterOptions<'_>,
    ) -> Result<(), ExcelComError> {
        if options.action == AdvancedFilterAction::COPY && options.copy_to_range.is_none() {
            return Err(ExcelComError::Unsupported {
                detail: "AdvancedFilter copy action requires copy_to_range",
            });
        }
        if options.action != AdvancedFilterAction::COPY && options.copy_to_range.is_some() {
            return Err(ExcelComError::Unsupported {
                detail: "AdvancedFilter in-place action does not accept copy_to_range",
            });
        }
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(options.action.raw()));
        args.push_optional_object(options.criteria_range.map(Range::dispatch_object));
        args.push_optional_object(options.copy_to_range.map(Range::dispatch_object));
        args.push_optional(options.unique.map(OwnedVariant::bool));
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.advancedfilter"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }
}
