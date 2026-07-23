use crate::ExcelComError;
use crate::automation::{OwnedVariant, PositionalArguments, invoke, property_get};
use crate::excel::{Application, PasteOperation, PasteSpecialOptions, PasteType, Range};
use crate::object_model::{MemberId, member};

use super::helpers::finite;
use super::{AutoFillType, DataSeriesOptions};

impl Range {
    /// Fills downward through Excel, including Excel's relative-formula adjustment.
    pub fn fill_down(&self) -> Result<(), ExcelComError> {
        fill(self, "excel.range.filldown")
    }
    /// Fills upward through Excel, including Excel's relative-formula adjustment.
    pub fn fill_up(&self) -> Result<(), ExcelComError> {
        fill(self, "excel.range.fillup")
    }
    /// Fills left through Excel, including Excel's relative-formula adjustment.
    pub fn fill_left(&self) -> Result<(), ExcelComError> {
        fill(self, "excel.range.fillleft")
    }
    /// Fills right through Excel, including Excel's relative-formula adjustment.
    pub fn fill_right(&self) -> Result<(), ExcelComError> {
        fill(self, "excel.range.fillright")
    }

    /// Extends this source Range to `destination` using Excel's AutoFill rules.
    /// The destination normally includes this source Range; Excel validates the precise geometry.
    pub fn auto_fill(
        &self,
        destination: &Range,
        fill_type: Option<AutoFillType>,
    ) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_object(destination.dispatch_object());
        args.push_optional(fill_type.map(|value| OwnedVariant::i32(value.raw())));
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.autofill"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }

    /// Generates a series through Excel's `DataSeries`; finite numeric options are checked first.
    pub fn fill_series(&self, options: &DataSeriesOptions) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_optional(
            options
                .orientation
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(
            options
                .series_type
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(
            options
                .date_unit
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(finite(
            options.step_value,
            "Range.DataSeries step_value must be finite",
        )?);
        args.push_optional(finite(
            options.stop_value,
            "Range.DataSeries stop_value must be finite",
        )?);
        args.push_optional(options.trend.map(OwnedVariant::bool));
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.dataseries"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }

    /// Invokes Excel Flash Fill. Results are heuristic and version-dependent.
    pub fn flash_fill(&self) -> Result<(), ExcelComError> {
        fill(self, "excel.range.flashfill")
    }

    /// Copies through Excel's own cut/copy state, transposes with `PasteSpecial`, and clears it.
    /// This method never reads the operating-system clipboard.
    pub fn copy_transposed_to(&self, destination: &Range) -> Result<(), ExcelComError> {
        self.copy(None)?;
        let operation = destination.paste_special(&PasteSpecialOptions {
            paste: PasteType::ALL,
            operation: PasteOperation::NONE,
            skip_blanks: false,
            transpose: true,
        });
        let cleanup =
            application_for_range(self).and_then(|application| application.clear_cut_copy_mode());
        match (operation, cleanup) {
            (Ok(()), Ok(())) => Ok(()),
            (Err(error), Ok(())) | (Ok(()), Err(error)) => Err(error),
            (Err(operation), Err(_restoration)) => Err(operation),
        }
    }
}

fn fill(range: &Range, id: &'static str) -> Result<(), ExcelComError> {
    let _ = invoke(
        &range.dispatch_object().dispatch,
        member(MemberId::new(id), false),
        vec![],
        false,
    )?;
    Ok(())
}

fn application_for_range(range: &Range) -> Result<Application, ExcelComError> {
    let mut value = property_get(
        &range.dispatch_object().dispatch,
        member(MemberId::new("excel.range.application"), false),
        vec![],
    )?;
    Ok(Application::from_dispatch(value.take_dispatch()?))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_non_finite_series_values_before_com() {
        assert!(finite(Some(f64::NAN), "test").is_err());
        assert!(finite(Some(f64::INFINITY), "test").is_err());
        assert!(finite(Some(1.0), "test").is_ok());
    }
}
