use crate::automation::{
    ConversionPolicy, OwnedVariant, PositionalArguments, activate_excel, decode_variant, invoke,
    property_get, property_put,
};
use crate::excel::text::text_bstr;
use crate::excel::{
    CalculationMode, CalculationState, DispatchObject, FormulaConversionOptions, Range,
    ReferenceStyle, Workbooks,
};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use crate::{ComApartment, ConversionError, ExcelComError};
use std::fmt::{Debug, Formatter};

/// Restores an [`Application`]'s prior `DisplayAlerts` value on drop.
///
/// Restoration failures cannot be returned from `Drop`; call [`Self::restore`]
/// when the caller needs to observe the result directly.
pub struct DisplayAlertsGuard<'a> {
    application: &'a Application,
    previous: bool,
    active: bool,
}

/// Restores an [`Application`]'s prior global `ReferenceStyle` on drop.
///
/// The setting is process-wide Excel state. Prefer this guard to an unscoped
/// mutation, and call [`Self::restore`] when restoration errors must be
/// observed directly.
pub struct ReferenceStyleGuard<'a> {
    application: &'a Application,
    previous: ReferenceStyle,
    active: bool,
}

/// Restores an [`Application`]'s prior global calculation mode on drop.
///
/// Calculation mode affects the entire Excel Application, not only the
/// workbook or Range that created this guard. Call [`Self::restore`] when a
/// restoration failure must be observed directly.
pub struct CalculationModeGuard<'a> {
    application: &'a Application,
    previous: CalculationMode,
    active: bool,
}

impl Debug for ReferenceStyleGuard<'_> {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ReferenceStyleGuard")
            .field("previous", &self.previous)
            .field("active", &self.active)
            .finish()
    }
}

impl ReferenceStyleGuard<'_> {
    /// Restores the prior reference style and disarms the guard.
    pub fn restore(mut self) -> Result<(), ExcelComError> {
        self.application.set_reference_style(self.previous)?;
        self.active = false;
        Ok(())
    }
}

impl Drop for ReferenceStyleGuard<'_> {
    fn drop(&mut self) {
        if self.active {
            let _ = self.application.set_reference_style(self.previous);
        }
    }
}

impl Debug for CalculationModeGuard<'_> {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CalculationModeGuard")
            .field("previous", &self.previous)
            .field("active", &self.active)
            .finish()
    }
}

impl CalculationModeGuard<'_> {
    /// Restores the prior calculation mode and disarms the guard.
    pub fn restore(mut self) -> Result<(), ExcelComError> {
        self.application.set_calculation_mode(self.previous)?;
        self.active = false;
        Ok(())
    }
}

impl Drop for CalculationModeGuard<'_> {
    fn drop(&mut self) {
        if self.active {
            let _ = self.application.set_calculation_mode(self.previous);
            self.active = false;
        }
    }
}

impl Debug for DisplayAlertsGuard<'_> {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("DisplayAlertsGuard")
            .field("previous", &self.previous)
            .field("active", &self.active)
            .finish()
    }
}

impl DisplayAlertsGuard<'_> {
    /// Restores the prior `DisplayAlerts` value and disarms the guard.
    pub fn restore(mut self) -> Result<(), ExcelComError> {
        self.application.set_display_alerts(self.previous)?;
        self.active = false;
        Ok(())
    }
}

impl Drop for DisplayAlertsGuard<'_> {
    fn drop(&mut self) {
        if self.active {
            let _ = self.application.set_display_alerts(self.previous);
        }
    }
}

/// Experimental wrapper for a crate-created local Excel Application instance.
pub struct Application {
    inner: DispatchObject,
}
impl Debug for Application {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Application")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for Application {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Application {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Application",
            },
        }
    }

    /// Starts a fresh local Excel `Application` in the supplied STA apartment.
    pub fn new(apartment: &ComApartment) -> Result<Self, ExcelComError> {
        apartment.assert_current()?;
        Ok(Self::from_dispatch(activate_excel()?))
    }
    /// Returns the server's Excel version string.
    pub fn version(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.version"), false),
            vec![],
        )?
        .as_string()
    }
    /// Returns the current visibility of the crate-created Excel window.
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.visible"), false),
            vec![],
        )?
        .as_bool()
        .ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType { vartype: 0 },
        ))
    }
    /// Sets the visibility of the crate-created Excel window.
    pub fn set_visible(&self, value: bool) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.visible"), true),
            OwnedVariant::bool(value),
        )?;
        Ok(())
    }
    /// Returns Excel's `DisplayAlerts` setting.
    pub fn display_alerts(&self) -> Result<bool, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.displayalerts"), false),
            vec![],
        )?;
        result.as_bool().ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType {
                vartype: result.vt(),
            },
        ))
    }
    /// Sets Excel's `DisplayAlerts` setting.
    pub fn set_display_alerts(&self, value: bool) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.displayalerts"), true),
            OwnedVariant::bool(value),
        )?;
        Ok(())
    }
    /// Sets `DisplayAlerts` and returns a guard that restores its prior value.
    pub fn display_alerts_guard(
        &self,
        value: bool,
    ) -> Result<DisplayAlertsGuard<'_>, ExcelComError> {
        let previous = self.display_alerts()?;
        self.set_display_alerts(value)?;
        Ok(DisplayAlertsGuard {
            application: self,
            previous,
            active: true,
        })
    }
    /// Returns Excel's process-wide reference notation setting.
    pub fn reference_style(&self) -> Result<ReferenceStyle, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.referencestyle"), false),
            vec![],
        )?;
        result
            .as_i32()
            .or_else(|| result.as_scode())
            .map(ReferenceStyle::from_raw)
            .ok_or(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType {
                    vartype: result.vt(),
                },
            ))
    }

    fn set_reference_style(&self, value: ReferenceStyle) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.referencestyle"), true),
            OwnedVariant::i32(value.raw()),
        )?;
        Ok(())
    }

    /// Temporarily changes Excel's global reference style and returns a restoring guard.
    pub fn reference_style_guard(
        &self,
        style: ReferenceStyle,
    ) -> Result<ReferenceStyleGuard<'_>, ExcelComError> {
        let previous = self.reference_style()?;
        self.set_reference_style(style)?;
        Ok(ReferenceStyleGuard {
            application: self,
            previous,
            active: true,
        })
    }

    /// Returns Excel's process-wide calculation mode.
    pub fn calculation_mode(&self) -> Result<CalculationMode, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.calculation"), false),
            vec![],
        )?;
        result
            .as_i32()
            .map(CalculationMode::from_raw)
            .ok_or(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType {
                    vartype: result.vt(),
                },
            ))
    }

    fn set_calculation_mode(&self, mode: CalculationMode) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.calculation"), true),
            OwnedVariant::i32(mode.raw()),
        )?;
        Ok(())
    }

    /// Sets a temporary global calculation mode and returns a restoring guard.
    ///
    /// Prefer this scoped mutation to persistent mode changes. The guard makes
    /// a best-effort restoration on drop and exposes [`CalculationModeGuard::restore`]
    /// for callers that need the restoration error.
    ///
    /// ```no_run
    /// # fn example(application: &excel_com::Application, range: &excel_com::Range) -> Result<(), excel_com::ExcelComError> {
    /// use excel_com::{AutomationValue, CalculationMode};
    /// let guard = application.calculation_mode_guard(CalculationMode::MANUAL)?;
    /// range.set_value2(AutomationValue::Number(10.0))?;
    /// range.calculate()?;
    /// guard.restore()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn calculation_mode_guard(
        &self,
        mode: CalculationMode,
    ) -> Result<CalculationModeGuard<'_>, ExcelComError> {
        let previous = self.calculation_mode()?;
        self.set_calculation_mode(mode)?;
        Ok(CalculationModeGuard {
            application: self,
            previous,
            active: true,
        })
    }

    /// Returns Excel's current calculation-state snapshot.
    ///
    /// A `DONE` snapshot after a calculation call is an observation, not a
    /// general promise that every future calculation completes synchronously.
    pub fn calculation_state(&self) -> Result<CalculationState, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.calculationstate"), false),
            vec![],
        )?;
        result
            .as_i32()
            .map(CalculationState::from_raw)
            .ok_or(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType {
                    vartype: result.vt(),
                },
            ))
    }

    /// Returns Excel's calculation-engine version number.
    pub fn calculation_version(&self) -> Result<i32, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.calculationversion"), false),
            vec![],
        )?;
        result.as_i32().ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType {
                vartype: result.vt(),
            },
        ))
    }

    /// Returns whether Excel calculates workbooks before saving in manual mode.
    pub fn calculate_before_save(&self) -> Result<bool, ExcelComError> {
        self.bool_property("excel.application.calculatebeforesave")
    }

    /// Changes Excel's calculate-before-save setting.
    ///
    /// This global setting is not mutated by [`Self::calculation_mode_guard`];
    /// callers that change it are responsible for restoring its prior value.
    pub fn set_calculate_before_save(&self, enabled: bool) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.calculatebeforesave"), true),
            OwnedVariant::bool(enabled),
        )?;
        Ok(())
    }

    /// Calculates all open workbooks through Excel.
    pub fn calculate(&self) -> Result<(), ExcelComError> {
        self.calculation_method("excel.application.calculate")
    }

    /// Forces a full calculation of all open workbooks through Excel.
    pub fn calculate_full(&self) -> Result<(), ExcelComError> {
        self.calculation_method("excel.application.calculatefull")
    }

    /// Forces a full calculation and dependency rebuild of all open workbooks.
    ///
    /// This can be expensive and is normally appropriate only for controlled
    /// maintenance or compatibility operations.
    pub fn calculate_full_rebuild(&self) -> Result<(), ExcelComError> {
        self.calculation_method("excel.application.calculatefullrebuild")
    }

    /// Converts a formula or address through Excel's `ConvertFormula` engine.
    ///
    /// The supplied `from` and `to` styles are explicit. Optional conversion
    /// mode and relative base preserve their positions as `Missing` when not
    /// supplied; no formula parser is implemented in Rust.
    ///
    /// ```no_run
    /// # fn example(application: &excel_com::Application) -> Result<(), excel_com::ExcelComError> {
    /// use excel_com::{FormulaConversionOptions, ReferenceStyle};
    /// let converted = application.convert_formula(
    ///     "=SUM(A1:X10)",
    ///     ReferenceStyle::A1,
    ///     ReferenceStyle::R1C1,
    ///     &FormulaConversionOptions::default(),
    /// )?;
    /// # assert!(!converted.is_empty());
    /// # Ok(())
    /// # }
    /// ```
    pub fn convert_formula(
        &self,
        formula: &str,
        from: ReferenceStyle,
        to: ReferenceStyle,
        options: &FormulaConversionOptions<'_>,
    ) -> Result<String, ExcelComError> {
        invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.convertformula"), false),
            convert_formula_arguments(formula, from, to, options)?,
            false,
        )?
        .as_string()
    }

    /// Evaluates an expression and requires a scalar or array Automation value.
    ///
    /// If Excel returns an object, use [`Self::evaluate_range`] instead.
    ///
    /// ```no_run
    /// # fn example(application: &excel_com::Application) -> Result<(), excel_com::ExcelComError> {
    /// let result = application.evaluate_value("SUM(A1:A10)")?;
    /// # drop(result);
    /// # Ok(())
    /// # }
    /// ```
    pub fn evaluate_value(
        &self,
        expression: &str,
    ) -> Result<crate::AutomationValue, ExcelComError> {
        let result = self.evaluate(expression)?;
        if result.vt() == windows_sys::Win32::System::Variant::VT_DISPATCH {
            return Err(ExcelComError::Unsupported {
                detail: "Application.evaluate_value returned an object; use Application.evaluate_range",
            });
        }
        decode_variant(&result, ConversionPolicy::default())
    }

    /// Evaluates an expression and requires an Excel Range dispatch result.
    ///
    /// If Excel returns a scalar or array value, use [`Self::evaluate_value`]
    /// instead.
    ///
    /// ```no_run
    /// # fn example(application: &excel_com::Application) -> Result<(), excel_com::ExcelComError> {
    /// let resolved = application.evaluate_range("InputRange")?;
    /// # drop(resolved);
    /// # Ok(())
    /// # }
    /// ```
    pub fn evaluate_range(&self, expression: &str) -> Result<Range, ExcelComError> {
        let mut result = self.evaluate(expression)?;
        if result.vt() != windows_sys::Win32::System::Variant::VT_DISPATCH {
            return Err(ExcelComError::Unsupported {
                detail: "Application.evaluate_range returned a value; use Application.evaluate_value",
            });
        }
        Ok(Range::from_dispatch(result.take_dispatch()?))
    }

    fn evaluate(&self, expression: &str) -> Result<OwnedVariant, ExcelComError> {
        invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.evaluate-1"), false),
            vec![text_bstr(expression)?],
            false,
        )
    }

    fn bool_property(&self, id: &'static str) -> Result<bool, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new(id), false),
            vec![],
        )?;
        result.as_bool().ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType {
                vartype: result.vt(),
            },
        ))
    }

    fn calculation_method(&self, id: &'static str) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new(id), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Returns the application's `Workbooks` collection.
    pub fn workbooks(&self) -> Result<Workbooks, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.workbooks"), false),
            vec![],
        )?;
        Ok(Workbooks::from_dispatch(result.take_dispatch()?))
    }
    /// Returns the union of two ranges from this Application's Excel server.
    ///
    /// Both Range values must belong to compatible workbooks in the same Excel
    /// instance; Excel reports any cross-server or cross-workbook restriction.
    pub fn union2(&self, left: &Range, right: &Range) -> Result<Range, ExcelComError> {
        let mut arguments = crate::automation::PositionalArguments::new();
        arguments.push_object(left.dispatch_object());
        arguments.push_object(right.dispatch_object());
        let mut result = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.union"), false),
            arguments.into_inner(),
            false,
        )?;
        Ok(Range::from_dispatch(result.take_dispatch()?))
    }
    /// Explicitly asks the crate-created application to quit. `Drop` never does this.
    pub fn quit(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.quit"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

fn convert_formula_arguments(
    formula: &str,
    from: ReferenceStyle,
    to: ReferenceStyle,
    options: &FormulaConversionOptions<'_>,
) -> Result<Vec<OwnedVariant>, ExcelComError> {
    let mut arguments = PositionalArguments::new();
    arguments.push_result(text_bstr(formula))?;
    arguments.push_required(OwnedVariant::i32(from.raw()));
    arguments.push_required(OwnedVariant::i32(to.raw()));
    arguments.push_optional(
        options
            .to_absolute
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    arguments.push_optional_object(options.relative_to.map(Range::dispatch_object));
    Ok(arguments.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ReferenceAbsoluteMode;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn convert_formula_preserves_optional_positions() {
        let values = convert_formula_arguments(
            "=A1",
            ReferenceStyle::A1,
            ReferenceStyle::R1C1,
            &FormulaConversionOptions::default(),
        )
        .expect("arguments");
        assert_eq!(values.len(), 5);
        assert_eq!(values[0].as_string().expect("formula"), "=A1");
        assert_eq!(values[1].as_i32(), Some(1));
        assert_eq!(values[2].as_i32(), Some(-4150));
        assert_eq!(values[3].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[4].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }

    #[test]
    fn convert_formula_encodes_absolute_mode_and_rejects_nul() {
        let values = convert_formula_arguments(
            "A1",
            ReferenceStyle::A1,
            ReferenceStyle::R1C1,
            &FormulaConversionOptions {
                to_absolute: Some(ReferenceAbsoluteMode::RELATIVE),
                relative_to: None,
            },
        )
        .expect("arguments");
        assert_eq!(values[3].as_i32(), Some(4));
        assert!(
            convert_formula_arguments(
                "A\0",
                ReferenceStyle::A1,
                ReferenceStyle::R1C1,
                &FormulaConversionOptions::default(),
            )
            .is_err()
        );
    }
}
