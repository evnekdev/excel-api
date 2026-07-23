//! Focused drawing implementation component.
#![allow(missing_docs)]
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
/// A Series trendline collection.
pub struct Trendlines {
    inner: DispatchObject,
}
impl Debug for Trendlines {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Trendlines").field(&self.inner).finish()
    }
}
impl Trendlines {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Trendlines", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, TRENDLINES)
    }
    pub fn item(&self, index: usize) -> Result<Trendline, ExcelComError> {
        collection_item(&self.inner, TRENDLINES, index, Trendline::from_dispatch)
    }
    pub fn iter(&self) -> Result<TrendlinesIter, ExcelComError> {
        Ok(TrendlinesIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(TRENDLINES.new_enum),
                TRENDLINES.name,
            )?,
            index: 0,
            done: false,
        })
    }
    pub fn add(&self, options: &TrendlineAddOptions<'_>) -> Result<Trendline, ExcelComError> {
        if let Some(value) = options.order {
            if value == 0 {
                return Err(ExcelComError::Unsupported {
                    detail: "trendline order must be positive",
                });
            }
        }
        if let Some(value) = options.period {
            if value == 0 {
                return Err(ExcelComError::Unsupported {
                    detail: "trendline period must be positive",
                });
            }
        }
        for value in [options.forward, options.backward, options.intercept]
            .into_iter()
            .flatten()
        {
            finite(value, "trendline value must be finite")?;
        }
        if let Some(value) = options.name {
            let _ = text_bstr(value)?;
        }
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(options.trendline_type.raw()));
        args.push_optional(
            options
                .order
                .map(|value| i32::try_from(value).map(OwnedVariant::i32))
                .transpose()
                .map_err(|_| ExcelComError::Unsupported {
                    detail: "trendline order exceeds i32",
                })?,
        );
        args.push_optional(
            options
                .period
                .map(|value| i32::try_from(value).map(OwnedVariant::i32))
                .transpose()
                .map_err(|_| ExcelComError::Unsupported {
                    detail: "trendline period exceeds i32",
                })?,
        );
        args.push_optional(options.forward.map(OwnedVariant::f64));
        args.push_optional(options.backward.map(OwnedVariant::f64));
        args.push_optional(options.intercept.map(OwnedVariant::f64));
        args.push_optional(options.display_equation.map(OwnedVariant::bool));
        args.push_optional(options.display_r_squared.map(OwnedVariant::bool));
        match options.name {
            Some(value) => args.push_result(text_bstr(value))?,
            None => args.push_optional(None),
        };
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.trendlines.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Trendline::from_dispatch(value.take_dispatch()?))
    }
}
pub struct TrendlinesIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for TrendlinesIter {
    type Item = Result<Trendline, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "Trendlines", index)
                        .map(Trendline::from_dispatch),
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
impl FusedIterator for TrendlinesIter {}
pub struct Trendline {
    inner: DispatchObject,
}
impl Debug for Trendline {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Trendline").field(&self.inner).finish()
    }
}
impl Trendline {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Trendline", value),
        }
    }
    pub fn trendline_type(&self) -> Result<TrendlineType, ExcelComError> {
        Ok(TrendlineType::from_raw(get_i32(
            &self.inner,
            "excel.trendline.type",
            "Trendline.Type was not an integer",
        )?))
    }
    /// Returns the trendline's display name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.trendline.name")
    }
    /// Sets the trendline's display name.
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.trendline.name", text_bstr(value)?)
    }
    /// Returns the polynomial order for an applicable trendline.
    pub fn order(&self) -> Result<usize, ExcelComError> {
        usize::try_from(get_i32(
            &self.inner,
            "excel.trendline.order",
            "Trendline.Order was not an integer",
        )?)
        .map_err(|_| ExcelComError::Unsupported {
            detail: "Trendline.Order was negative",
        })
    }
    /// Sets an applicable polynomial order from 2 through 6.
    pub fn set_order(&self, value: usize) -> Result<(), ExcelComError> {
        if !(2..=6).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "polynomial Trendline.Order must be between 2 and 6",
            });
        }
        put(
            &self.inner,
            "excel.trendline.order",
            one_based(value, "trendline order must be positive")?,
        )
    }
    /// Returns the moving-average period for an applicable trendline.
    pub fn period(&self) -> Result<usize, ExcelComError> {
        usize::try_from(get_i32(
            &self.inner,
            "excel.trendline.period",
            "Trendline.Period was not an integer",
        )?)
        .map_err(|_| ExcelComError::Unsupported {
            detail: "Trendline.Period was negative",
        })
    }
    /// Sets an applicable moving-average period of at least 2.
    pub fn set_period(&self, value: usize) -> Result<(), ExcelComError> {
        if value < 2 {
            return Err(ExcelComError::Unsupported {
                detail: "moving-average Trendline.Period must be at least 2",
            });
        }
        put(
            &self.inner,
            "excel.trendline.period",
            one_based(value, "trendline period must be positive")?,
        )
    }
    /// Returns the finite forward forecast amount for this trendline.
    pub fn forward(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.trendline.forward",
            "Trendline.Forward was not numeric",
        )
    }
    /// Sets a finite forward forecast amount.
    pub fn set_forward(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "Trendline.Forward must be finite")?;
        put(
            &self.inner,
            "excel.trendline.forward",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the finite backward forecast amount for this trendline.
    pub fn backward(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.trendline.backward",
            "Trendline.Backward was not numeric",
        )
    }
    /// Sets a finite backward forecast amount.
    pub fn set_backward(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "Trendline.Backward must be finite")?;
        put(
            &self.inner,
            "excel.trendline.backward",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the finite intercept for an applicable trendline.
    pub fn intercept(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.trendline.intercept",
            "Trendline.Intercept was not numeric",
        )
    }
    /// Sets a finite intercept for an applicable trendline.
    pub fn set_intercept(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "Trendline.Intercept must be finite")?;
        put(
            &self.inner,
            "excel.trendline.intercept",
            OwnedVariant::f64(value),
        )
    }
    pub fn display_equation(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.trendline.displayequation")
    }
    pub fn set_display_equation(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.trendline.displayequation",
            OwnedVariant::bool(value),
        )
    }
    pub fn display_r_squared(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.trendline.displayrsquared")
    }
    pub fn set_display_r_squared(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.trendline.displayrsquared",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the trendline's data label when Excel exposes one.
    pub fn data_label(&self) -> Result<Option<DataLabel>, ExcelComError> {
        optional_dispatch(
            &self.inner,
            "excel.trendline.datalabel",
            DataLabel::from_dispatch,
        )
    }
    /// Returns Office drawing formatting for this trendline.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.trendline.format",
            ChartFormat::from_dispatch,
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.trendline.delete", vec![])
    }
}
