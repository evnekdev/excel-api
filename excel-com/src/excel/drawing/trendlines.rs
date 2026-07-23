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
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.trendline.delete", vec![])
    }
}
