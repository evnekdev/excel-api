//! Exact COM `CY` fixed-scale values.

use super::ConversionError;

#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd, Hash)]
pub(crate) struct Currency(pub(super) i64);

impl Currency {
    pub(crate) const SCALE: i64 = 10_000;
    pub(crate) const fn from_scaled(scaled: i64) -> Self { Self(scaled) }
    pub(crate) const fn scaled(self) -> i64 { self.0 }

    #[allow(dead_code)]
    pub(crate) fn from_decimal_parts(whole: i64, fractional_ten_thousandths: u16) -> Result<Self, ConversionError> {
        if fractional_ten_thousandths >= Self::SCALE as u16 {
            return Err(ConversionError::CurrencyOverflow);
        }
        whole.checked_mul(Self::SCALE)
            .and_then(|scaled| scaled.checked_add(if whole < 0 { -(fractional_ten_thousandths as i64) } else { fractional_ten_thousandths as i64 }))
            .map(Self)
            .ok_or(ConversionError::CurrencyOverflow)
    }
}
