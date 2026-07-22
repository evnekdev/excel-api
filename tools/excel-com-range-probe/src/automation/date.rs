//! OLE Automation date serials.

use super::ConversionError;

#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct OaDate(pub(super) f64);

impl OaDate {
    pub(crate) fn new(serial: f64) -> Result<Self, ConversionError> {
        serial.is_finite().then_some(Self(serial)).ok_or(ConversionError::NonFiniteNumber)
    }

    pub(crate) const fn serial(self) -> f64 { self.0 }
}
