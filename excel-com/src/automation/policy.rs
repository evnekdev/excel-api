/// Explicit, strict conversion policy; worksheet-specific limits stay at wrapper boundaries.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ConversionPolicy {
    /// Whether non-finite floating-point values are rejected before COM calls.
    pub reject_non_finite_numbers: bool,
    /// Whether strings containing an embedded NUL are rejected before BSTR allocation.
    pub reject_embedded_nul: bool,
    /// Date representation selected by the calling Range member.
    pub(crate) date_write: super::DateWriteMode,
}
impl Default for ConversionPolicy {
    fn default() -> Self {
        Self {
            reject_non_finite_numbers: true,
            reject_embedded_nul: true,
            date_write: super::DateWriteMode::Value,
        }
    }
}
