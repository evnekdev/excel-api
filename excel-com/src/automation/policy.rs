/// Explicit, strict conversion policy; worksheet-specific limits stay at wrapper boundaries.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ConversionPolicy {
    pub reject_non_finite_numbers: bool,
    pub reject_embedded_nul: bool,
}
impl Default for ConversionPolicy {
    fn default() -> Self {
        Self {
            reject_non_finite_numbers: true,
            reject_embedded_nul: true,
        }
    }
}
