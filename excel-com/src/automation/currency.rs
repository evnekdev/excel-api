/// Exact COM `CY` storage scaled by 10,000.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd, Hash)]
pub struct Currency(i64);
impl Currency {
    /// Number of scaled units in one currency unit.
    pub const SCALE: i64 = 10_000;
    /// Creates currency from its exact COM scaled integer representation.
    pub const fn from_scaled(value: i64) -> Self {
        Self(value)
    }
    /// Returns the exact COM scaled integer representation.
    pub const fn scaled(self) -> i64 {
        self.0
    }
}
