//! Forward-compatible Excel calculation representations.

/// An Excel `XlCalculation` value.
///
/// The value is process-wide Application state. Unknown values are retained so
/// a newer Excel version does not make an otherwise readable setting
/// unrepresentable.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct CalculationMode(i32);

impl CalculationMode {
    /// `xlCalculationAutomatic`.
    pub const AUTOMATIC: Self = Self(-4105);
    /// `xlCalculationManual`.
    pub const MANUAL: Self = Self(-4135);
    /// `xlCalculationSemiautomatic`.
    pub const SEMIAUTOMATIC: Self = Self(2);

    /// Preserves an Excel calculation-mode value without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw `XlCalculation` value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// An Excel `XlCalculationState` value.
///
/// This is a snapshot reported by Excel, not a completion guarantee for a
/// preceding calculation call. Unknown values are retained.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct CalculationState(i32);

impl CalculationState {
    /// `xlDone`.
    pub const DONE: Self = Self(0);
    /// `xlCalculating`.
    pub const CALCULATING: Self = Self(1);
    /// `xlPending`.
    pub const PENDING: Self = Self(2);

    /// Preserves an Excel calculation-state value without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw `XlCalculationState` value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

#[cfg(test)]
mod tests {
    use super::{CalculationMode, CalculationState};

    #[test]
    fn calculation_values_preserve_known_and_unknown_values() {
        assert_eq!(CalculationMode::AUTOMATIC.raw(), -4105);
        assert_eq!(CalculationMode::MANUAL.raw(), -4135);
        assert_eq!(CalculationMode::SEMIAUTOMATIC.raw(), 2);
        assert_eq!(CalculationMode::from_raw(701).raw(), 701);
        assert_eq!(CalculationState::DONE.raw(), 0);
        assert_eq!(CalculationState::CALCULATING.raw(), 1);
        assert_eq!(CalculationState::PENDING.raw(), 2);
        assert_eq!(CalculationState::from_raw(702).raw(), 702);
    }
}
