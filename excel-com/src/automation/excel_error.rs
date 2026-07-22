/// A lossless signed Excel Automation error SCODE.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct ExcelError(i32);
impl ExcelError {
    /// The Excel `#N/A` error encoded as its full signed Automation SCODE.
    pub const NOT_AVAILABLE: Self = Self(0x800A_07FA_u32 as i32);
    /// Creates an error from its full signed Automation SCODE.
    pub const fn from_scode(value: i32) -> Self {
        Self(value)
    }
    /// Returns the full signed Automation SCODE without normalizing it.
    pub const fn scode(self) -> i32 {
        self.0
    }
}
