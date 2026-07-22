/// A lossless signed Excel Automation error SCODE.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct ExcelError(i32);
impl ExcelError {
    pub const NOT_AVAILABLE: Self = Self(0x800A_07FA_u32 as i32);
    pub const fn from_scode(value: i32) -> Self {
        Self(value)
    }
    pub const fn scode(self) -> i32 {
        self.0
    }
}
