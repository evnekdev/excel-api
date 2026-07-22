use crate::ExcelComError;
/// Finite OLE Automation date serial; calendar interpretation is out of scope.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct OaDate(f64);
impl OaDate {
    pub fn new(value: f64) -> Result<Self, ExcelComError> {
        value
            .is_finite()
            .then_some(Self(value))
            .ok_or(ExcelComError::Conversion {
                detail: "date serial must be finite",
            })
    }
    pub const fn serial(self) -> f64 {
        self.0
    }
}
