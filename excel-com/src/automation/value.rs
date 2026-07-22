use super::{AutomationArray, Currency, ExcelError, OaDate};
/// Experimental pointer-free Automation semantic value.
#[derive(Clone, Debug, PartialEq)]
pub enum AutomationValue {
    Empty,
    Null,
    Bool(bool),
    Number(f64),
    Text(String),
    Error(ExcelError),
    Date(OaDate),
    Currency(Currency),
    Array(AutomationArray),
}
