//! Pointer-free internal Automation values.

use super::{AutomationArray, Currency, ExcelError, OaDate};

#[derive(Clone, Debug, PartialEq)]
pub(crate) enum AutomationValue {
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
