use super::AutomationValue;

/// Invocation-only distinction between a supplied Automation value and Missing.
#[derive(Clone, Debug, PartialEq)]
pub enum AutomationArgument {
    Value(AutomationValue),
    Missing,
}
