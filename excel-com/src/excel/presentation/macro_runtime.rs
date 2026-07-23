//! Macro-security guard and safe-open option types.
#![allow(missing_docs)]

use super::*;

/// Restores the previous process-global Automation macro-security setting on drop.
pub struct AutomationSecurityGuard<'a> {
    pub(super) application: &'a Application,
    pub(super) previous: AutomationSecurity,
    pub(super) active: bool,
}
impl Debug for AutomationSecurityGuard<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("AutomationSecurityGuard")
            .field("previous", &self.previous)
            .field("active", &self.active)
            .finish()
    }
}
impl AutomationSecurityGuard<'_> {
    /// Restores the previous setting and disarms the guard.
    pub fn restore(mut self) -> Result<(), ExcelComError> {
        self.application.set_automation_security(self.previous)?;
        self.active = false;
        Ok(())
    }
}
impl Drop for AutomationSecurityGuard<'_> {
    fn drop(&mut self) {
        if self.active {
            let _ = self.application.set_automation_security(self.previous);
            self.active = false;
        }
    }
}

/// Options for opening a workbook with macros forcibly disabled.
#[derive(Clone, Default, PartialEq)]
pub struct SafeWorkbookOpenOptions<'a> {
    pub open: WorkbookOpenOptions<'a>,
    /// Optional temporary value for the separate Excel external-link prompt.
    /// Macro execution remains controlled independently by `AutomationSecurity`.
    pub link_prompt: Option<bool>,
}
impl Debug for SafeWorkbookOpenOptions<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("SafeWorkbookOpenOptions")
            .field("open", &self.open)
            .field("link_prompt", &self.link_prompt)
            .finish()
    }
}
