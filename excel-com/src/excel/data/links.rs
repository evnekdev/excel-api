use crate::ExcelComError;
use crate::automation::{OwnedVariant, invoke, property_get};
use crate::excel::{Application, Workbook};
use crate::object_model::{MemberId, member};

use super::helpers::text_argument;
use super::{ExternalLinkSource, LinkStatus, LinkType};

/// Restores Excel's process-wide `AskToUpdateLinks` state when explicitly restored or dropped.
pub struct AskToUpdateLinksGuard<'a> {
    application: &'a Application,
    previous: bool,
    active: bool,
}
impl std::fmt::Debug for AskToUpdateLinksGuard<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("AskToUpdateLinksGuard")
            .field("previous", &self.previous)
            .field("active", &self.active)
            .finish()
    }
}
impl AskToUpdateLinksGuard<'_> {
    /// Restores the captured prompt setting and disarms this guard.
    pub fn restore(mut self) -> Result<(), ExcelComError> {
        self.application.set_ask_to_update_links(self.previous)?;
        self.active = false;
        Ok(())
    }
}
impl Drop for AskToUpdateLinksGuard<'_> {
    fn drop(&mut self) {
        if self.active {
            let _ = self.application.set_ask_to_update_links(self.previous);
        }
    }
}

impl Application {
    /// Returns Excel's global external-link prompt setting. Macro security is a separate control.
    pub fn ask_to_update_links(&self) -> Result<bool, ExcelComError> {
        property_get(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.application.asktoupdatelinks"), false),
            vec![],
        )?
        .as_bool()
        .ok_or(ExcelComError::Unsupported {
            detail: "Application.AskToUpdateLinks did not return VT_BOOL",
        })
    }
    fn set_ask_to_update_links(&self, value: bool) -> Result<(), ExcelComError> {
        let _ = crate::automation::property_put(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.application.asktoupdatelinks"), true),
            OwnedVariant::bool(value),
        )?;
        Ok(())
    }
    /// Temporarily changes only the external-link prompt and restores it on drop.
    pub fn ask_to_update_links_guard(
        &self,
        value: bool,
    ) -> Result<AskToUpdateLinksGuard<'_>, ExcelComError> {
        let previous = self.ask_to_update_links()?;
        self.set_ask_to_update_links(value)?;
        Ok(AskToUpdateLinksGuard {
            application: self,
            previous,
            active: true,
        })
    }
}

impl Workbook {
    /// Enumerates exact external-link source strings returned by Excel without normalizing paths.
    pub fn link_sources(
        &self,
        link_type: Option<LinkType>,
    ) -> Result<Vec<ExternalLinkSource>, ExcelComError> {
        let link_type = link_type.unwrap_or(LinkType::EXCEL_LINK);
        let value = property_get(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.workbook.linksources"), false),
            vec![OwnedVariant::i32(link_type.raw())],
        )?;
        let decoded =
            crate::automation::decode_variant(&value, crate::ConversionPolicy::default())?;
        let crate::AutomationValue::Array(array) = decoded else {
            return Ok(Vec::new());
        };
        let mut result = Vec::new();
        for value in array.values() {
            if let crate::AutomationValue::Text(name) = value {
                result.push(ExternalLinkSource {
                    name: name.clone(),
                    link_type,
                });
            }
        }
        Ok(result)
    }
    /// Requests an update for an explicitly named, caller-owned link. This wrapper does not resolve paths.
    pub fn update_link(&self, name: &str, link_type: LinkType) -> Result<(), ExcelComError> {
        link_operation(
            self,
            "excel.workbook.updatelink",
            vec![text_argument(name)?, OwnedVariant::i32(link_type.raw())],
        )
    }
    /// Destructively breaks an explicitly named link; Excel replaces formulas with their current values.
    pub fn break_link(&self, name: &str, link_type: LinkType) -> Result<(), ExcelComError> {
        link_operation(
            self,
            "excel.workbook.breaklink",
            vec![text_argument(name)?, OwnedVariant::i32(link_type.raw())],
        )
    }
    /// Changes one explicit link source name to another without path normalization.
    pub fn change_link(
        &self,
        old_name: &str,
        new_name: &str,
        link_type: LinkType,
    ) -> Result<(), ExcelComError> {
        link_operation(
            self,
            "excel.workbook.changelink",
            vec![
                text_argument(old_name)?,
                text_argument(new_name)?,
                OwnedVariant::i32(link_type.raw()),
            ],
        )
    }
    /// Returns Excel's status for an explicitly named external link.
    pub fn link_status(
        &self,
        name: &str,
        link_type: LinkType,
    ) -> Result<LinkStatus, ExcelComError> {
        let result = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.workbook.linkinfo"), false),
            vec![
                text_argument(name)?,
                OwnedVariant::i32(link_type.raw()),
                OwnedVariant::i32(1),
            ],
            false,
        )?;
        result
            .as_i32()
            .map(LinkStatus::from_raw)
            .ok_or(ExcelComError::Unsupported {
                detail: "Workbook.LinkInfo did not return VT_I4",
            })
    }
}

fn link_operation(
    workbook: &Workbook,
    id: &'static str,
    values: Vec<OwnedVariant>,
) -> Result<(), ExcelComError> {
    let _ = invoke(
        &workbook.dispatch_object().dispatch,
        member(MemberId::new(id), false),
        values,
        false,
    )?;
    Ok(())
}
