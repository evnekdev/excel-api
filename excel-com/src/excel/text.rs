use crate::{ExcelComError, automation::OwnedVariant};

/// Encodes caller text after rejecting the NUL character before COM.
pub(crate) fn text_bstr(value: &str) -> Result<OwnedVariant, ExcelComError> {
    if value.contains('\0') {
        return Err(ExcelComError::Unsupported {
            detail: "embedded NUL is not supported by Excel Automation text input",
        });
    }
    OwnedVariant::bstr(value)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn embedded_nul_is_rejected_before_com() {
        assert!(matches!(
            text_bstr("A\0B"),
            Err(ExcelComError::Unsupported { .. })
        ));
    }
}
