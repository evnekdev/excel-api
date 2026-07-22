use std::path::Path;

use crate::{ExcelComError, automation::OwnedVariant};

/// Converts a caller path directly from its Windows UTF-16 representation.
pub(crate) fn path_bstr(path: &Path) -> Result<OwnedVariant, ExcelComError> {
    use std::os::windows::ffi::OsStrExt;

    let units: Vec<u16> = path.as_os_str().encode_wide().collect();
    if units.contains(&0) {
        return Err(ExcelComError::InvalidPath {
            detail: "embedded NUL is not supported by Excel Automation",
        });
    }
    OwnedVariant::bstr_wide(&units)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn windows_path_units_preserve_unicode_and_spaces() {
        let value = path_bstr(Path::new("relative folder/資料 book.xlsx"));
        assert!(value.is_ok());
    }

    #[test]
    fn embedded_nul_path_is_rejected_without_lossy_conversion() {
        assert!(matches!(
            path_bstr(Path::new("bad\0path.xlsx")),
            Err(ExcelComError::InvalidPath { .. })
        ));
    }
}
