use crate::automation::{OwnedVariant, PositionalArguments, invoke, property_get};
use crate::excel::{DispatchObject, Workbook, WorkbookOpenOptions};
use crate::internal::{ComPtr, Dispatch, path_bstr};
use crate::object_model::{MemberId, member};
use crate::{ConversionError, ExcelComError};
use std::fmt::{Debug, Formatter};
use std::path::Path;

/// Experimental wrapper for an Excel Workbooks collection.
pub struct Workbooks {
    inner: DispatchObject,
}
impl Debug for Workbooks {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Workbooks")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for Workbooks {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Workbooks {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Workbooks",
            },
        }
    }
    /// Returns the number of open workbooks.
    pub fn count(&self) -> Result<i32, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbooks.count"), false),
            vec![],
        )?
        .as_i32()
        .ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType { vartype: 0 },
        ))
    }
    /// Adds a default workbook using Excel's proven property-get invocation form.
    pub fn add(&self) -> Result<Workbook, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbooks.add"), false),
            vec![],
        )?;
        Ok(Workbook::from_dispatch(result.take_dispatch()?))
    }
    /// Opens a workbook using the exact 15-position `Workbooks.Open` signature.
    ///
    /// The provided path is passed as its original Windows UTF-16 `OsStr`
    /// units, without canonicalization or lossy conversion. Every omitted
    /// optional position remains an explicit Automation `Missing` argument.
    pub fn open<P: AsRef<Path>>(
        &self,
        filename: P,
        options: WorkbookOpenOptions<'_>,
    ) -> Result<Workbook, ExcelComError> {
        let arguments = open_arguments(filename.as_ref(), options)?;
        let mut result = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbooks.open-1923"), false),
            arguments,
            false,
        )?;
        Ok(Workbook::from_dispatch(result.take_dispatch()?))
    }
    /// Opens a workbook with every optional `Open` position explicitly missing.
    pub fn open_default<P: AsRef<Path>>(&self, filename: P) -> Result<Workbook, ExcelComError> {
        self.open(filename, WorkbookOpenOptions::new())
    }
}

fn open_arguments(
    filename: &Path,
    options: WorkbookOpenOptions<'_>,
) -> Result<Vec<OwnedVariant>, ExcelComError> {
    let mut arguments = PositionalArguments::new();
    arguments.push_result(path_bstr(filename))?;
    arguments.push_optional(
        options
            .update_links
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    arguments.push_optional(options.read_only.map(OwnedVariant::bool));
    arguments.push_optional(options.format.map(|value| OwnedVariant::i32(value.raw())));
    push_optional_text(&mut arguments, options.password)?;
    push_optional_text(&mut arguments, options.write_res_password)?;
    arguments.push_optional(options.ignore_read_only_recommended.map(OwnedVariant::bool));
    arguments.push_optional(options.origin.map(|value| OwnedVariant::i32(value.raw())));
    if let Some(delimiter) = options.delimiter {
        arguments.push_result(OwnedVariant::bstr(&delimiter.to_string()))?;
    } else {
        arguments.push_optional(None);
    }
    arguments.push_optional(options.editable.map(OwnedVariant::bool));
    arguments.push_optional(options.notify.map(OwnedVariant::bool));
    arguments.push_optional(options.converter.map(OwnedVariant::i32));
    arguments.push_optional(options.add_to_mru.map(OwnedVariant::bool));
    arguments.push_optional(options.local.map(OwnedVariant::bool));
    arguments.push_optional(
        options
            .corrupt_load
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    Ok(arguments.into_inner())
}

fn push_optional_text(
    arguments: &mut PositionalArguments,
    value: Option<&str>,
) -> Result<(), ExcelComError> {
    match value {
        Some(value) => arguments.push_result(OwnedVariant::bstr(value)),
        None => {
            arguments.push_optional(None);
            Ok(())
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn open_options_fill_all_exact_positions_without_trimming() {
        let values = open_arguments(
            Path::new("relative folder/資料 book.xlsx"),
            WorkbookOpenOptions {
                read_only: Some(true),
                converter: Some(9),
                local: Some(false),
                ..WorkbookOpenOptions::new()
            },
        )
        .expect("path arguments");
        assert_eq!(values.len(), 15);
        assert_eq!(values[1].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[2].as_bool(), Some(true));
        assert_eq!(values[11].as_i32(), Some(9));
        assert_eq!(values[13].as_bool(), Some(false));
        assert_eq!(values[14].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }

    #[test]
    fn open_option_debug_redacts_passwords() {
        let text = format!(
            "{:?}",
            WorkbookOpenOptions {
                password: Some("redaction-input-one"),
                write_res_password: Some("redaction-input-two"),
                ..WorkbookOpenOptions::new()
            }
        );
        assert!(text.contains("REDACTED"));
        assert!(!text.contains("redaction-input-one"));
        assert!(!text.contains("redaction-input-two"));
    }
}
