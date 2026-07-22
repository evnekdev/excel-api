use crate::automation::{OwnedVariant, PositionalArguments, invoke, property_get, property_put};
use crate::excel::{
    DispatchObject, SaveChanges, WorkbookCloseOptions, WorkbookSaveAsOptions, Worksheets,
    XlFileFormat,
};
use crate::internal::{ComPtr, Dispatch, path_bstr};
use crate::object_model::{MemberId, member};
use crate::{ConversionError, ExcelComError};
use std::fmt::{Debug, Formatter};
use std::path::Path;

/// Experimental wrapper for a single Excel Workbook.
pub struct Workbook {
    inner: DispatchObject,
}
impl Debug for Workbook {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Workbook")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for Workbook {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Workbook {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Workbook",
            },
        }
    }
    /// Returns the workbook name reported by Excel.
    pub fn name(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.name"), false),
            vec![],
        )?
        .as_string()
    }
    /// Returns the workbook's worksheet collection.
    pub fn worksheets(&self) -> Result<Worksheets, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.worksheets"), false),
            vec![],
        )?;
        Ok(Worksheets::from_dispatch(result.take_dispatch()?))
    }
    /// Returns Excel's current saved-state flag.
    pub fn saved(&self) -> Result<bool, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.saved"), false),
            vec![],
        )?
        .as_bool()
        .ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType { vartype: 0 },
        ))
    }
    /// Sets Excel's saved-state flag without saving a file.
    pub fn set_saved(&self, value: bool) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.saved"), true),
            OwnedVariant::bool(value),
        )?;
        Ok(())
    }
    /// Returns the workbook's current full Excel name.
    pub fn full_name(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.fullname"), false),
            vec![],
        )?
        .as_string()
    }
    /// Returns the workbook's directory path reported by Excel.
    pub fn path(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.path"), false),
            vec![],
        )?
        .as_string()
    }
    /// Returns the workbook's current Excel file-format value.
    ///
    /// The registered type library describes this as an integer enum. Current
    /// Excel runtime evidence also returns an exact integral `VT_R8`, which is
    /// accepted without changing the public newtype.
    pub fn file_format(&self) -> Result<XlFileFormat, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.fileformat"), false),
            vec![],
        )?;
        if let Some(value) = result.as_i32() {
            return Ok(XlFileFormat::from_raw(value));
        }
        if let Some(value) = result.as_f64()
            && value.is_finite()
            && value.fract() == 0.0
            && value >= i32::MIN as f64
            && value <= i32::MAX as f64
        {
            return Ok(XlFileFormat::from_raw(value as i32));
        }
        Err(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType {
                vartype: result.vt(),
            },
        ))
    }
    /// Returns whether Excel has opened the workbook read-only.
    pub fn read_only(&self) -> Result<bool, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.readonly"), false),
            vec![],
        )?;
        result.as_bool().ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType {
                vartype: result.vt(),
            },
        ))
    }
    /// Saves the workbook at its current file identity.
    pub fn save(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.save"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Saves the workbook under a caller-supplied Windows path and options.
    pub fn save_as<P: AsRef<Path>>(
        &self,
        filename: P,
        options: WorkbookSaveAsOptions<'_>,
    ) -> Result<(), ExcelComError> {
        let arguments = save_as_arguments(filename.as_ref(), options)?;
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.saveas-3174"), false),
            arguments,
            false,
        )?;
        Ok(())
    }
    /// Saves a copy without changing this workbook's current file identity.
    pub fn save_copy_as<P: AsRef<Path>>(&self, filename: P) -> Result<(), ExcelComError> {
        let mut arguments = PositionalArguments::new();
        arguments.push_result(path_bstr(filename.as_ref()))?;
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.savecopyas"), false),
            arguments.into_inner(),
            false,
        )?;
        Ok(())
    }
    /// Closes this workbook with explicit Excel close options.
    pub fn close(self, options: WorkbookCloseOptions<'_>) -> Result<(), ExcelComError> {
        let arguments = close_arguments(options)?;
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.close"), false),
            arguments,
            false,
        )?;
        Ok(())
    }
    /// Closes this workbook while explicitly declining to save changes.
    pub fn close_without_saving(self) -> Result<(), ExcelComError> {
        self.close(WorkbookCloseOptions {
            save_changes: SaveChanges::Discard,
            ..WorkbookCloseOptions::new()
        })
    }
}

fn save_as_arguments(
    filename: &Path,
    options: WorkbookSaveAsOptions<'_>,
) -> Result<Vec<OwnedVariant>, ExcelComError> {
    let mut arguments = PositionalArguments::new();
    arguments.push_result(path_bstr(filename))?;
    arguments.push_optional(
        options
            .file_format
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    push_optional_text(&mut arguments, options.password)?;
    push_optional_text(&mut arguments, options.write_res_password)?;
    arguments.push_optional(options.read_only_recommended.map(OwnedVariant::bool));
    arguments.push_optional(options.create_backup.map(OwnedVariant::bool));
    arguments.push_optional(
        options
            .access_mode
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    arguments.push_optional(
        options
            .conflict_resolution
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    arguments.push_optional(options.add_to_mru.map(OwnedVariant::bool));
    arguments.push_optional(options.text_codepage.map(OwnedVariant::i32));
    arguments.push_optional(options.text_visual_layout.map(OwnedVariant::i32));
    arguments.push_optional(options.local.map(OwnedVariant::bool));
    arguments.push_optional(None);
    Ok(arguments.into_inner())
}

fn close_arguments(options: WorkbookCloseOptions<'_>) -> Result<Vec<OwnedVariant>, ExcelComError> {
    let mut arguments = PositionalArguments::new();
    match options.save_changes {
        SaveChanges::Prompt => arguments.push_optional(None),
        SaveChanges::Save => arguments.push_required(OwnedVariant::bool(true)),
        SaveChanges::Discard => arguments.push_required(OwnedVariant::bool(false)),
    }
    match options.filename {
        Some(value) => arguments.push_result(path_bstr(value))?,
        None => arguments.push_optional(None),
    }
    arguments.push_optional(options.route_workbook.map(OwnedVariant::bool));
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
    fn save_as_keeps_its_13th_workidentity_position_missing() {
        let values = save_as_arguments(
            Path::new("output.xlsx"),
            WorkbookSaveAsOptions {
                file_format: Some(XlFileFormat::OPEN_XML_WORKBOOK),
                access_mode: Some(crate::XlSaveAsAccessMode::EXCLUSIVE),
                local: Some(true),
                ..WorkbookSaveAsOptions::new()
            },
        )
        .expect("save arguments");
        assert_eq!(values.len(), 13);
        assert_eq!(values[1].as_i32(), Some(51));
        assert_eq!(values[6].as_i32(), Some(3));
        assert_eq!(values[11].as_bool(), Some(true));
        assert_eq!(values[12].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }

    #[test]
    fn close_uses_missing_for_prompt_and_preserves_optional_positions() {
        let values = close_arguments(WorkbookCloseOptions::new()).expect("close arguments");
        assert_eq!(values.len(), 3);
        assert_eq!(values[0].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[1].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[2].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }

    #[test]
    fn save_as_option_debug_redacts_passwords() {
        let text = format!(
            "{:?}",
            WorkbookSaveAsOptions {
                password: Some("redaction-input-one"),
                write_res_password: Some("redaction-input-two"),
                ..WorkbookSaveAsOptions::new()
            }
        );
        assert!(text.contains("REDACTED"));
        assert!(!text.contains("redaction-input-one"));
        assert!(!text.contains("redaction-input-two"));
    }
}
