use crate::ExcelComError;
use crate::automation::{
    EnumVariant, OwnedVariant, PositionalArguments, enumerated_dispatch, invoke, property_get,
};
use crate::excel::collection::{
    CollectionDescriptor, count as collection_count, enumerator, item_by_index, item_by_name,
};
use crate::excel::{DispatchObject, Workbook, WorkbookOpenOptions};
use crate::internal::{ComPtr, Dispatch, path_bstr};
use crate::object_model::{MemberId, member};
use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;
use std::path::Path;

const DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Workbooks",
    count: MemberId::new("excel.workbooks.count"),
    item: MemberId::new("excel.workbooks.item"),
    new_enum: MemberId::new("excel.workbooks.newenum"),
};

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
    pub(crate) fn dispatch_object(&self) -> &DispatchObject {
        &self.inner
    }
    /// Returns the number of open workbooks.
    pub fn count(&self) -> Result<i32, ExcelComError> {
        i32::try_from(collection_count(&self.inner, DESCRIPTOR)?).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "Workbooks.Count exceeds i32",
            }
        })
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
    /// Opens a workbook with only Excel's required filename argument.
    ///
    /// This deliberately differs from [`Self::open_default`]: it omits all
    /// optional positions from the COM invocation instead of supplying them as
    /// explicit `Missing` values. Some legacy workbooks with data-model or
    /// external-link metadata accept the normal Office automation form but
    /// reject a fully expanded `Missing` argument list.
    pub fn open_minimal<P: AsRef<Path>>(&self, filename: P) -> Result<Workbook, ExcelComError> {
        let mut result = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbooks.open-1923"), false),
            open_minimal_arguments(filename.as_ref())?,
            false,
        )?;
        Ok(Workbook::from_dispatch(result.take_dispatch()?))
    }
    /// Opens a workbook with the safe prefix of Excel's `Workbooks.Open`
    /// signature, without adding trailing `Missing` positions.
    ///
    /// `UpdateLinks := 0` prevents external-link updates, while `read_only`
    /// selects whether the caller needs to mutate the workbook. This matches
    /// the documented Office automation shape and avoids both the rejected
    /// one-argument form and legacy files that reject a fully expanded
    /// `Missing` argument list.
    pub fn open_without_link_updates<P: AsRef<Path>>(
        &self,
        filename: P,
        read_only: bool,
    ) -> Result<Workbook, ExcelComError> {
        let mut result = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbooks.open-1923"), false),
            open_without_link_updates_arguments(filename.as_ref(), read_only)?,
            false,
        )?;
        Ok(Workbook::from_dispatch(result.take_dispatch()?))
    }
    /// Returns the one-based workbook at `index`.
    pub fn item_by_index(&self, index: usize) -> Result<Workbook, ExcelComError> {
        Ok(Workbook::from_dispatch(item_by_index(
            &self.inner,
            DESCRIPTOR,
            index,
        )?))
    }
    /// Returns the workbook selected by its current name.
    pub fn item_by_name(&self, name: &str) -> Result<Workbook, ExcelComError> {
        Ok(Workbook::from_dispatch(item_by_name(
            &self.inner,
            DESCRIPTOR,
            name,
        )?))
    }
    /// Iterates workbook objects in Excel's `_NewEnum` order.
    pub fn iter(&self) -> Result<WorkbooksIter, ExcelComError> {
        Ok(WorkbooksIter {
            enumerator: enumerator(&self.inner, DESCRIPTOR)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Typed, single-pass workbook collection iterator.
pub struct WorkbooksIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for WorkbooksIter {
    type Item = Result<Workbook, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(
                    enumerated_dispatch(&mut value, "Workbooks", index)
                        .map(Workbook::from_dispatch),
                )
            }
            Ok(None) => {
                self.terminal = true;
                None
            }
            Err(error) => {
                self.terminal = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for WorkbooksIter {}

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

fn open_minimal_arguments(filename: &Path) -> Result<Vec<OwnedVariant>, ExcelComError> {
    Ok(vec![path_bstr(filename)?])
}

fn open_without_link_updates_arguments(
    filename: &Path,
    read_only: bool,
) -> Result<Vec<OwnedVariant>, ExcelComError> {
    Ok(vec![
        path_bstr(filename)?,
        OwnedVariant::i32(0),
        OwnedVariant::bool(read_only),
    ])
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
    fn minimal_open_omits_optional_positions() {
        let values =
            open_minimal_arguments(Path::new("relative folder/\u{8cc7}\u{6599} book.xlsx"))
                .expect("path argument");
        assert_eq!(values.len(), 1);
    }

    #[test]
    fn safe_open_prefix_disables_link_updates_and_sets_read_only() {
        let values = open_without_link_updates_arguments(
            Path::new("relative folder/\u{8cc7}\u{6599} book.xlsx"),
            true,
        )
        .expect("safe-open arguments");
        assert_eq!(values.len(), 3);
        assert_eq!(values[1].as_i32(), Some(0));
        assert_eq!(values[2].as_bool(), Some(true));
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
