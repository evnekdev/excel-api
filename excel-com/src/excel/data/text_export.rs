use std::path::Path;

use crate::ExcelComError;
use crate::excel::{Workbook, WorkbookSaveAsOptions, XlFileFormat};

use super::{TextExportOptions, TextFileFormat};

impl<'a> TextExportOptions<'a> {
    /// Builds a UTF-8 CSV export request. Excel normally exports only the active worksheet.
    pub fn csv_utf8(path: &'a Path) -> Self {
        Self {
            path,
            format: TextFileFormat::CSV_UTF8,
            local: None,
            create_backup: Some(false),
        }
    }
    /// Builds a Unicode tab-delimited export request. Excel normally exports only the active worksheet.
    pub fn tsv_unicode(path: &'a Path) -> Self {
        Self {
            path,
            format: TextFileFormat::UNICODE_TEXT,
            local: None,
            create_backup: Some(false),
        }
    }
}

impl Workbook {
    /// Saves the active worksheet in an Excel text format through the existing `SaveAs` implementation.
    ///
    /// Text export normally uses displayed/calculated cell values and Excel decides locale and
    /// encoding behavior. It is not a Rust CSV writer.
    ///
    /// ```no_run
    /// # fn example(workbook: &excel_com::Workbook, output: &std::path::Path) -> Result<(), excel_com::ExcelComError> {
    /// use excel_com::TextExportOptions;
    /// workbook.save_as_text(&TextExportOptions::csv_utf8(output))?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save_as_text(&self, options: &TextExportOptions<'_>) -> Result<(), ExcelComError> {
        self.save_as(
            options.path,
            WorkbookSaveAsOptions {
                file_format: Some(XlFileFormat::from_raw(options.format.raw())),
                create_backup: options.create_backup,
                local: options.local,
                ..WorkbookSaveAsOptions::new()
            },
        )
    }
}
