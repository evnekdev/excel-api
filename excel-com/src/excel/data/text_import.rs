use std::path::Path;

use crate::ExcelComError;
use crate::automation::{OwnedVariant, PositionalArguments, invoke};
use crate::excel::{Workbook, Workbooks};
use crate::object_model::{MemberId, member};

use super::helpers::{delimiter_arguments, field_info, one_based, text_path};
use super::{OpenTextOptions, TextDelimiter, TextParsingType, TextQualifier, TextToColumnsOptions};

impl<'a> OpenTextOptions<'a> {
    /// Builds a locale-explicit comma-delimited import request. Excel still owns type inference.
    pub fn csv(path: &'a Path) -> Self {
        Self {
            path,
            origin: None,
            start_row: None,
            parsing_type: TextParsingType::DELIMITED,
            text_qualifier: Some(TextQualifier::DOUBLE_QUOTE),
            consecutive_delimiters: Some(false),
            delimiter: Some(TextDelimiter::Comma),
            columns: Vec::new(),
            decimal_separator: None,
            thousands_separator: None,
            trailing_minus_numbers: None,
            local: None,
        }
    }
    /// Builds a locale-explicit tab-delimited import request. Excel still owns type inference.
    pub fn tsv(path: &'a Path) -> Self {
        Self {
            delimiter: Some(TextDelimiter::Tab),
            ..Self::csv(path)
        }
    }
}

impl Workbooks {
    /// Opens a text file through Excel's `OpenText` import engine and creates a workbook.
    ///
    /// Delimiters, quotes, date and number inference remain Excel-owned. This does not create a
    /// persistent QueryTable or external connection.
    ///
    /// ```no_run
    /// # fn example(application: &excel_com::Application, path: &std::path::Path) -> Result<(), excel_com::ExcelComError> {
    /// use excel_com::OpenTextOptions;
    /// let workbook = application.workbooks()?.open_text(&OpenTextOptions::csv(path))?;
    /// # drop(workbook); Ok(())
    /// # }
    /// ```
    pub fn open_text(&self, options: &OpenTextOptions<'_>) -> Result<Workbook, ExcelComError> {
        let mut result = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.workbooks.opentext-1924"), false),
            open_text_arguments(options)?,
            false,
        )?;
        Ok(Workbook::from_dispatch(result.take_dispatch()?))
    }
}

impl crate::excel::Range {
    /// Splits this Range through Excel's `TextToColumns` engine, modifying worksheet cells.
    /// Excel owns parsing, locale conversion, and unsupported-shape errors.
    pub fn text_to_columns(&self, options: &TextToColumnsOptions<'_>) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.texttocolumns"), false),
            text_to_columns_arguments(options)?,
            false,
        )?;
        Ok(())
    }
}

fn open_text_arguments(options: &OpenTextOptions<'_>) -> Result<Vec<OwnedVariant>, ExcelComError> {
    let (tab, semicolon, comma, space, other_enabled, other) =
        delimiter_arguments(options.delimiter.as_ref())?;
    let mut args = PositionalArguments::new();
    args.push_result(text_path(options.path))?;
    args.push_optional(options.origin.map(|value| OwnedVariant::i32(value.raw())));
    args.push_optional(
        options
            .start_row
            .map(|value| one_based(value, "OpenText start_row is one-based"))
            .transpose()?
            .map(OwnedVariant::i32),
    );
    args.push_required(OwnedVariant::i32(options.parsing_type.raw()));
    args.push_optional(
        options
            .text_qualifier
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    args.push_optional(options.consecutive_delimiters.map(OwnedVariant::bool));
    args.push_optional(tab.map(OwnedVariant::bool));
    args.push_optional(semicolon.map(OwnedVariant::bool));
    args.push_optional(comma.map(OwnedVariant::bool));
    args.push_optional(space.map(OwnedVariant::bool));
    args.push_optional(other_enabled.map(OwnedVariant::bool));
    args.push_optional(other);
    args.push_optional(field_info(&options.columns, options.parsing_type)?);
    // OpenText exposes TextVisualLayout between FieldInfo and decimal handling;
    // Prompt 18 deliberately leaves visual-layout selection out of its public API.
    args.push_optional(None);
    args.push_optional(
        options
            .decimal_separator
            .map(|value| OwnedVariant::bstr(&separator(value, "decimal")?))
            .transpose()?,
    );
    args.push_optional(
        options
            .thousands_separator
            .map(|value| OwnedVariant::bstr(&separator(value, "thousands")?))
            .transpose()?,
    );
    args.push_optional(options.trailing_minus_numbers.map(OwnedVariant::bool));
    args.push_optional(options.local.map(OwnedVariant::bool));
    Ok(args.into_inner())
}

fn text_to_columns_arguments(
    options: &TextToColumnsOptions<'_>,
) -> Result<Vec<OwnedVariant>, ExcelComError> {
    let (tab, semicolon, comma, space, other_enabled, other) =
        delimiter_arguments(options.delimiter.as_ref())?;
    let mut args = PositionalArguments::new();
    args.push_optional_object(
        options
            .destination
            .map(crate::excel::Range::dispatch_object),
    );
    args.push_required(OwnedVariant::i32(options.parsing_type.raw()));
    args.push_optional(
        options
            .text_qualifier
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    args.push_optional(options.consecutive_delimiters.map(OwnedVariant::bool));
    args.push_optional(tab.map(OwnedVariant::bool));
    args.push_optional(semicolon.map(OwnedVariant::bool));
    args.push_optional(comma.map(OwnedVariant::bool));
    args.push_optional(space.map(OwnedVariant::bool));
    args.push_optional(other_enabled.map(OwnedVariant::bool));
    args.push_optional(other);
    args.push_optional(field_info(&options.columns, options.parsing_type)?);
    args.push_optional(
        options
            .decimal_separator
            .map(|value| OwnedVariant::bstr(&separator(value, "decimal")?))
            .transpose()?,
    );
    args.push_optional(
        options
            .thousands_separator
            .map(|value| OwnedVariant::bstr(&separator(value, "thousands")?))
            .transpose()?,
    );
    args.push_optional(options.trailing_minus_numbers.map(OwnedVariant::bool));
    Ok(args.into_inner())
}

fn separator(value: char, name: &'static str) -> Result<String, ExcelComError> {
    if value == '\0' {
        return Err(ExcelComError::Unsupported {
            detail: if name == "decimal" {
                "decimal separator cannot be NUL"
            } else {
                "thousands separator cannot be NUL"
            },
        });
    }
    Ok(value.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn open_text_keeps_all_sixteen_positions() {
        let options = OpenTextOptions::csv(Path::new("sample.csv"));
        let values = open_text_arguments(&options).expect("arguments");
        assert_eq!(values.len(), 18);
        assert_eq!(values[3].as_i32(), Some(TextParsingType::DELIMITED.raw()));
        assert_eq!(values[8].as_bool(), Some(true));
        assert_eq!(values[10].as_bool(), Some(false));
        assert_eq!(values[17].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }
}
