//! Private COM conversion and optional-argument helpers shared by presentation modules.

use super::*;

pub(super) fn count_i32(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
) -> Result<i32, ExcelComError> {
    i32::try_from(collection_count(target, descriptor)?).map_err(|_| ExcelComError::Unsupported {
        detail: "collection Count exceeds i32",
    })
}
pub(super) fn sheet_from_dispatch(dispatch: ComPtr<Dispatch>) -> Result<Sheet, ExcelComError> {
    let object = DispatchObject {
        dispatch,
        kind: "Sheet",
    };
    let sheet_type = get_i32(&object, "excel.worksheet.type", "Sheet.Type")?;
    if sheet_type == -4167 {
        Ok(Sheet::Worksheet(Worksheet::from_dispatch(object.dispatch)))
    } else {
        Ok(Sheet::Other(SheetObject::from_dispatch(object.dispatch)))
    }
}
pub(super) fn get_object<T>(
    target: &DispatchObject,
    id: &'static str,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    Ok(from(result.take_dispatch()?))
}
pub(super) fn get_sheet(target: &DispatchObject, id: &'static str) -> Result<Sheet, ExcelComError> {
    let mut result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    sheet_from_dispatch(result.take_dispatch()?)
}
pub(super) fn get_range(target: &DispatchObject, id: &'static str) -> Result<Range, ExcelComError> {
    get_object(target, id, Range::from_dispatch)
}
pub(super) fn get_string(
    target: &DispatchObject,
    id: &'static str,
) -> Result<String, ExcelComError> {
    property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?.as_string()
}
pub(super) fn optional_string(
    target: &DispatchObject,
    id: &'static str,
) -> Result<Option<String>, ExcelComError> {
    let value = get_string(target, id)?;
    Ok((!value.is_empty()).then_some(value))
}
pub(super) fn get_bool(target: &DispatchObject, id: &'static str) -> Result<bool, ExcelComError> {
    let result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    result.as_bool().ok_or(ExcelComError::Conversion(
        ConversionError::UnsupportedVariantType {
            vartype: result.vt(),
        },
    ))
}
pub(super) fn integer_variant(value: &OwnedVariant) -> Option<i32> {
    value.as_i32().or_else(|| {
        value
            .as_f64()
            .filter(|v| {
                v.is_finite() && v.fract() == 0.0 && *v >= i32::MIN as f64 && *v <= i32::MAX as f64
            })
            .map(|v| v as i32)
    })
}
pub(super) fn get_i32(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<i32, ExcelComError> {
    let result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    integer_variant(&result).ok_or(ExcelComError::Unsupported { detail })
}
pub(super) fn get_f64(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<f64, ExcelComError> {
    let result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    result
        .as_f64()
        .or_else(|| result.as_i32().map(f64::from))
        .ok_or(ExcelComError::Unsupported { detail })
}
pub(super) fn put(
    target: &DispatchObject,
    id: &'static str,
    value: OwnedVariant,
) -> Result<(), ExcelComError> {
    let _ = property_put(&target.dispatch, member(MemberId::new(id), true), value)?;
    Ok(())
}
pub(super) fn call(
    target: &DispatchObject,
    id: &'static str,
    arguments: Vec<OwnedVariant>,
) -> Result<(), ExcelComError> {
    let _ = invoke(
        &target.dispatch,
        member(MemberId::new(id), false),
        arguments,
        false,
    )?;
    Ok(())
}
pub(super) fn one_based(value: usize, detail: &'static str) -> Result<OwnedVariant, ExcelComError> {
    if value == 0 {
        return Err(ExcelComError::Unsupported { detail });
    }
    Ok(OwnedVariant::i32(
        i32::try_from(value).map_err(|_| ExcelComError::Unsupported { detail })?,
    ))
}
pub(super) fn finite(value: f64) -> Result<(), ExcelComError> {
    if value.is_finite() {
        Ok(())
    } else {
        Err(ExcelComError::Conversion(ConversionError::NonFiniteNumber))
    }
}
pub(super) fn nonnegative(value: f64) -> Result<(), ExcelComError> {
    finite(value)?;
    if value >= 0.0 {
        Ok(())
    } else {
        Err(ExcelComError::Unsupported {
            detail: "page margin must be nonnegative",
        })
    }
}
pub(super) fn optional_positive(
    value: Option<usize>,
    detail: &'static str,
) -> Result<OwnedVariant, ExcelComError> {
    match value {
        Some(value) => one_based(value, detail),
        None => Ok(OwnedVariant::bool(false)),
    }
}
pub(super) fn push_optional_text(
    arguments: &mut PositionalArguments,
    value: Option<&str>,
) -> Result<(), ExcelComError> {
    match value {
        Some(value) => arguments.push_result(text_bstr(value)),
        None => {
            arguments.push_optional(None);
            Ok(())
        }
    }
}
pub(super) fn page_zoom_get(
    target: &DispatchObject,
    id: &'static str,
) -> Result<PageZoom, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    if value.as_bool() == Some(false) {
        Ok(PageZoom::Automatic)
    } else {
        integer_variant(&value)
            .map(PageZoom::Percent)
            .ok_or(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType {
                    vartype: value.vt(),
                },
            ))
    }
}
pub(super) fn page_zoom_put(
    target: &DispatchObject,
    id: &'static str,
    value: PageZoom,
) -> Result<(), ExcelComError> {
    match value {
        PageZoom::Automatic => put(target, id, OwnedVariant::bool(false)),
        PageZoom::Percent(value) if value > 0 => put(target, id, OwnedVariant::i32(value)),
        PageZoom::Percent(_) => Err(ExcelComError::Unsupported {
            detail: "zoom percentage must be positive",
        }),
    }
}
pub(super) fn page_fit_dimension_get(
    target: &DispatchObject,
    id: &'static str,
) -> Result<Option<usize>, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    if value.as_bool() == Some(false) {
        return Ok(None);
    }
    let value = integer_variant(&value).ok_or(ExcelComError::Conversion(
        ConversionError::UnsupportedVariantType {
            vartype: value.vt(),
        },
    ))?;
    usize::try_from(value)
        .ok()
        .filter(|value| *value > 0)
        .ok_or(ExcelComError::Unsupported {
            detail: "PageSetup fit-to-pages dimension was not positive",
        })
        .map(Some)
}
pub(super) fn add_page_break<T>(
    target: &DispatchObject,
    id: &'static str,
    before: &Range,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_object(before.dispatch_object());
    let mut result = invoke(
        &target.dispatch,
        member(MemberId::new(id), false),
        a.into_inner(),
        false,
    )?;
    Ok(from(result.take_dispatch()?))
}
pub(super) fn worksheet_copy_move(
    source: &Worksheet,
    id: &'static str,
    destination: SheetDestination<'_>,
) -> Result<(), ExcelComError> {
    let mut a = PositionalArguments::new();
    match destination {
        SheetDestination::Before(value) => {
            a.push_object(value.dispatch_object());
            a.push_optional(None);
        }
        SheetDestination::After(value) => {
            a.push_optional(None);
            a.push_object(value.dispatch_object());
        }
        SheetDestination::NewWorkbook => {
            a.push_optional(None);
            a.push_optional(None);
        }
    };
    call(source.dispatch_object(), id, a.into_inner())
}
pub(super) fn worksheet_protect(
    target: &DispatchObject,
    options: &WorksheetProtectOptions<'_>,
) -> Result<(), ExcelComError> {
    let mut a = PositionalArguments::new();
    push_optional_text(&mut a, options.password)?;
    for value in [
        options.drawing_objects,
        options.contents,
        options.scenarios,
        options.user_interface_only,
        options.allow_formatting_cells,
        options.allow_formatting_columns,
        options.allow_formatting_rows,
        options.allow_inserting_columns,
        options.allow_inserting_rows,
        options.allow_inserting_hyperlinks,
        options.allow_deleting_columns,
        options.allow_deleting_rows,
        options.allow_sorting,
        options.allow_filtering,
        options.allow_using_pivot_tables,
    ] {
        a.push_optional(value.map(OwnedVariant::bool));
    }
    call(target, "excel.worksheet.protect-2029", a.into_inner())
}
pub(super) fn print_out(
    target: &DispatchObject,
    id: &'static str,
    options: &PrintOutOptions<'_>,
) -> Result<(), ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_optional(
        options
            .from
            .map(|v| one_based(v, "PrintOut.From"))
            .transpose()?,
    );
    a.push_optional(
        options
            .to
            .map(|v| one_based(v, "PrintOut.To"))
            .transpose()?,
    );
    a.push_optional(
        options
            .copies
            .map(|v| one_based(v, "PrintOut.Copies"))
            .transpose()?,
    );
    a.push_optional(options.preview.map(OwnedVariant::bool));
    push_optional_text(&mut a, options.active_printer)?;
    a.push_optional(options.print_to_file.map(OwnedVariant::bool));
    a.push_optional(options.collate.map(OwnedVariant::bool));
    push_optional_text(&mut a, options.pr_to_file_name)?;
    a.push_optional(options.ignore_print_areas.map(OwnedVariant::bool));
    call(target, id, a.into_inner())
}
pub(super) fn fixed_format(
    target: &DispatchObject,
    id: &'static str,
    format: FixedFormatType,
    options: &FixedFormatOptions<'_>,
) -> Result<(), ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_required(OwnedVariant::i32(format.raw()));
    match options.output {
        Some(path) => a.push_result(path_bstr(path))?,
        None => a.push_optional(None),
    };
    a.push_optional(options.quality.map(|v| OwnedVariant::i32(v.raw())));
    a.push_optional(options.include_doc_properties.map(OwnedVariant::bool));
    a.push_optional(options.ignore_print_areas.map(OwnedVariant::bool));
    a.push_optional(
        options
            .from
            .map(|v| one_based(v, "ExportAsFixedFormat.From"))
            .transpose()?,
    );
    a.push_optional(
        options
            .to
            .map(|v| one_based(v, "ExportAsFixedFormat.To"))
            .transpose()?,
    );
    a.push_optional(options.open_after_publish.map(OwnedVariant::bool));
    a.push_optional(None);
    call(target, id, a.into_inner())
}
