//! Protection options and print/export/page-break wrappers.
#![allow(missing_docs)]

use super::*;

/// Password and permission positions accepted by `Worksheet.Protect`.
#[derive(Clone, Default, PartialEq)]
pub struct WorksheetProtectOptions<'a> {
    pub password: Option<&'a str>,
    pub drawing_objects: Option<bool>,
    pub contents: Option<bool>,
    pub scenarios: Option<bool>,
    pub user_interface_only: Option<bool>,
    pub allow_formatting_cells: Option<bool>,
    pub allow_formatting_columns: Option<bool>,
    pub allow_formatting_rows: Option<bool>,
    pub allow_inserting_columns: Option<bool>,
    pub allow_inserting_rows: Option<bool>,
    pub allow_inserting_hyperlinks: Option<bool>,
    pub allow_deleting_columns: Option<bool>,
    pub allow_deleting_rows: Option<bool>,
    pub allow_sorting: Option<bool>,
    pub allow_filtering: Option<bool>,
    pub allow_using_pivot_tables: Option<bool>,
}
impl Debug for WorksheetProtectOptions<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WorksheetProtectOptions")
            .field("password", &self.password.map(|_| "REDACTED"))
            .field("drawing_objects", &self.drawing_objects)
            .field("contents", &self.contents)
            .field("scenarios", &self.scenarios)
            .field("user_interface_only", &self.user_interface_only)
            .field("allow_formatting_cells", &self.allow_formatting_cells)
            .field("allow_formatting_columns", &self.allow_formatting_columns)
            .field("allow_formatting_rows", &self.allow_formatting_rows)
            .field("allow_inserting_columns", &self.allow_inserting_columns)
            .field("allow_inserting_rows", &self.allow_inserting_rows)
            .field(
                "allow_inserting_hyperlinks",
                &self.allow_inserting_hyperlinks,
            )
            .field("allow_deleting_columns", &self.allow_deleting_columns)
            .field("allow_deleting_rows", &self.allow_deleting_rows)
            .field("allow_sorting", &self.allow_sorting)
            .field("allow_filtering", &self.allow_filtering)
            .field("allow_using_pivot_tables", &self.allow_using_pivot_tables)
            .finish()
    }
}

/// Password and permission positions accepted by `Workbook.Protect`.
#[derive(Clone, Default, PartialEq)]
pub struct WorkbookProtectOptions<'a> {
    pub password: Option<&'a str>,
    pub structure: Option<bool>,
    pub windows: Option<bool>,
}
impl Debug for WorkbookProtectOptions<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WorkbookProtectOptions")
            .field("password", &self.password.map(|_| "REDACTED"))
            .field("structure", &self.structure)
            .field("windows", &self.windows)
            .finish()
    }
}

/// Typed options for Excel's `PrintOut` method.
#[derive(Clone, Default, PartialEq)]
pub struct PrintOutOptions<'a> {
    pub from: Option<usize>,
    pub to: Option<usize>,
    pub copies: Option<usize>,
    pub preview: Option<bool>,
    pub active_printer: Option<&'a str>,
    pub print_to_file: Option<bool>,
    pub collate: Option<bool>,
    pub pr_to_file_name: Option<&'a str>,
    pub ignore_print_areas: Option<bool>,
}
impl Debug for PrintOutOptions<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("PrintOutOptions")
            .field("from", &self.from)
            .field("to", &self.to)
            .field("copies", &self.copies)
            .field("preview", &self.preview)
            .field("active_printer", &self.active_printer.map(|_| "REDACTED"))
            .field("print_to_file", &self.print_to_file)
            .field("collate", &self.collate)
            .field("pr_to_file_name", &self.pr_to_file_name.map(|_| "REDACTED"))
            .field("ignore_print_areas", &self.ignore_print_areas)
            .finish()
    }
}

/// Typed optional positions for `ExportAsFixedFormat`.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct FixedFormatOptions<'a> {
    pub output: Option<&'a Path>,
    pub quality: Option<FixedFormatQuality>,
    pub include_doc_properties: Option<bool>,
    pub ignore_print_areas: Option<bool>,
    pub from: Option<usize>,
    pub to: Option<usize>,
    pub open_after_publish: Option<bool>,
}

/// A horizontal page break.
pub struct HPageBreak {
    inner: DispatchObject,
}
impl Debug for HPageBreak {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("HPageBreak").field(&self.inner).finish()
    }
}
impl Clone for HPageBreak {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl HPageBreak {
    pub(super) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "HPageBreak",
            },
        }
    }
    pub fn location(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.hpagebreak.location")
    }
    pub fn set_location(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.hpagebreak.location",
            OwnedVariant::dispatch_borrowed(&range.dispatch_object().dispatch),
        )
    }
    pub fn break_type(&self) -> Result<PageBreakType, ExcelComError> {
        Ok(PageBreakType::from_raw(get_i32(
            &self.inner,
            "excel.hpagebreak.type",
            "HPageBreak.Type",
        )?))
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.hpagebreak.delete", vec![])
    }
}

/// A vertical page break.
pub struct VPageBreak {
    inner: DispatchObject,
}
impl Debug for VPageBreak {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("VPageBreak").field(&self.inner).finish()
    }
}
impl Clone for VPageBreak {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl VPageBreak {
    pub(super) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "VPageBreak",
            },
        }
    }
    pub fn location(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.vpagebreak.location")
    }
    pub fn set_location(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.vpagebreak.location",
            OwnedVariant::dispatch_borrowed(&range.dispatch_object().dispatch),
        )
    }
    pub fn break_type(&self) -> Result<PageBreakType, ExcelComError> {
        Ok(PageBreakType::from_raw(get_i32(
            &self.inner,
            "excel.vpagebreak.type",
            "VPageBreak.Type",
        )?))
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.vpagebreak.delete", vec![])
    }
}

const HPAGE_BREAKS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "HPageBreaks",
    count: MemberId::new("excel.hpagebreaks.count"),
    item: MemberId::new("excel.hpagebreaks.item"),
    new_enum: MemberId::new("excel.hpagebreaks.newenum"),
};
const VPAGE_BREAKS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "VPageBreaks",
    count: MemberId::new("excel.vpagebreaks.count"),
    item: MemberId::new("excel.vpagebreaks.item"),
    new_enum: MemberId::new("excel.vpagebreaks.newenum"),
};

/// Horizontal page-break collection for one worksheet.
pub struct HPageBreaks {
    inner: DispatchObject,
}
impl Debug for HPageBreaks {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("HPageBreaks").field(&self.inner).finish()
    }
}
impl Clone for HPageBreaks {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl HPageBreaks {
    pub(super) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "HPageBreaks",
            },
        }
    }
    pub fn count(&self) -> Result<i32, ExcelComError> {
        count_i32(&self.inner, HPAGE_BREAKS_DESCRIPTOR)
    }
    pub fn item_by_index(&self, index: usize) -> Result<HPageBreak, ExcelComError> {
        Ok(HPageBreak::from_dispatch(item_by_index(
            &self.inner,
            HPAGE_BREAKS_DESCRIPTOR,
            index,
        )?))
    }
    pub fn add(&self, before: &Range) -> Result<HPageBreak, ExcelComError> {
        add_page_break(
            &self.inner,
            "excel.hpagebreaks.add",
            before,
            HPageBreak::from_dispatch,
        )
    }
}

/// Vertical page-break collection for one worksheet.
pub struct VPageBreaks {
    inner: DispatchObject,
}
impl Debug for VPageBreaks {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("VPageBreaks").field(&self.inner).finish()
    }
}
impl Clone for VPageBreaks {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl VPageBreaks {
    pub(super) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "VPageBreaks",
            },
        }
    }
    pub fn count(&self) -> Result<i32, ExcelComError> {
        count_i32(&self.inner, VPAGE_BREAKS_DESCRIPTOR)
    }
    pub fn item_by_index(&self, index: usize) -> Result<VPageBreak, ExcelComError> {
        Ok(VPageBreak::from_dispatch(item_by_index(
            &self.inner,
            VPAGE_BREAKS_DESCRIPTOR,
            index,
        )?))
    }
    pub fn add(&self, before: &Range) -> Result<VPageBreak, ExcelComError> {
        add_page_break(
            &self.inner,
            "excel.vpagebreaks.add",
            before,
            VPageBreak::from_dispatch,
        )
    }
}
