//! Page layout, print-title, and header/footer settings.
#![allow(missing_docs)]

use super::*;

/// Page layout, print-title, and header/footer settings for a worksheet.
pub struct PageSetup {
    inner: DispatchObject,
}
impl Debug for PageSetup {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PageSetup").field(&self.inner).finish()
    }
}
impl Clone for PageSetup {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PageSetup {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PageSetup",
            },
        }
    }
    pub fn orientation(&self) -> Result<PageOrientation, ExcelComError> {
        Ok(PageOrientation::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.orientation",
            "PageSetup.Orientation",
        )?))
    }
    pub fn set_orientation(&self, value: PageOrientation) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.orientation",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn paper_size(&self) -> Result<PaperSize, ExcelComError> {
        Ok(PaperSize::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.papersize",
            "PageSetup.PaperSize",
        )?))
    }
    pub fn set_paper_size(&self, value: PaperSize) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.papersize",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn order(&self) -> Result<PrintOrder, ExcelComError> {
        Ok(PrintOrder::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.order",
            "PageSetup.Order",
        )?))
    }
    pub fn set_order(&self, value: PrintOrder) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.order",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn zoom(&self) -> Result<PageZoom, ExcelComError> {
        page_zoom_get(&self.inner, "excel.pagesetup.zoom")
    }
    pub fn set_zoom(&self, value: PageZoom) -> Result<(), ExcelComError> {
        page_zoom_put(&self.inner, "excel.pagesetup.zoom", value)
    }
    pub fn set_fit_to_pages(&self, value: PageFit) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.fittopageswide",
            optional_positive(value.wide, "PageSetup.FitToPagesWide")?,
        )?;
        put(
            &self.inner,
            "excel.pagesetup.fittopagestall",
            optional_positive(value.tall, "PageSetup.FitToPagesTall")?,
        )
    }
    pub fn fit_to_pages(&self) -> Result<PageFit, ExcelComError> {
        Ok(PageFit {
            wide: page_fit_dimension_get(&self.inner, "excel.pagesetup.fittopageswide")?,
            tall: page_fit_dimension_get(&self.inner, "excel.pagesetup.fittopagestall")?,
        })
    }
    pub fn left_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.leftmargin",
            "PageSetup.LeftMargin",
        )
    }
    pub fn set_left_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.leftmargin",
            OwnedVariant::f64(value),
        )
    }
    pub fn right_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.rightmargin",
            "PageSetup.RightMargin",
        )
    }
    pub fn set_right_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.rightmargin",
            OwnedVariant::f64(value),
        )
    }
    pub fn top_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.topmargin",
            "PageSetup.TopMargin",
        )
    }
    pub fn set_top_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.topmargin",
            OwnedVariant::f64(value),
        )
    }
    pub fn bottom_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.bottommargin",
            "PageSetup.BottomMargin",
        )
    }
    pub fn set_bottom_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.bottommargin",
            OwnedVariant::f64(value),
        )
    }
    pub fn header_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.headermargin",
            "PageSetup.HeaderMargin",
        )
    }
    pub fn set_header_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.headermargin",
            OwnedVariant::f64(value),
        )
    }
    pub fn footer_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.footermargin",
            "PageSetup.FooterMargin",
        )
    }
    pub fn set_footer_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.footermargin",
            OwnedVariant::f64(value),
        )
    }
    pub fn print_headings(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.printheadings")
    }
    pub fn set_print_headings(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printheadings",
            OwnedVariant::bool(value),
        )
    }
    pub fn print_gridlines(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.printgridlines")
    }
    pub fn set_print_gridlines(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printgridlines",
            OwnedVariant::bool(value),
        )
    }
    pub fn center_horizontally(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.centerhorizontally")
    }
    pub fn set_center_horizontally(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.centerhorizontally",
            OwnedVariant::bool(value),
        )
    }
    pub fn center_vertically(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.centervertically")
    }
    pub fn set_center_vertically(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.centervertically",
            OwnedVariant::bool(value),
        )
    }
    pub fn black_and_white(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.blackandwhite")
    }
    pub fn set_black_and_white(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.blackandwhite",
            OwnedVariant::bool(value),
        )
    }
    pub fn draft(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.draft")
    }
    pub fn set_draft(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.draft",
            OwnedVariant::bool(value),
        )
    }
    pub fn first_page_number(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.pagesetup.firstpagenumber",
            "PageSetup.FirstPageNumber",
        )
    }
    pub fn set_first_page_number(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.firstpagenumber",
            OwnedVariant::i32(value),
        )
    }
    pub fn print_comments(&self) -> Result<PrintLocation, ExcelComError> {
        Ok(PrintLocation::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.printcomments",
            "PageSetup.PrintComments",
        )?))
    }
    pub fn set_print_comments(&self, value: PrintLocation) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printcomments",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn print_quality(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.pagesetup.printquality",
            "PageSetup.PrintQuality",
        )
    }
    pub fn set_print_quality(&self, value: i32) -> Result<(), ExcelComError> {
        if value <= 0 {
            return Err(ExcelComError::Unsupported {
                detail: "PageSetup.PrintQuality must be positive",
            });
        }
        put(
            &self.inner,
            "excel.pagesetup.printquality",
            OwnedVariant::i32(value),
        )
    }
    pub fn print_errors(&self) -> Result<PrintErrors, ExcelComError> {
        Ok(PrintErrors::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.printerrors",
            "PageSetup.PrintErrors",
        )?))
    }
    pub fn set_print_errors(&self, value: PrintErrors) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printerrors",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn print_area(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.printarea")
    }
    pub fn set_print_area(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printarea",
            text_bstr(&range.external_address(ReferenceStyle::A1)?)?,
        )
    }
    pub fn clear_print_area(&self) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.pagesetup.printarea", text_bstr("")?)
    }
    pub fn print_title_rows(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.printtitlerows")
    }
    pub fn set_print_title_rows(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printtitlerows",
            text_bstr(&range.external_address(ReferenceStyle::A1)?)?,
        )
    }
    pub fn clear_print_title_rows(&self) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printtitlerows",
            text_bstr("")?,
        )
    }
    pub fn print_title_columns(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.printtitlecolumns")
    }
    pub fn set_print_title_columns(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printtitlecolumns",
            text_bstr(&range.external_address(ReferenceStyle::A1)?)?,
        )
    }
    pub fn clear_print_title_columns(&self) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printtitlecolumns",
            text_bstr("")?,
        )
    }
    pub fn left_header(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.leftheader")
    }
    pub fn set_left_header(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.pagesetup.leftheader", text_bstr(value)?)
    }
    pub fn center_header(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.centerheader")
    }
    pub fn set_center_header(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.centerheader",
            text_bstr(value)?,
        )
    }
    pub fn right_header(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.rightheader")
    }
    pub fn set_right_header(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.rightheader",
            text_bstr(value)?,
        )
    }
    pub fn left_footer(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.leftfooter")
    }
    pub fn set_left_footer(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.pagesetup.leftfooter", text_bstr(value)?)
    }
    pub fn center_footer(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.centerfooter")
    }
    pub fn set_center_footer(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.centerfooter",
            text_bstr(value)?,
        )
    }
    pub fn right_footer(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.rightfooter")
    }
    pub fn set_right_footer(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.rightfooter",
            text_bstr(value)?,
        )
    }
}
