//! Presentation operations implemented on existing workbook, worksheet, and range wrappers.
#![allow(missing_docs)]

use super::*;

impl Application {
    pub fn active_workbook(&self) -> Result<Workbook, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.activeworkbook",
            Workbook::from_dispatch,
        )
    }
    pub fn active_sheet(&self) -> Result<Sheet, ExcelComError> {
        get_sheet(self.dispatch_object(), "excel.application.activesheet")
    }
    pub fn active_cell(&self) -> Result<Range, ExcelComError> {
        get_range(self.dispatch_object(), "excel.application.activecell")
    }
    pub fn active_range(&self) -> Result<Range, ExcelComError> {
        get_range(self.dispatch_object(), "excel.application.selection")
    }
    pub fn active_window(&self) -> Result<Window, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.activewindow",
            Window::from_dispatch,
        )
    }
    pub fn windows(&self) -> Result<Windows, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.windows",
            Windows::from_dispatch,
        )
    }
    pub fn sheets(&self) -> Result<Sheets, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.sheets",
            Sheets::from_dispatch,
        )
    }
    pub fn worksheets(&self) -> Result<crate::excel::Worksheets, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.worksheets",
            crate::excel::Worksheets::from_dispatch,
        )
    }
    pub fn go_to(&self, range: &Range, select: Option<bool>) -> Result<(), ExcelComError> {
        let mut arguments = PositionalArguments::new();
        arguments.push_object(range.dispatch_object());
        arguments.push_optional(select.map(OwnedVariant::bool));
        call(
            self.dispatch_object(),
            "excel.application.goto",
            arguments.into_inner(),
        )
    }
    pub fn automation_security(&self) -> Result<AutomationSecurity, ExcelComError> {
        Ok(AutomationSecurity::from_raw(get_i32(
            self.dispatch_object(),
            "excel.application.automationsecurity",
            "Application.AutomationSecurity",
        )?))
    }
    pub fn set_automation_security(&self, value: AutomationSecurity) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.application.automationsecurity",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn automation_security_guard(
        &self,
        value: AutomationSecurity,
    ) -> Result<AutomationSecurityGuard<'_>, ExcelComError> {
        let previous = self.automation_security()?;
        self.set_automation_security(value)?;
        Ok(AutomationSecurityGuard {
            application: self,
            previous,
            active: true,
        })
    }
    pub fn run_macro(
        &self,
        macro_name: &str,
        arguments: &[AutomationValue],
    ) -> Result<AutomationValue, ExcelComError> {
        if arguments.len() > 30 {
            return Err(ExcelComError::Unsupported {
                detail: "Application.Run accepts at most 30 macro arguments",
            });
        }
        let mut values = PositionalArguments::new();
        values.push_result(text_bstr(macro_name))?;
        for argument in arguments {
            values.push_argument(
                AutomationArgument::Value(argument.clone()),
                ConversionPolicy::default(),
            )?;
        }
        let result = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.application.run"), false),
            values.into_inner(),
            false,
        )?;
        if result.vt() == VT_DISPATCH {
            return Err(ExcelComError::Unsupported {
                detail: "Application.Run returned a dispatch object",
            });
        }
        decode_variant(&result, ConversionPolicy::default())
    }
}

impl Workbooks {
    pub fn open_safely<P: AsRef<Path>>(
        &self,
        filename: P,
        options: SafeWorkbookOpenOptions<'_>,
    ) -> Result<Workbook, ExcelComError> {
        let application: Application = get_object(
            self.dispatch_object(),
            "excel.workbooks.application",
            Application::from_dispatch,
        )?;
        let guard = application.automation_security_guard(AutomationSecurity::FORCE_DISABLE)?;
        let opened = self.open(filename, options.open);
        let restored = guard.restore();
        match (opened, restored) {
            (Ok(workbook), Ok(())) => Ok(workbook),
            (Err(error), _) | (_, Err(error)) => Err(error),
        }
    }
}

impl Workbook {
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.workbook.activate", vec![])
    }
    pub fn active_sheet(&self) -> Result<Sheet, ExcelComError> {
        get_sheet(self.dispatch_object(), "excel.workbook.activesheet")
    }
    pub fn sheets(&self) -> Result<Sheets, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.workbook.sheets",
            Sheets::from_dispatch,
        )
    }
    pub fn windows(&self) -> Result<Windows, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.workbook.windows",
            Windows::from_dispatch,
        )
    }
    pub fn has_vb_project(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.workbook.hasvbproject")
    }
    pub fn protect_structure(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.workbook.protectstructure")
    }
    pub fn protect_windows(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.workbook.protectwindows")
    }
    pub fn protect(&self, options: &WorkbookProtectOptions<'_>) -> Result<(), ExcelComError> {
        let mut arguments = PositionalArguments::new();
        push_optional_text(&mut arguments, options.password)?;
        arguments.push_optional(options.structure.map(OwnedVariant::bool));
        arguments.push_optional(options.windows.map(OwnedVariant::bool));
        call(
            self.dispatch_object(),
            "excel.workbook.protect-2029",
            arguments.into_inner(),
        )
    }
    pub fn unprotect(&self, password: Option<&str>) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        push_optional_text(&mut a, password)?;
        call(
            self.dispatch_object(),
            "excel.workbook.unprotect",
            a.into_inner(),
        )
    }
    pub fn print_preview(&self) -> Result<(), ExcelComError> {
        call(
            self.dispatch_object(),
            "excel.workbook.printpreview",
            vec![],
        )
    }
    pub fn print_out(&self, options: &PrintOutOptions<'_>) -> Result<(), ExcelComError> {
        print_out(
            self.dispatch_object(),
            "excel.workbook.printout-2361",
            options,
        )
    }
    pub fn export_as_fixed_format(
        &self,
        format: FixedFormatType,
        options: &FixedFormatOptions<'_>,
    ) -> Result<(), ExcelComError> {
        fixed_format(
            self.dispatch_object(),
            "excel.workbook.exportasfixedformat-3175",
            format,
            options,
        )
    }
}

impl Worksheet {
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.worksheet.activate", vec![])
    }
    pub fn select(&self, replace: Option<bool>) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_optional(replace.map(OwnedVariant::bool));
        call(
            self.dispatch_object(),
            "excel.worksheet.select",
            a.into_inner(),
        )
    }
    pub fn copy_to(&self, destination: SheetDestination<'_>) -> Result<(), ExcelComError> {
        worksheet_copy_move(self, "excel.worksheet.copy", destination)
    }
    pub fn move_to(&self, destination: SheetDestination<'_>) -> Result<(), ExcelComError> {
        worksheet_copy_move(self, "excel.worksheet.move", destination)
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.worksheet.delete", vec![])
    }
    pub fn tab(&self) -> Result<Tab, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.tab",
            Tab::from_dispatch,
        )
    }
    pub fn page_setup(&self) -> Result<PageSetup, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.pagesetup",
            PageSetup::from_dispatch,
        )
    }
    pub fn outline(&self) -> Result<Outline, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.outline",
            Outline::from_dispatch,
        )
    }
    pub fn h_page_breaks(&self) -> Result<HPageBreaks, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.hpagebreaks",
            HPageBreaks::from_dispatch,
        )
    }
    pub fn v_page_breaks(&self) -> Result<VPageBreaks, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.vpagebreaks",
            VPageBreaks::from_dispatch,
        )
    }
    pub fn reset_all_page_breaks(&self) -> Result<(), ExcelComError> {
        call(
            self.dispatch_object(),
            "excel.worksheet.resetallpagebreaks",
            vec![],
        )
    }
    pub fn protect(&self, options: &WorksheetProtectOptions<'_>) -> Result<(), ExcelComError> {
        worksheet_protect(self.dispatch_object(), options)
    }
    pub fn unprotect(&self, password: Option<&str>) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        push_optional_text(&mut a, password)?;
        call(
            self.dispatch_object(),
            "excel.worksheet.unprotect",
            a.into_inner(),
        )
    }
    pub fn protect_contents(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.worksheet.protectcontents")
    }
    pub fn protect_drawing_objects(&self) -> Result<bool, ExcelComError> {
        get_bool(
            self.dispatch_object(),
            "excel.worksheet.protectdrawingobjects",
        )
    }
    pub fn protect_scenarios(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.worksheet.protectscenarios")
    }
    pub fn protection_mode(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.worksheet.protectionmode")
    }
    pub fn print_preview(&self) -> Result<(), ExcelComError> {
        call(
            self.dispatch_object(),
            "excel.worksheet.printpreview",
            vec![],
        )
    }
    pub fn print_out(&self, options: &PrintOutOptions<'_>) -> Result<(), ExcelComError> {
        print_out(
            self.dispatch_object(),
            "excel.worksheet.printout-2361",
            options,
        )
    }
    pub fn export_as_fixed_format(
        &self,
        format: FixedFormatType,
        options: &FixedFormatOptions<'_>,
    ) -> Result<(), ExcelComError> {
        fixed_format(
            self.dispatch_object(),
            "excel.worksheet.exportasfixedformat-3175",
            format,
            options,
        )
    }
}

impl Range {
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.range.activate", vec![])
    }
    pub fn select(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.range.select", vec![])
    }
    pub fn merge(&self, across: Option<bool>) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_optional(across.map(OwnedVariant::bool));
        call(self.dispatch_object(), "excel.range.merge", a.into_inner())
    }
    pub fn unmerge(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.range.unmerge", vec![])
    }
    pub fn merge_cells(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.mergecells", mixed_bool)
    }
    pub fn merge_area(&self) -> Result<Range, ExcelComError> {
        get_range(self.dispatch_object(), "excel.range.mergearea")
    }
    pub fn orientation(&self) -> Result<MixedValue<i32>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.orientation", mixed_i32)
    }
    pub fn set_orientation(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.orientation",
            OwnedVariant::i32(value),
        )
    }
    pub fn indent_level(&self) -> Result<MixedValue<i32>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.indentlevel", mixed_i32)
    }
    pub fn set_indent_level(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.indentlevel",
            OwnedVariant::i32(value),
        )
    }
    pub fn shrink_to_fit(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(
            self.dispatch_object(),
            "excel.range.shrinktofit",
            mixed_bool,
        )
    }
    pub fn set_shrink_to_fit(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.shrinktofit",
            OwnedVariant::bool(value),
        )
    }
    pub fn reading_order(&self) -> Result<MixedValue<ReadingOrder>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.readingorder", |v| {
            mixed_i32(v).map(|m| map_mixed(m, ReadingOrder::from_raw))
        })
    }
    pub fn set_reading_order(&self, value: ReadingOrder) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.readingorder",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn locked(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.locked", mixed_bool)
    }
    pub fn set_locked(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.locked",
            OwnedVariant::bool(value),
        )
    }
    pub fn formula_hidden(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(
            self.dispatch_object(),
            "excel.range.formulahidden",
            mixed_bool,
        )
    }
    pub fn set_formula_hidden(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.formulahidden",
            OwnedVariant::bool(value),
        )
    }
    pub fn group(&self) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        for _ in 0..4 {
            a.push_optional(None);
        }
        call(self.dispatch_object(), "excel.range.group", a.into_inner())
    }
    pub fn ungroup(&self) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        for _ in 0..4 {
            a.push_optional(None);
        }
        call(
            self.dispatch_object(),
            "excel.range.ungroup",
            a.into_inner(),
        )
    }
    pub fn show_detail(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.range.showdetail")
    }
    pub fn set_show_detail(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.showdetail",
            OwnedVariant::bool(value),
        )
    }
    pub fn print_preview(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.range.printpreview", vec![])
    }
    pub fn print_out(&self, options: &PrintOutOptions<'_>) -> Result<(), ExcelComError> {
        print_out(self.dispatch_object(), "excel.range.printout-2361", options)
    }
    pub fn export_as_fixed_format(
        &self,
        format: FixedFormatType,
        options: &FixedFormatOptions<'_>,
    ) -> Result<(), ExcelComError> {
        fixed_format(
            self.dispatch_object(),
            "excel.range.exportasfixedformat-3175",
            format,
            options,
        )
    }
}
