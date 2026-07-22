use super::MemberId;
use crate::automation::{MemberDescriptor, MemberKind};

const APPLICATION_VERSION: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.version"),
    name: "Version",
    kind: MemberKind::PropertyGet,
};
const APPLICATION_VISIBLE_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.visible"),
    name: "Visible",
    kind: MemberKind::PropertyGet,
};
const APPLICATION_VISIBLE_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.visible"),
    name: "Visible",
    kind: MemberKind::PropertyPut,
};
const APPLICATION_DISPLAY_ALERTS_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.displayalerts"),
    name: "DisplayAlerts",
    kind: MemberKind::PropertyGet,
};
const APPLICATION_DISPLAY_ALERTS_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.displayalerts"),
    name: "DisplayAlerts",
    kind: MemberKind::PropertyPut,
};
const APPLICATION_WORKBOOKS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.workbooks"),
    name: "Workbooks",
    kind: MemberKind::PropertyGet,
};
const APPLICATION_QUIT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.quit"),
    name: "Quit",
    kind: MemberKind::Method,
};
const APPLICATION_UNION: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.union"),
    name: "Union",
    kind: MemberKind::Method,
};
const APPLICATION_REFERENCE_STYLE_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.referencestyle"),
    name: "ReferenceStyle",
    kind: MemberKind::PropertyGet,
};
const APPLICATION_REFERENCE_STYLE_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.referencestyle"),
    name: "ReferenceStyle",
    kind: MemberKind::PropertyPut,
};
const APPLICATION_CONVERT_FORMULA: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.convertformula"),
    name: "ConvertFormula",
    kind: MemberKind::Method,
};
const APPLICATION_EVALUATE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.application.evaluate-1"),
    name: "Evaluate",
    kind: MemberKind::Method,
};
const WORKBOOKS_COUNT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbooks.count"),
    name: "Count",
    kind: MemberKind::PropertyGet,
};
const WORKBOOKS_ITEM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbooks.item"),
    name: "Item",
    kind: MemberKind::PropertyGet,
};
const WORKBOOKS_NEW_ENUM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbooks.newenum"),
    name: "_NewEnum",
    kind: MemberKind::PropertyGet,
};
// Microsoft's C++ reference and runtime evidence classify Add as PROPERTYGET.
const WORKBOOKS_ADD: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbooks.add"),
    name: "Add",
    kind: MemberKind::PropertyGet,
};
const WORKBOOKS_OPEN: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbooks.open-1923"),
    name: "Open",
    kind: MemberKind::Method,
};
const WORKBOOK_NAME: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.name"),
    name: "Name",
    kind: MemberKind::PropertyGet,
};
const WORKBOOK_SAVED_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.saved"),
    name: "Saved",
    kind: MemberKind::PropertyGet,
};
const WORKBOOK_SAVED_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.saved"),
    name: "Saved",
    kind: MemberKind::PropertyPut,
};
const WORKBOOK_CLOSE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.close"),
    name: "Close",
    kind: MemberKind::Method,
};
const WORKBOOK_FULL_NAME: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.fullname"),
    name: "FullName",
    kind: MemberKind::PropertyGet,
};
const WORKBOOK_PATH: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.path"),
    name: "Path",
    kind: MemberKind::PropertyGet,
};
const WORKBOOK_FILE_FORMAT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.fileformat"),
    name: "FileFormat",
    kind: MemberKind::PropertyGet,
};
const WORKBOOK_READ_ONLY: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.readonly"),
    name: "ReadOnly",
    kind: MemberKind::PropertyGet,
};
const WORKBOOK_SAVE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.save"),
    name: "Save",
    kind: MemberKind::Method,
};
const WORKBOOK_SAVE_AS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.saveas-3174"),
    name: "SaveAs",
    kind: MemberKind::Method,
};
const WORKBOOK_SAVE_COPY_AS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.savecopyas"),
    name: "SaveCopyAs",
    kind: MemberKind::Method,
};
const WORKBOOK_WORKSHEETS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.worksheets"),
    name: "Worksheets",
    kind: MemberKind::PropertyGet,
};
const WORKBOOK_NAMES: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.names"),
    name: "Names",
    kind: MemberKind::PropertyGet,
};
const WORKSHEETS_COUNT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheets.count"),
    name: "Count",
    kind: MemberKind::PropertyGet,
};
const WORKSHEETS_ITEM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheets.item"),
    name: "Item",
    kind: MemberKind::PropertyGet,
};
const WORKSHEETS_NEW_ENUM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheets.newenum"),
    name: "_NewEnum",
    kind: MemberKind::PropertyGet,
};
const WORKSHEETS_ADD: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheets.add"),
    name: "Add",
    kind: MemberKind::Method,
};
const WORKSHEET_NAME_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.name"),
    name: "Name",
    kind: MemberKind::PropertyGet,
};
const WORKSHEET_NAME_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.name"),
    name: "Name",
    kind: MemberKind::PropertyPut,
};
const WORKSHEET_INDEX: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.index"),
    name: "Index",
    kind: MemberKind::PropertyGet,
};
const WORKSHEET_VISIBLE_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.visible"),
    name: "Visible",
    kind: MemberKind::PropertyGet,
};
const WORKSHEET_VISIBLE_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.visible"),
    name: "Visible",
    kind: MemberKind::PropertyPut,
};
const WORKSHEET_RANGE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.range"),
    name: "Range",
    kind: MemberKind::PropertyGet,
};
const WORKSHEET_USED_RANGE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.usedrange"),
    name: "UsedRange",
    kind: MemberKind::PropertyGet,
};
const WORKSHEET_APPLICATION: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.application"),
    name: "Application",
    kind: MemberKind::PropertyGet,
};
const WORKSHEET_CELLS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.cells"),
    name: "Cells",
    kind: MemberKind::PropertyGet,
};
const WORKSHEET_NAMES: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.names"),
    name: "Names",
    kind: MemberKind::PropertyGet,
};
const WORKSHEET_EVALUATE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.worksheet.evaluate-1"),
    name: "Evaluate",
    kind: MemberKind::Method,
};
const RANGE_ADDRESS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.address"),
    name: "Address",
    kind: MemberKind::PropertyGet,
};
const RANGE_ROW: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.row"),
    name: "Row",
    kind: MemberKind::PropertyGet,
};
const RANGE_COLUMN: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.column"),
    name: "Column",
    kind: MemberKind::PropertyGet,
};
const RANGE_COUNT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.count"),
    name: "Count",
    kind: MemberKind::PropertyGet,
};
const RANGE_ROWS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.rows"),
    name: "Rows",
    kind: MemberKind::PropertyGet,
};
const RANGE_COLUMNS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.columns"),
    name: "Columns",
    kind: MemberKind::PropertyGet,
};
const RANGE_NEW_ENUM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.newenum"),
    name: "_NewEnum",
    kind: MemberKind::PropertyGet,
};
const RANGE_CELLS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.cells"),
    name: "Cells",
    kind: MemberKind::PropertyGet,
};
const RANGE_ITEM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.item"),
    name: "Item",
    kind: MemberKind::PropertyGet,
};
const RANGE_OFFSET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.offset"),
    name: "Offset",
    kind: MemberKind::PropertyGet,
};
const RANGE_RESIZE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.resize"),
    name: "Resize",
    kind: MemberKind::PropertyGet,
};
const RANGE_AREAS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.areas"),
    name: "Areas",
    kind: MemberKind::PropertyGet,
};
const RANGE_ENTIRE_ROW: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.entirerow"),
    name: "EntireRow",
    kind: MemberKind::PropertyGet,
};
const RANGE_ENTIRE_COLUMN: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.entirecolumn"),
    name: "EntireColumn",
    kind: MemberKind::PropertyGet,
};
const AREAS_COUNT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.areas.count"),
    name: "Count",
    kind: MemberKind::PropertyGet,
};
const AREAS_ITEM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.areas.item"),
    name: "Item",
    kind: MemberKind::PropertyGet,
};
const AREAS_NEW_ENUM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.areas.newenum"),
    name: "_NewEnum",
    kind: MemberKind::PropertyGet,
};
const NAMES_COUNT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.names.count"),
    name: "Count",
    kind: MemberKind::PropertyGet,
};
const NAMES_ITEM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.names.item"),
    name: "Item",
    kind: MemberKind::Method,
};
const NAMES_NEW_ENUM: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.names.newenum"),
    name: "_NewEnum",
    kind: MemberKind::PropertyGet,
};
const NAMES_ADD: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.names.add"),
    name: "Add",
    kind: MemberKind::Method,
};
const NAME_NAME: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.name.name"),
    name: "Name",
    kind: MemberKind::PropertyGet,
};
const NAME_REFERS_TO: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.name.refersto"),
    name: "RefersTo",
    kind: MemberKind::PropertyGet,
};
const NAME_REFERS_TO_R1C1: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.name.referstor1c1"),
    name: "RefersToR1C1",
    kind: MemberKind::PropertyGet,
};
const NAME_REFERS_TO_RANGE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.name.referstorange"),
    name: "RefersToRange",
    kind: MemberKind::PropertyGet,
};
const NAME_VISIBLE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.name.visible"),
    name: "Visible",
    kind: MemberKind::PropertyGet,
};
const NAME_DELETE: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.name.delete"),
    name: "Delete",
    kind: MemberKind::Method,
};
const RANGE_VALUE_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.value"),
    name: "Value",
    kind: MemberKind::PropertyGet,
};
const RANGE_VALUE_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.value"),
    name: "Value",
    kind: MemberKind::PropertyPut,
};
const RANGE_VALUE2_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.value2"),
    name: "Value2",
    kind: MemberKind::PropertyGet,
};
const RANGE_VALUE2_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.value2"),
    name: "Value2",
    kind: MemberKind::PropertyPut,
};
const RANGE_FORMULA_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.formula"),
    name: "Formula",
    kind: MemberKind::PropertyGet,
};
const RANGE_FORMULA_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.formula"),
    name: "Formula",
    kind: MemberKind::PropertyPut,
};
const RANGE_FORMULA2_GET: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.formula2"),
    name: "Formula2",
    kind: MemberKind::PropertyGet,
};
const RANGE_FORMULA2_PUT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.formula2"),
    name: "Formula2",
    kind: MemberKind::PropertyPut,
};
const RANGE_CLEAR_CONTENTS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.range.clearcontents"),
    name: "_ClearContents",
    kind: MemberKind::Method,
};

/// Inventory IDs implemented by the experimental wrapper slice.
pub const IMPLEMENTED_MEMBER_IDS: &[&str] = &[
    "excel.application.version",
    "excel.application.visible",
    "excel.application.displayalerts",
    "excel.application.workbooks",
    "excel.application.quit",
    "excel.application.union",
    "excel.application.referencestyle",
    "excel.application.calculation",
    "excel.application.calculationstate",
    "excel.application.calculationversion",
    "excel.application.calculatebeforesave",
    "excel.application.calculate",
    "excel.application.calculatefull",
    "excel.application.calculatefullrebuild",
    "excel.application.convertformula",
    "excel.application.evaluate-1",
    "excel.workbooks.count",
    "excel.workbooks.item",
    "excel.workbooks.newenum",
    "excel.workbooks.add",
    "excel.workbooks.open-1923",
    "excel.workbook.name",
    "excel.workbook.saved",
    "excel.workbook.close",
    "excel.workbook.fullname",
    "excel.workbook.path",
    "excel.workbook.fileformat",
    "excel.workbook.readonly",
    "excel.workbook.save",
    "excel.workbook.saveas-3174",
    "excel.workbook.savecopyas",
    "excel.workbook.worksheets",
    "excel.workbook.names",
    "excel.worksheets.count",
    "excel.worksheets.item",
    "excel.worksheets.newenum",
    "excel.worksheets.add",
    "excel.worksheet.name",
    "excel.worksheet.index",
    "excel.worksheet.visible",
    "excel.worksheet.range",
    "excel.worksheet.usedrange",
    "excel.worksheet.application",
    "excel.worksheet.cells",
    "excel.worksheet.names",
    "excel.worksheet.evaluate-1",
    "excel.worksheet.calculate",
    "excel.range.address",
    "excel.range.row",
    "excel.range.column",
    "excel.range.count",
    "excel.range.rows",
    "excel.range.columns",
    "excel.range.newenum",
    "excel.range.cells",
    "excel.range.item",
    "excel.range.offset",
    "excel.range.resize",
    "excel.range.areas",
    "excel.range.entirerow",
    "excel.range.entirecolumn",
    "excel.range.value",
    "excel.range.value2",
    "excel.range.formula",
    "excel.range.formula2",
    "excel.range.formular1c1",
    "excel.range.formula2r1c1",
    "excel.range.formulalocal",
    "excel.range.formular1c1local",
    "excel.range.formulaarray",
    "excel.range.hasformula",
    "excel.range.hasarray",
    "excel.range.currentarray",
    "excel.range.hasspill",
    "excel.range.spillingtorange",
    "excel.range.spillparent",
    "excel.range.calculate",
    "excel.range.dirty",
    "excel.range.directprecedents",
    "excel.range.directdependents",
    "excel.range.precedents",
    "excel.range.dependents",
    "excel.range.specialcells",
    "excel.range.find",
    "excel.range.findnext",
    "excel.range.findprevious",
    "excel.range.replace-3305",
    "excel.range.clearcontents",
    "excel.range.font",
    "excel.range.interior",
    "excel.range.borders",
    "excel.range.numberformat",
    "excel.range.horizontalalignment",
    "excel.range.verticalalignment",
    "excel.range.wraptext",
    "excel.range.rowheight",
    "excel.range.columnwidth",
    "excel.range.autofit",
    "excel.font.name",
    "excel.font.size",
    "excel.font.bold",
    "excel.font.italic",
    "excel.font.underline",
    "excel.font.strikethrough",
    "excel.font.color",
    "excel.font.colorindex",
    "excel.interior.color",
    "excel.interior.colorindex",
    "excel.interior.pattern",
    "excel.interior.patterncolor",
    "excel.interior.patterncolorindex",
    "excel.borders.count",
    "excel.borders.item",
    "excel.borders.newenum",
    "excel.borders.color",
    "excel.borders.colorindex",
    "excel.borders.linestyle",
    "excel.borders.weight",
    "excel.border.color",
    "excel.border.colorindex",
    "excel.border.linestyle",
    "excel.border.weight",
    "excel.areas.count",
    "excel.areas.item",
    "excel.areas.newenum",
    "excel.names.count",
    "excel.names.item",
    "excel.names.newenum",
    "excel.names.add",
    "excel.name.name",
    "excel.name.refersto",
    "excel.name.referstor1c1",
    "excel.name.referstorange",
    "excel.name.visible",
    "excel.name.delete",
    "excel.worksheet.listobjects",
    "excel.worksheet.autofiltermode",
    "excel.worksheet.filtermode",
    "excel.worksheet.showalldata",
    "excel.worksheet.autofilter-3289",
    "excel.range.worksheet",
    "excel.listobjects.count",
    "excel.listobjects.item",
    "excel.listobjects.newenum",
    "excel.listobjects.add",
    "excel.listobject.name",
    "excel.listobject.displayname",
    "excel.listobject.range",
    "excel.listobject.headerrowrange",
    "excel.listobject.databodyrange",
    "excel.listobject.totalsrowrange",
    "excel.listobject.insertrowrange",
    "excel.listobject.showheaders",
    "excel.listobject.showtotals",
    "excel.listobject.showautofilter",
    "excel.listobject.tablestyle",
    "excel.listobject.listcolumns",
    "excel.listobject.listrows",
    "excel.listobject.autofilter-3289",
    "excel.listobject.sort-3288",
    "excel.listobject.resize",
    "excel.listobject.delete",
    "excel.listobject.unlist",
    "excel.listcolumns.count",
    "excel.listcolumns.item",
    "excel.listcolumns.newenum",
    "excel.listcolumns.add",
    "excel.listcolumn.index",
    "excel.listcolumn.name",
    "excel.listcolumn.range",
    "excel.listcolumn.databodyrange",
    "excel.listcolumn.total",
    "excel.listcolumn.totalscalculation",
    "excel.listcolumn.delete",
    "excel.listrows.count",
    "excel.listrows.item",
    "excel.listrows.newenum",
    "excel.listrows.add",
    "excel.listrow.index",
    "excel.listrow.range",
    "excel.listrow.delete",
    "excel.autofilter.range",
    "excel.autofilter.filters",
    "excel.filters.count",
    "excel.filters.item",
    "excel.filters.newenum",
    "excel.filter.on",
    "excel.filter.operator-2641",
    "excel.filter.criteria1",
    "excel.filter.criteria2",
    "excel.range.autofilter-3289",
    "excel.range.sort",
    "excel.sort.sortfields",
    "excel.sort.setrange",
    "excel.sort.header",
    "excel.sort.matchcase",
    "excel.sort.orientation",
    "excel.sort.apply",
    "excel.sortfields.count",
    "excel.sortfields.clear",
    "excel.sortfields.add",
    "excel.range.validation",
    "excel.validation.add",
    "excel.validation.delete",
    "excel.validation.value",
    "excel.validation.type",
    "excel.validation.alertstyle",
    "excel.validation.operator",
    "excel.validation.formula1",
    "excel.validation.formula2",
    "excel.validation.ignoreblank",
    "excel.validation.incelldropdown",
    "excel.validation.showinput",
    "excel.validation.showerror",
    "excel.validation.inputtitle",
    "excel.validation.inputmessage",
    "excel.validation.errortitle",
    "excel.validation.errormessage",
    "excel.range.removeduplicates",
    "excel.range.currentregion",
    "excel.range.hidden",
    "excel.range.insert",
    "excel.range.delete",
    "excel.range.clear",
    "excel.range.clearformats",
    "excel.range.clearcomments",
    "excel.range.clearhyperlinks",
    "excel.range.copy",
    "excel.range.cut",
    "excel.range.pastespecial-1928",
];

pub(crate) fn member(id: MemberId, put: bool) -> MemberDescriptor {
    if let Some(descriptor) = structured_data_member(id, put) {
        return descriptor;
    }
    match (id.as_str(), put) {
        ("excel.application.version", _) => APPLICATION_VERSION,
        ("excel.application.visible", false) => APPLICATION_VISIBLE_GET,
        ("excel.application.visible", true) => APPLICATION_VISIBLE_PUT,
        ("excel.application.displayalerts", false) => APPLICATION_DISPLAY_ALERTS_GET,
        ("excel.application.displayalerts", true) => APPLICATION_DISPLAY_ALERTS_PUT,
        ("excel.application.workbooks", _) => APPLICATION_WORKBOOKS,
        ("excel.application.quit", _) => APPLICATION_QUIT,
        ("excel.application.union", _) => APPLICATION_UNION,
        ("excel.application.referencestyle", false) => APPLICATION_REFERENCE_STYLE_GET,
        ("excel.application.referencestyle", true) => APPLICATION_REFERENCE_STYLE_PUT,
        ("excel.application.calculation", false) => MemberDescriptor {
            id: MemberId::new("excel.application.calculation"),
            name: "Calculation",
            kind: MemberKind::PropertyGet,
        },
        ("excel.application.calculation", true) => MemberDescriptor {
            id: MemberId::new("excel.application.calculation"),
            name: "Calculation",
            kind: MemberKind::PropertyPut,
        },
        ("excel.application.calculationstate", _) => MemberDescriptor {
            id: MemberId::new("excel.application.calculationstate"),
            name: "CalculationState",
            kind: MemberKind::PropertyGet,
        },
        ("excel.application.calculationversion", _) => MemberDescriptor {
            id: MemberId::new("excel.application.calculationversion"),
            name: "CalculationVersion",
            kind: MemberKind::PropertyGet,
        },
        ("excel.application.calculatebeforesave", false) => MemberDescriptor {
            id: MemberId::new("excel.application.calculatebeforesave"),
            name: "CalculateBeforeSave",
            kind: MemberKind::PropertyGet,
        },
        ("excel.application.calculatebeforesave", true) => MemberDescriptor {
            id: MemberId::new("excel.application.calculatebeforesave"),
            name: "CalculateBeforeSave",
            kind: MemberKind::PropertyPut,
        },
        ("excel.application.calculate", _) => MemberDescriptor {
            id: MemberId::new("excel.application.calculate"),
            name: "Calculate",
            kind: MemberKind::Method,
        },
        ("excel.application.calculatefull", _) => MemberDescriptor {
            id: MemberId::new("excel.application.calculatefull"),
            name: "CalculateFull",
            kind: MemberKind::Method,
        },
        ("excel.application.calculatefullrebuild", _) => MemberDescriptor {
            id: MemberId::new("excel.application.calculatefullrebuild"),
            name: "CalculateFullRebuild",
            kind: MemberKind::Method,
        },
        ("excel.application.convertformula", _) => APPLICATION_CONVERT_FORMULA,
        ("excel.application.evaluate-1", _) => APPLICATION_EVALUATE,
        ("excel.workbooks.count", _) => WORKBOOKS_COUNT,
        ("excel.workbooks.item", _) => WORKBOOKS_ITEM,
        ("excel.workbooks.newenum", _) => WORKBOOKS_NEW_ENUM,
        ("excel.workbooks.add", _) => WORKBOOKS_ADD,
        ("excel.workbooks.open-1923", _) => WORKBOOKS_OPEN,
        ("excel.workbook.name", _) => WORKBOOK_NAME,
        ("excel.workbook.saved", false) => WORKBOOK_SAVED_GET,
        ("excel.workbook.saved", true) => WORKBOOK_SAVED_PUT,
        ("excel.workbook.close", _) => WORKBOOK_CLOSE,
        ("excel.workbook.fullname", _) => WORKBOOK_FULL_NAME,
        ("excel.workbook.path", _) => WORKBOOK_PATH,
        ("excel.workbook.fileformat", _) => WORKBOOK_FILE_FORMAT,
        ("excel.workbook.readonly", _) => WORKBOOK_READ_ONLY,
        ("excel.workbook.save", _) => WORKBOOK_SAVE,
        ("excel.workbook.saveas-3174", _) => WORKBOOK_SAVE_AS,
        ("excel.workbook.savecopyas", _) => WORKBOOK_SAVE_COPY_AS,
        ("excel.workbook.worksheets", _) => WORKBOOK_WORKSHEETS,
        ("excel.workbook.names", _) => WORKBOOK_NAMES,
        ("excel.worksheets.count", _) => WORKSHEETS_COUNT,
        ("excel.worksheets.item", _) => WORKSHEETS_ITEM,
        ("excel.worksheets.newenum", _) => WORKSHEETS_NEW_ENUM,
        ("excel.worksheets.add", _) => WORKSHEETS_ADD,
        ("excel.worksheet.name", false) => WORKSHEET_NAME_GET,
        ("excel.worksheet.name", true) => WORKSHEET_NAME_PUT,
        ("excel.worksheet.index", _) => WORKSHEET_INDEX,
        ("excel.worksheet.visible", false) => WORKSHEET_VISIBLE_GET,
        ("excel.worksheet.visible", true) => WORKSHEET_VISIBLE_PUT,
        ("excel.worksheet.range", _) => WORKSHEET_RANGE,
        ("excel.worksheet.usedrange", _) => WORKSHEET_USED_RANGE,
        ("excel.worksheet.application", _) => WORKSHEET_APPLICATION,
        ("excel.worksheet.cells", _) => WORKSHEET_CELLS,
        ("excel.worksheet.names", _) => WORKSHEET_NAMES,
        ("excel.worksheet.evaluate-1", _) => WORKSHEET_EVALUATE,
        ("excel.worksheet.calculate", _) => MemberDescriptor {
            id: MemberId::new("excel.worksheet.calculate"),
            name: "Calculate",
            kind: MemberKind::Method,
        },
        ("excel.range.address", _) => RANGE_ADDRESS,
        ("excel.range.row", _) => RANGE_ROW,
        ("excel.range.column", _) => RANGE_COLUMN,
        ("excel.range.count", _) => RANGE_COUNT,
        ("excel.range.rows", _) => RANGE_ROWS,
        ("excel.range.columns", _) => RANGE_COLUMNS,
        ("excel.range.newenum", _) => RANGE_NEW_ENUM,
        ("excel.range.cells", _) => RANGE_CELLS,
        ("excel.range.item", _) => RANGE_ITEM,
        ("excel.range.offset", _) => RANGE_OFFSET,
        ("excel.range.resize", _) => RANGE_RESIZE,
        ("excel.range.areas", _) => RANGE_AREAS,
        ("excel.range.entirerow", _) => RANGE_ENTIRE_ROW,
        ("excel.range.entirecolumn", _) => RANGE_ENTIRE_COLUMN,
        ("excel.range.value", false) => RANGE_VALUE_GET,
        ("excel.range.value", true) => RANGE_VALUE_PUT,
        ("excel.range.value2", false) => RANGE_VALUE2_GET,
        ("excel.range.value2", true) => RANGE_VALUE2_PUT,
        ("excel.range.formula", false) => RANGE_FORMULA_GET,
        ("excel.range.formula", true) => RANGE_FORMULA_PUT,
        ("excel.range.formula2", false) => RANGE_FORMULA2_GET,
        ("excel.range.formula2", true) => RANGE_FORMULA2_PUT,
        ("excel.range.formular1c1", false) => MemberDescriptor {
            id: MemberId::new("excel.range.formular1c1"),
            name: "FormulaR1C1",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.formular1c1", true) => MemberDescriptor {
            id: MemberId::new("excel.range.formular1c1"),
            name: "FormulaR1C1",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.formula2r1c1", false) => MemberDescriptor {
            id: MemberId::new("excel.range.formula2r1c1"),
            name: "Formula2R1C1",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.formula2r1c1", true) => MemberDescriptor {
            id: MemberId::new("excel.range.formula2r1c1"),
            name: "Formula2R1C1",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.formulalocal", false) => MemberDescriptor {
            id: MemberId::new("excel.range.formulalocal"),
            name: "FormulaLocal",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.formulalocal", true) => MemberDescriptor {
            id: MemberId::new("excel.range.formulalocal"),
            name: "FormulaLocal",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.formular1c1local", false) => MemberDescriptor {
            id: MemberId::new("excel.range.formular1c1local"),
            name: "FormulaR1C1Local",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.formular1c1local", true) => MemberDescriptor {
            id: MemberId::new("excel.range.formular1c1local"),
            name: "FormulaR1C1Local",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.formulaarray", false) => MemberDescriptor {
            id: MemberId::new("excel.range.formulaarray"),
            name: "FormulaArray",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.formulaarray", true) => MemberDescriptor {
            id: MemberId::new("excel.range.formulaarray"),
            name: "FormulaArray",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.hasformula", _) => MemberDescriptor {
            id: MemberId::new("excel.range.hasformula"),
            name: "HasFormula",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.hasarray", _) => MemberDescriptor {
            id: MemberId::new("excel.range.hasarray"),
            name: "HasArray",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.currentarray", _) => MemberDescriptor {
            id: MemberId::new("excel.range.currentarray"),
            name: "CurrentArray",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.hasspill", _) => MemberDescriptor {
            id: MemberId::new("excel.range.hasspill"),
            name: "HasSpill",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.spillingtorange", _) => MemberDescriptor {
            id: MemberId::new("excel.range.spillingtorange"),
            name: "SpillingToRange",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.spillparent", _) => MemberDescriptor {
            id: MemberId::new("excel.range.spillparent"),
            name: "SpillParent",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.calculate", _) => MemberDescriptor {
            id: MemberId::new("excel.range.calculate"),
            name: "Calculate",
            kind: MemberKind::Method,
        },
        ("excel.range.dirty", _) => MemberDescriptor {
            id: MemberId::new("excel.range.dirty"),
            name: "Dirty",
            kind: MemberKind::Method,
        },
        ("excel.range.directprecedents", _) => MemberDescriptor {
            id: MemberId::new("excel.range.directprecedents"),
            name: "DirectPrecedents",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.directdependents", _) => MemberDescriptor {
            id: MemberId::new("excel.range.directdependents"),
            name: "DirectDependents",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.precedents", _) => MemberDescriptor {
            id: MemberId::new("excel.range.precedents"),
            name: "Precedents",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.dependents", _) => MemberDescriptor {
            id: MemberId::new("excel.range.dependents"),
            name: "Dependents",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.specialcells", _) => MemberDescriptor {
            id: MemberId::new("excel.range.specialcells"),
            name: "SpecialCells",
            kind: MemberKind::Method,
        },
        ("excel.range.find", _) => MemberDescriptor {
            id: MemberId::new("excel.range.find"),
            name: "Find",
            kind: MemberKind::Method,
        },
        ("excel.range.findnext", _) => MemberDescriptor {
            id: MemberId::new("excel.range.findnext"),
            name: "FindNext",
            kind: MemberKind::Method,
        },
        ("excel.range.findprevious", _) => MemberDescriptor {
            id: MemberId::new("excel.range.findprevious"),
            name: "FindPrevious",
            kind: MemberKind::Method,
        },
        ("excel.range.replace-3305", _) => MemberDescriptor {
            id: MemberId::new("excel.range.replace-3305"),
            name: "Replace",
            kind: MemberKind::Method,
        },
        ("excel.range.clearcontents", _) => RANGE_CLEAR_CONTENTS,
        ("excel.range.font", _) => MemberDescriptor {
            id: MemberId::new("excel.range.font"),
            name: "Font",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.interior", _) => MemberDescriptor {
            id: MemberId::new("excel.range.interior"),
            name: "Interior",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.borders", _) => MemberDescriptor {
            id: MemberId::new("excel.range.borders"),
            name: "Borders",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.numberformat", false) => MemberDescriptor {
            id: MemberId::new("excel.range.numberformat"),
            name: "NumberFormat",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.numberformat", true) => MemberDescriptor {
            id: MemberId::new("excel.range.numberformat"),
            name: "NumberFormat",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.horizontalalignment", false) => MemberDescriptor {
            id: MemberId::new("excel.range.horizontalalignment"),
            name: "HorizontalAlignment",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.horizontalalignment", true) => MemberDescriptor {
            id: MemberId::new("excel.range.horizontalalignment"),
            name: "HorizontalAlignment",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.verticalalignment", false) => MemberDescriptor {
            id: MemberId::new("excel.range.verticalalignment"),
            name: "VerticalAlignment",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.verticalalignment", true) => MemberDescriptor {
            id: MemberId::new("excel.range.verticalalignment"),
            name: "VerticalAlignment",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.wraptext", false) => MemberDescriptor {
            id: MemberId::new("excel.range.wraptext"),
            name: "WrapText",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.wraptext", true) => MemberDescriptor {
            id: MemberId::new("excel.range.wraptext"),
            name: "WrapText",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.rowheight", false) => MemberDescriptor {
            id: MemberId::new("excel.range.rowheight"),
            name: "RowHeight",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.rowheight", true) => MemberDescriptor {
            id: MemberId::new("excel.range.rowheight"),
            name: "RowHeight",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.columnwidth", false) => MemberDescriptor {
            id: MemberId::new("excel.range.columnwidth"),
            name: "ColumnWidth",
            kind: MemberKind::PropertyGet,
        },
        ("excel.range.columnwidth", true) => MemberDescriptor {
            id: MemberId::new("excel.range.columnwidth"),
            name: "ColumnWidth",
            kind: MemberKind::PropertyPut,
        },
        ("excel.range.autofit", _) => MemberDescriptor {
            id: MemberId::new("excel.range.autofit"),
            name: "AutoFit",
            kind: MemberKind::Method,
        },
        ("excel.font.name", false) => MemberDescriptor {
            id: MemberId::new("excel.font.name"),
            name: "Name",
            kind: MemberKind::PropertyGet,
        },
        ("excel.font.name", true) => MemberDescriptor {
            id: MemberId::new("excel.font.name"),
            name: "Name",
            kind: MemberKind::PropertyPut,
        },
        ("excel.font.size", false) => MemberDescriptor {
            id: MemberId::new("excel.font.size"),
            name: "Size",
            kind: MemberKind::PropertyGet,
        },
        ("excel.font.size", true) => MemberDescriptor {
            id: MemberId::new("excel.font.size"),
            name: "Size",
            kind: MemberKind::PropertyPut,
        },
        ("excel.font.bold", false) => MemberDescriptor {
            id: MemberId::new("excel.font.bold"),
            name: "Bold",
            kind: MemberKind::PropertyGet,
        },
        ("excel.font.bold", true) => MemberDescriptor {
            id: MemberId::new("excel.font.bold"),
            name: "Bold",
            kind: MemberKind::PropertyPut,
        },
        ("excel.font.italic", false) => MemberDescriptor {
            id: MemberId::new("excel.font.italic"),
            name: "Italic",
            kind: MemberKind::PropertyGet,
        },
        ("excel.font.italic", true) => MemberDescriptor {
            id: MemberId::new("excel.font.italic"),
            name: "Italic",
            kind: MemberKind::PropertyPut,
        },
        ("excel.font.underline", false) => MemberDescriptor {
            id: MemberId::new("excel.font.underline"),
            name: "Underline",
            kind: MemberKind::PropertyGet,
        },
        ("excel.font.underline", true) => MemberDescriptor {
            id: MemberId::new("excel.font.underline"),
            name: "Underline",
            kind: MemberKind::PropertyPut,
        },
        ("excel.font.strikethrough", false) => MemberDescriptor {
            id: MemberId::new("excel.font.strikethrough"),
            name: "Strikethrough",
            kind: MemberKind::PropertyGet,
        },
        ("excel.font.strikethrough", true) => MemberDescriptor {
            id: MemberId::new("excel.font.strikethrough"),
            name: "Strikethrough",
            kind: MemberKind::PropertyPut,
        },
        ("excel.font.color", false) => MemberDescriptor {
            id: MemberId::new("excel.font.color"),
            name: "Color",
            kind: MemberKind::PropertyGet,
        },
        ("excel.font.color", true) => MemberDescriptor {
            id: MemberId::new("excel.font.color"),
            name: "Color",
            kind: MemberKind::PropertyPut,
        },
        ("excel.font.colorindex", false) => MemberDescriptor {
            id: MemberId::new("excel.font.colorindex"),
            name: "ColorIndex",
            kind: MemberKind::PropertyGet,
        },
        ("excel.font.colorindex", true) => MemberDescriptor {
            id: MemberId::new("excel.font.colorindex"),
            name: "ColorIndex",
            kind: MemberKind::PropertyPut,
        },
        ("excel.interior.color", false) => MemberDescriptor {
            id: MemberId::new("excel.interior.color"),
            name: "Color",
            kind: MemberKind::PropertyGet,
        },
        ("excel.interior.color", true) => MemberDescriptor {
            id: MemberId::new("excel.interior.color"),
            name: "Color",
            kind: MemberKind::PropertyPut,
        },
        ("excel.interior.colorindex", false) => MemberDescriptor {
            id: MemberId::new("excel.interior.colorindex"),
            name: "ColorIndex",
            kind: MemberKind::PropertyGet,
        },
        ("excel.interior.colorindex", true) => MemberDescriptor {
            id: MemberId::new("excel.interior.colorindex"),
            name: "ColorIndex",
            kind: MemberKind::PropertyPut,
        },
        ("excel.interior.pattern", false) => MemberDescriptor {
            id: MemberId::new("excel.interior.pattern"),
            name: "Pattern",
            kind: MemberKind::PropertyGet,
        },
        ("excel.interior.pattern", true) => MemberDescriptor {
            id: MemberId::new("excel.interior.pattern"),
            name: "Pattern",
            kind: MemberKind::PropertyPut,
        },
        ("excel.interior.patterncolor", false) => MemberDescriptor {
            id: MemberId::new("excel.interior.patterncolor"),
            name: "PatternColor",
            kind: MemberKind::PropertyGet,
        },
        ("excel.interior.patterncolor", true) => MemberDescriptor {
            id: MemberId::new("excel.interior.patterncolor"),
            name: "PatternColor",
            kind: MemberKind::PropertyPut,
        },
        ("excel.interior.patterncolorindex", false) => MemberDescriptor {
            id: MemberId::new("excel.interior.patterncolorindex"),
            name: "PatternColorIndex",
            kind: MemberKind::PropertyGet,
        },
        ("excel.interior.patterncolorindex", true) => MemberDescriptor {
            id: MemberId::new("excel.interior.patterncolorindex"),
            name: "PatternColorIndex",
            kind: MemberKind::PropertyPut,
        },
        ("excel.borders.count", _) => MemberDescriptor {
            id: MemberId::new("excel.borders.count"),
            name: "Count",
            kind: MemberKind::PropertyGet,
        },
        ("excel.borders.item", _) => MemberDescriptor {
            id: MemberId::new("excel.borders.item"),
            name: "Item",
            kind: MemberKind::PropertyGet,
        },
        ("excel.borders.newenum", _) => MemberDescriptor {
            id: MemberId::new("excel.borders.newenum"),
            name: "_NewEnum",
            kind: MemberKind::PropertyGet,
        },
        ("excel.borders.color", false) => MemberDescriptor {
            id: MemberId::new("excel.borders.color"),
            name: "Color",
            kind: MemberKind::PropertyGet,
        },
        ("excel.borders.color", true) => MemberDescriptor {
            id: MemberId::new("excel.borders.color"),
            name: "Color",
            kind: MemberKind::PropertyPut,
        },
        ("excel.borders.colorindex", false) => MemberDescriptor {
            id: MemberId::new("excel.borders.colorindex"),
            name: "ColorIndex",
            kind: MemberKind::PropertyGet,
        },
        ("excel.borders.colorindex", true) => MemberDescriptor {
            id: MemberId::new("excel.borders.colorindex"),
            name: "ColorIndex",
            kind: MemberKind::PropertyPut,
        },
        ("excel.borders.linestyle", false) => MemberDescriptor {
            id: MemberId::new("excel.borders.linestyle"),
            name: "LineStyle",
            kind: MemberKind::PropertyGet,
        },
        ("excel.borders.linestyle", true) => MemberDescriptor {
            id: MemberId::new("excel.borders.linestyle"),
            name: "LineStyle",
            kind: MemberKind::PropertyPut,
        },
        ("excel.borders.weight", false) => MemberDescriptor {
            id: MemberId::new("excel.borders.weight"),
            name: "Weight",
            kind: MemberKind::PropertyGet,
        },
        ("excel.borders.weight", true) => MemberDescriptor {
            id: MemberId::new("excel.borders.weight"),
            name: "Weight",
            kind: MemberKind::PropertyPut,
        },
        ("excel.border.color", false) => MemberDescriptor {
            id: MemberId::new("excel.border.color"),
            name: "Color",
            kind: MemberKind::PropertyGet,
        },
        ("excel.border.color", true) => MemberDescriptor {
            id: MemberId::new("excel.border.color"),
            name: "Color",
            kind: MemberKind::PropertyPut,
        },
        ("excel.border.colorindex", false) => MemberDescriptor {
            id: MemberId::new("excel.border.colorindex"),
            name: "ColorIndex",
            kind: MemberKind::PropertyGet,
        },
        ("excel.border.colorindex", true) => MemberDescriptor {
            id: MemberId::new("excel.border.colorindex"),
            name: "ColorIndex",
            kind: MemberKind::PropertyPut,
        },
        ("excel.border.linestyle", false) => MemberDescriptor {
            id: MemberId::new("excel.border.linestyle"),
            name: "LineStyle",
            kind: MemberKind::PropertyGet,
        },
        ("excel.border.linestyle", true) => MemberDescriptor {
            id: MemberId::new("excel.border.linestyle"),
            name: "LineStyle",
            kind: MemberKind::PropertyPut,
        },
        ("excel.border.weight", false) => MemberDescriptor {
            id: MemberId::new("excel.border.weight"),
            name: "Weight",
            kind: MemberKind::PropertyGet,
        },
        ("excel.border.weight", true) => MemberDescriptor {
            id: MemberId::new("excel.border.weight"),
            name: "Weight",
            kind: MemberKind::PropertyPut,
        },
        ("excel.areas.count", _) => AREAS_COUNT,
        ("excel.areas.item", _) => AREAS_ITEM,
        ("excel.areas.newenum", _) => AREAS_NEW_ENUM,
        ("excel.names.count", _) => NAMES_COUNT,
        ("excel.names.item", _) => NAMES_ITEM,
        ("excel.names.newenum", _) => NAMES_NEW_ENUM,
        ("excel.names.add", _) => NAMES_ADD,
        ("excel.name.name", _) => NAME_NAME,
        ("excel.name.refersto", _) => NAME_REFERS_TO,
        ("excel.name.referstor1c1", _) => NAME_REFERS_TO_R1C1,
        ("excel.name.referstorange", _) => NAME_REFERS_TO_RANGE,
        ("excel.name.visible", _) => NAME_VISIBLE,
        ("excel.name.delete", _) => NAME_DELETE,
        _ => unreachable!("implemented member ID must be registered"),
    }
}

fn structured_data_member(id: MemberId, put: bool) -> Option<MemberDescriptor> {
    let (name, kind) = match (id.as_str(), put) {
        ("excel.worksheet.listobjects", _) => ("ListObjects", MemberKind::PropertyGet),
        ("excel.worksheet.autofiltermode", _) => ("AutoFilterMode", MemberKind::PropertyGet),
        ("excel.worksheet.filtermode", _) => ("FilterMode", MemberKind::PropertyGet),
        ("excel.worksheet.showalldata", _) => ("ShowAllData", MemberKind::Method),
        ("excel.worksheet.autofilter-3289", _) => ("AutoFilter", MemberKind::PropertyGet),
        ("excel.range.worksheet", _) => ("Worksheet", MemberKind::PropertyGet),
        ("excel.listobjects.count", _) => ("Count", MemberKind::PropertyGet),
        ("excel.listobjects.item", _) => ("Item", MemberKind::PropertyGet),
        ("excel.listobjects.newenum", _) => ("_NewEnum", MemberKind::PropertyGet),
        ("excel.listobjects.add", _) => ("Add", MemberKind::Method),
        ("excel.listobject.name", false) => ("Name", MemberKind::PropertyGet),
        ("excel.listobject.name", true) => ("Name", MemberKind::PropertyPut),
        ("excel.listobject.displayname", false) => ("DisplayName", MemberKind::PropertyGet),
        ("excel.listobject.displayname", true) => ("DisplayName", MemberKind::PropertyPut),
        ("excel.listobject.range", _) => ("Range", MemberKind::PropertyGet),
        ("excel.listobject.headerrowrange", _) => ("HeaderRowRange", MemberKind::PropertyGet),
        ("excel.listobject.databodyrange", _) => ("DataBodyRange", MemberKind::PropertyGet),
        ("excel.listobject.totalsrowrange", _) => ("TotalsRowRange", MemberKind::PropertyGet),
        ("excel.listobject.insertrowrange", _) => ("InsertRowRange", MemberKind::PropertyGet),
        ("excel.listobject.showheaders", false) => ("ShowHeaders", MemberKind::PropertyGet),
        ("excel.listobject.showheaders", true) => ("ShowHeaders", MemberKind::PropertyPut),
        ("excel.listobject.showtotals", false) => ("ShowTotals", MemberKind::PropertyGet),
        ("excel.listobject.showtotals", true) => ("ShowTotals", MemberKind::PropertyPut),
        ("excel.listobject.showautofilter", false) => ("ShowAutoFilter", MemberKind::PropertyGet),
        ("excel.listobject.showautofilter", true) => ("ShowAutoFilter", MemberKind::PropertyPut),
        ("excel.listobject.tablestyle", false) => ("TableStyle", MemberKind::PropertyGet),
        ("excel.listobject.tablestyle", true) => ("TableStyle", MemberKind::PropertyPut),
        ("excel.listobject.listcolumns", _) => ("ListColumns", MemberKind::PropertyGet),
        ("excel.listobject.listrows", _) => ("ListRows", MemberKind::PropertyGet),
        ("excel.listobject.autofilter-3289", _) => ("AutoFilter", MemberKind::PropertyGet),
        ("excel.listobject.sort-3288", _) => ("Sort", MemberKind::PropertyGet),
        ("excel.listobject.resize", _) => ("Resize", MemberKind::Method),
        ("excel.listobject.delete", _) => ("Delete", MemberKind::Method),
        ("excel.listobject.unlist", _) => ("Unlist", MemberKind::Method),
        ("excel.listcolumns.count", _) => ("Count", MemberKind::PropertyGet),
        ("excel.listcolumns.item", _) => ("Item", MemberKind::PropertyGet),
        ("excel.listcolumns.newenum", _) => ("_NewEnum", MemberKind::PropertyGet),
        ("excel.listcolumns.add", _) => ("Add", MemberKind::Method),
        ("excel.listcolumn.index", _) => ("Index", MemberKind::PropertyGet),
        ("excel.listcolumn.name", false) => ("Name", MemberKind::PropertyGet),
        ("excel.listcolumn.name", true) => ("Name", MemberKind::PropertyPut),
        ("excel.listcolumn.range", _) => ("Range", MemberKind::PropertyGet),
        ("excel.listcolumn.databodyrange", _) => ("DataBodyRange", MemberKind::PropertyGet),
        ("excel.listcolumn.total", _) => ("Total", MemberKind::PropertyGet),
        ("excel.listcolumn.totalscalculation", false) => {
            ("TotalsCalculation", MemberKind::PropertyGet)
        }
        ("excel.listcolumn.totalscalculation", true) => {
            ("TotalsCalculation", MemberKind::PropertyPut)
        }
        ("excel.listcolumn.delete", _) => ("Delete", MemberKind::Method),
        ("excel.listrows.count", _) => ("Count", MemberKind::PropertyGet),
        ("excel.listrows.item", _) => ("Item", MemberKind::PropertyGet),
        ("excel.listrows.newenum", _) => ("_NewEnum", MemberKind::PropertyGet),
        ("excel.listrows.add", _) => ("Add", MemberKind::Method),
        ("excel.listrow.index", _) => ("Index", MemberKind::PropertyGet),
        ("excel.listrow.range", _) => ("Range", MemberKind::PropertyGet),
        ("excel.listrow.delete", _) => ("Delete", MemberKind::Method),
        ("excel.autofilter.range", _) => ("Range", MemberKind::PropertyGet),
        ("excel.autofilter.filters", _) => ("Filters", MemberKind::PropertyGet),
        ("excel.filters.count", _) => ("Count", MemberKind::PropertyGet),
        ("excel.filters.item", _) => ("Item", MemberKind::PropertyGet),
        ("excel.filters.newenum", _) => ("_NewEnum", MemberKind::PropertyGet),
        ("excel.filter.on", _) => ("On", MemberKind::PropertyGet),
        ("excel.filter.operator-2641", _) => ("_Operator", MemberKind::PropertyGet),
        ("excel.filter.criteria1", _) => ("Criteria1", MemberKind::PropertyGet),
        ("excel.filter.criteria2", _) => ("Criteria2", MemberKind::PropertyGet),
        ("excel.range.autofilter-3289", _) => ("AutoFilter", MemberKind::Method),
        ("excel.range.sort", _) => ("_Sort", MemberKind::Method),
        ("excel.sort.sortfields", _) => ("SortFields", MemberKind::PropertyGet),
        ("excel.sort.setrange", _) => ("SetRange", MemberKind::Method),
        ("excel.sort.header", false) => ("Header", MemberKind::PropertyGet),
        ("excel.sort.header", true) => ("Header", MemberKind::PropertyPut),
        ("excel.sort.matchcase", false) => ("MatchCase", MemberKind::PropertyGet),
        ("excel.sort.matchcase", true) => ("MatchCase", MemberKind::PropertyPut),
        ("excel.sort.orientation", false) => ("Orientation", MemberKind::PropertyGet),
        ("excel.sort.orientation", true) => ("Orientation", MemberKind::PropertyPut),
        ("excel.sort.apply", _) => ("Apply", MemberKind::Method),
        ("excel.sortfields.count", _) => ("Count", MemberKind::PropertyGet),
        ("excel.sortfields.clear", _) => ("Clear", MemberKind::Method),
        ("excel.sortfields.add", _) => ("Add", MemberKind::Method),
        ("excel.range.validation", _) => ("Validation", MemberKind::PropertyGet),
        ("excel.validation.add", _) => ("Add", MemberKind::Method),
        ("excel.validation.delete", _) => ("Delete", MemberKind::Method),
        ("excel.validation.value", _) => ("Value", MemberKind::PropertyGet),
        ("excel.validation.type", _) => ("Type", MemberKind::PropertyGet),
        ("excel.validation.alertstyle", _) => ("AlertStyle", MemberKind::PropertyGet),
        ("excel.validation.operator", _) => ("Operator", MemberKind::PropertyGet),
        ("excel.validation.formula1", _) => ("Formula1", MemberKind::PropertyGet),
        ("excel.validation.formula2", _) => ("Formula2", MemberKind::PropertyGet),
        ("excel.validation.ignoreblank", false) => ("IgnoreBlank", MemberKind::PropertyGet),
        ("excel.validation.ignoreblank", true) => ("IgnoreBlank", MemberKind::PropertyPut),
        ("excel.validation.incelldropdown", false) => ("InCellDropdown", MemberKind::PropertyGet),
        ("excel.validation.incelldropdown", true) => ("InCellDropdown", MemberKind::PropertyPut),
        ("excel.validation.showinput", false) => ("ShowInput", MemberKind::PropertyGet),
        ("excel.validation.showinput", true) => ("ShowInput", MemberKind::PropertyPut),
        ("excel.validation.showerror", false) => ("ShowError", MemberKind::PropertyGet),
        ("excel.validation.showerror", true) => ("ShowError", MemberKind::PropertyPut),
        ("excel.validation.inputtitle", false) => ("InputTitle", MemberKind::PropertyGet),
        ("excel.validation.inputtitle", true) => ("InputTitle", MemberKind::PropertyPut),
        ("excel.validation.inputmessage", false) => ("InputMessage", MemberKind::PropertyGet),
        ("excel.validation.inputmessage", true) => ("InputMessage", MemberKind::PropertyPut),
        ("excel.validation.errortitle", false) => ("ErrorTitle", MemberKind::PropertyGet),
        ("excel.validation.errortitle", true) => ("ErrorTitle", MemberKind::PropertyPut),
        ("excel.validation.errormessage", false) => ("ErrorMessage", MemberKind::PropertyGet),
        ("excel.validation.errormessage", true) => ("ErrorMessage", MemberKind::PropertyPut),
        ("excel.range.removeduplicates", _) => ("RemoveDuplicates", MemberKind::Method),
        ("excel.range.currentregion", _) => ("CurrentRegion", MemberKind::PropertyGet),
        ("excel.range.hidden", false) => ("Hidden", MemberKind::PropertyGet),
        ("excel.range.hidden", true) => ("Hidden", MemberKind::PropertyPut),
        ("excel.range.insert", _) => ("Insert", MemberKind::Method),
        ("excel.range.delete", _) => ("Delete", MemberKind::Method),
        ("excel.range.clear", _) => ("Clear", MemberKind::Method),
        ("excel.range.clearformats", _) => ("ClearFormats", MemberKind::Method),
        ("excel.range.clearcomments", _) => ("ClearComments", MemberKind::Method),
        ("excel.range.clearhyperlinks", _) => ("ClearHyperlinks", MemberKind::Method),
        ("excel.range.copy", _) => ("Copy", MemberKind::Method),
        ("excel.range.cut", _) => ("Cut", MemberKind::Method),
        ("excel.range.pastespecial-1928", _) => ("PasteSpecial", MemberKind::Method),
        _ => return None,
    };
    Some(MemberDescriptor { id, name, kind })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn descriptor_lookup_is_centralized() {
        let descriptor = member(MemberId::new("excel.workbooks.add"), false);
        assert_eq!(descriptor.name, "Add");
        assert_eq!(descriptor.kind, MemberKind::PropertyGet);
    }
}
