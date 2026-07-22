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
    "excel.range.clearcontents",
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
];

pub(crate) fn member(id: MemberId, put: bool) -> MemberDescriptor {
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
        ("excel.range.clearcontents", _) => RANGE_CLEAR_CONTENTS,
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
