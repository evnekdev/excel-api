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
const WORKBOOKS_COUNT: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbooks.count"),
    name: "Count",
    kind: MemberKind::PropertyGet,
};
// Microsoft's C++ reference and runtime evidence classify Add as PROPERTYGET.
const WORKBOOKS_ADD: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbooks.add"),
    name: "Add",
    kind: MemberKind::PropertyGet,
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
const WORKBOOK_WORKSHEETS: MemberDescriptor = MemberDescriptor {
    id: MemberId::new("excel.workbook.worksheets"),
    name: "Worksheets",
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
    "excel.application.workbooks",
    "excel.application.quit",
    "excel.workbooks.count",
    "excel.workbooks.add",
    "excel.workbook.name",
    "excel.workbook.saved",
    "excel.workbook.close",
    "excel.workbook.worksheets",
    "excel.worksheets.count",
    "excel.worksheets.item",
    "excel.worksheets.add",
    "excel.worksheet.name",
    "excel.worksheet.index",
    "excel.worksheet.visible",
    "excel.worksheet.range",
    "excel.worksheet.usedrange",
    "excel.range.address",
    "excel.range.row",
    "excel.range.column",
    "excel.range.count",
    "excel.range.rows",
    "excel.range.columns",
    "excel.range.value",
    "excel.range.value2",
    "excel.range.formula",
    "excel.range.formula2",
    "excel.range.clearcontents",
];

pub(crate) fn member(id: MemberId, put: bool) -> MemberDescriptor {
    match (id.as_str(), put) {
        ("excel.application.version", _) => APPLICATION_VERSION,
        ("excel.application.visible", false) => APPLICATION_VISIBLE_GET,
        ("excel.application.visible", true) => APPLICATION_VISIBLE_PUT,
        ("excel.application.workbooks", _) => APPLICATION_WORKBOOKS,
        ("excel.application.quit", _) => APPLICATION_QUIT,
        ("excel.workbooks.count", _) => WORKBOOKS_COUNT,
        ("excel.workbooks.add", _) => WORKBOOKS_ADD,
        ("excel.workbook.name", _) => WORKBOOK_NAME,
        ("excel.workbook.saved", false) => WORKBOOK_SAVED_GET,
        ("excel.workbook.saved", true) => WORKBOOK_SAVED_PUT,
        ("excel.workbook.close", _) => WORKBOOK_CLOSE,
        ("excel.workbook.worksheets", _) => WORKBOOK_WORKSHEETS,
        ("excel.worksheets.count", _) => WORKSHEETS_COUNT,
        ("excel.worksheets.item", _) => WORKSHEETS_ITEM,
        ("excel.worksheets.add", _) => WORKSHEETS_ADD,
        ("excel.worksheet.name", false) => WORKSHEET_NAME_GET,
        ("excel.worksheet.name", true) => WORKSHEET_NAME_PUT,
        ("excel.worksheet.index", _) => WORKSHEET_INDEX,
        ("excel.worksheet.visible", false) => WORKSHEET_VISIBLE_GET,
        ("excel.worksheet.visible", true) => WORKSHEET_VISIBLE_PUT,
        ("excel.worksheet.range", _) => WORKSHEET_RANGE,
        ("excel.worksheet.usedrange", _) => WORKSHEET_USED_RANGE,
        ("excel.range.address", _) => RANGE_ADDRESS,
        ("excel.range.row", _) => RANGE_ROW,
        ("excel.range.column", _) => RANGE_COLUMN,
        ("excel.range.count", _) => RANGE_COUNT,
        ("excel.range.rows", _) => RANGE_ROWS,
        ("excel.range.columns", _) => RANGE_COLUMNS,
        ("excel.range.value", false) => RANGE_VALUE_GET,
        ("excel.range.value", true) => RANGE_VALUE_PUT,
        ("excel.range.value2", false) => RANGE_VALUE2_GET,
        ("excel.range.value2", true) => RANGE_VALUE2_PUT,
        ("excel.range.formula", false) => RANGE_FORMULA_GET,
        ("excel.range.formula", true) => RANGE_FORMULA_PUT,
        ("excel.range.formula2", false) => RANGE_FORMULA2_GET,
        ("excel.range.formula2", true) => RANGE_FORMULA2_PUT,
        ("excel.range.clearcontents", _) => RANGE_CLEAR_CONTENTS,
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
