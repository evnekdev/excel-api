use excel_com::{Workbook, WorkbookCloseOptions};

fn use_after_close(workbook: Workbook) {
    let _ = workbook.close(WorkbookCloseOptions::new());
    let _ = workbook.name();
}

fn main() {}
