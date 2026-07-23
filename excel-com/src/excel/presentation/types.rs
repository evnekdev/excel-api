//! Transparent Excel constants and small presentation option types.

use super::*;

macro_rules! excel_i32 {
    ($(#[$meta:meta])* $name:ident { $($(#[$constant_meta:meta])* $constant:ident = $value:expr;)* }) => {
        $(#[$meta])*
        #[repr(transparent)]
        #[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
        pub struct $name(i32);
        impl $name {
            $($(#[$constant_meta])* pub const $constant: Self = Self($value);)*
            /// Builds the value from Excel's raw integer without discarding an unknown value.
            pub const fn from_raw(value: i32) -> Self { Self(value) }
            /// Returns the raw integer passed to or read from Excel.
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}

excel_i32! {
    /// A forward-compatible Excel window view.
    WindowView {
        /// Normal worksheet view.
        NORMAL = 1;
        /// Page-break preview.
        PAGE_BREAK_PREVIEW = 2;
        /// Page-layout view.
        PAGE_LAYOUT = 3;
    }
}

/// Compatibility name for a sheet-window view value.
pub type SheetView = WindowView;

excel_i32! {
    /// A forward-compatible page orientation.
    PageOrientation {
        /// Portrait orientation.
        PORTRAIT = 1;
        /// Landscape orientation.
        LANDSCAPE = 2;
    }
}

excel_i32! {
    /// A forward-compatible paper-size value.
    PaperSize {
        /// Letter paper.
        LETTER = 1;
        /// Legal paper.
        LEGAL = 5;
        /// A4 paper.
        A4 = 9;
        /// Excel's printer-defined paper size.
        USER = 256;
    }
}

excel_i32! {
    /// Page traversal order.
    PrintOrder {
        /// Print down rows before moving across columns.
        DOWN_THEN_OVER = 1;
        /// Print across columns before moving down rows.
        OVER_THEN_DOWN = 2;
    }
}

excel_i32! {
    /// Header/footer print location for comments.
    PrintLocation {
        /// Do not print comments.
        NO_COMMENTS = -4142;
        /// Print comments at the end of the sheet.
        SHEET_END = 1;
        /// Print comments in place.
        IN_PLACE = 16;
    }
}

excel_i32! {
    /// How Excel renders cell errors in printed output.
    PrintErrors {
        /// Display errors as shown in the worksheet.
        DISPLAYED = 0;
        /// Print errors as blank cells.
        BLANK = 1;
        /// Print errors as dashes.
        DASH = 2;
        /// Print errors as `#N/A`.
        NA = 3;
    }
}

excel_i32! {
    /// Excel's range reading-order setting.
    ReadingOrder {
        /// Let Excel choose from the current language context.
        CONTEXT = -5002;
        /// Left-to-right reading order.
        LEFT_TO_RIGHT = -5003;
        /// Right-to-left reading order.
        RIGHT_TO_LEFT = -5004;
    }
}

excel_i32! {
    /// Location of outline summary rows.
    SummaryRow {
        /// Summary rows appear above details.
        ABOVE = 0;
        /// Summary rows appear below details.
        BELOW = 1;
    }
}

excel_i32! {
    /// Location of outline summary columns.
    SummaryColumn {
        /// Summary columns appear to the left of details.
        LEFT = -4131;
        /// Summary columns appear to the right of details.
        RIGHT = -4152;
    }
}

excel_i32! {
    /// A manual or automatic page-break classification.
    PageBreakType {
        /// Excel chose the break automatically.
        AUTOMATIC = -4105;
        /// A user-controlled manual break.
        MANUAL = -4135;
        /// No break is present.
        NONE = -4142;
    }
}

excel_i32! {
    /// Output format for `ExportAsFixedFormat`.
    FixedFormatType {
        /// Portable Document Format.
        PDF = 0;
        /// XML Paper Specification.
        XPS = 1;
    }
}

excel_i32! {
    /// Output quality for `ExportAsFixedFormat`.
    FixedFormatQuality {
        /// Standard quality.
        STANDARD = 0;
        /// Smaller, minimum-quality output.
        MINIMUM = 1;
    }
}

excel_i32! {
    /// Excel's process-global macro Automation security setting.
    AutomationSecurity {
        /// Enable macros using Excel's low-security behavior.
        LOW = 1;
        /// Follow the user's Excel UI setting.
        BY_UI = 2;
        /// Disable macros while a file is opened through Automation.
        FORCE_DISABLE = 3;
    }
}

/// A page zoom representation that preserves Excel's `bool | number` contract.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum PageZoom {
    /// Excel's automatic zoom behavior (`false`).
    Automatic,
    /// A numeric zoom percentage.
    Percent(i32),
}

/// Explicit page-fitting dimensions.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct PageFit {
    /// Number of pages wide, or `None` to leave that dimension Excel-controlled.
    pub wide: Option<usize>,
    /// Number of pages tall, or `None` to leave that dimension Excel-controlled.
    pub tall: Option<usize>,
}

/// A typed destination for worksheet copy or move operations.
#[derive(Clone, Copy, Debug)]
pub enum SheetDestination<'a> {
    /// Place the resulting sheet before this worksheet.
    Before(&'a Worksheet),
    /// Place the resulting sheet after this worksheet.
    After(&'a Worksheet),
    /// Let Excel create a new workbook containing the resulting sheet.
    NewWorkbook,
}
