use std::fmt::{Debug, Formatter};
use std::path::Path;

/// Numeric value from Excel's `XlFileFormat` enumeration.
#[repr(transparent)]
#[derive(Clone, Copy, Eq, Hash, PartialEq)]
pub struct XlFileFormat(i32);

impl XlFileFormat {
    /// Excel Binary Workbook (`.xlsb`).
    pub const EXCEL12: Self = Self(50);
    /// Open XML Workbook (`.xlsx`).
    pub const OPEN_XML_WORKBOOK: Self = Self(51);
    /// Open XML macro-enabled workbook (`.xlsm`).
    pub const OPEN_XML_WORKBOOK_MACRO_ENABLED: Self = Self(52);
    /// Open XML template (`.xltx`).
    pub const OPEN_XML_TEMPLATE: Self = Self(54);
    /// OpenDocument Spreadsheet (`.ods`).
    pub const OPEN_DOCUMENT_SPREADSHEET: Self = Self(60);
    /// Comma-separated values (`.csv`).
    pub const CSV: Self = Self(6);
    /// UTF-8 comma-separated values (`.csv`).
    pub const CSV_UTF8: Self = Self(62);
    /// Windows text (`.txt`).
    pub const TEXT_WINDOWS: Self = Self(20);

    /// Builds a value retained from Excel without discarding unknown values.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the numeric Excel enumeration value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

impl Debug for XlFileFormat {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("XlFileFormat")
            .field(&self.0)
            .finish()
    }
}

/// Numeric value from Excel's `XlSaveAsAccessMode` enumeration.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct XlSaveAsAccessMode(i32);

impl XlSaveAsAccessMode {
    /// Leaves the workbook's current access mode unchanged.
    pub const NO_CHANGE: Self = Self(1);
    /// Saves as a shared workbook.
    pub const SHARED: Self = Self(2);
    /// Saves as an exclusive workbook.
    pub const EXCLUSIVE: Self = Self(3);

    /// Builds a value retained from Excel without discarding unknown values.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the numeric Excel enumeration value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Numeric value from Excel's `XlSaveConflictResolution` enumeration.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct XlSaveConflictResolution(i32);

impl XlSaveConflictResolution {
    /// Lets Excel ask the user to resolve a conflict.
    pub const USER_RESOLUTION: Self = Self(1);
    /// Keeps this session's changes.
    pub const LOCAL_SESSION_CHANGES: Self = Self(2);
    /// Keeps another session's changes.
    pub const OTHER_SESSION_CHANGES: Self = Self(3);

    /// Builds a value retained from Excel without discarding unknown values.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the numeric Excel enumeration value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Numeric value from Excel's `XlUpdateLinks` enumeration.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct XlUpdateLinks(i32);

impl XlUpdateLinks {
    /// Uses the user's current update-link setting.
    pub const USER_SETTING: Self = Self(1);
    /// Does not update external links.
    pub const NEVER: Self = Self(2);
    /// Always updates external links.
    pub const ALWAYS: Self = Self(3);

    /// Returns the numeric Excel enumeration value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Numeric value for `Workbooks.Open` text-format interpretation.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct WorkbookOpenFormat(i32);

impl WorkbookOpenFormat {
    /// Lets Excel choose the normal workbook format.
    pub const NORMAL: Self = Self(-4143);
    /// Treats the input as an Excel template.
    pub const TEMPLATE: Self = Self(17);
    /// Uses Excel's current default workbook format.
    pub const DEFAULT: Self = Self(51);

    /// Returns the numeric Excel value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Numeric value from Excel's `XlPlatform` enumeration.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct XlPlatform(i32);

impl XlPlatform {
    /// Macintosh text-file origin.
    pub const MACINTOSH: Self = Self(1);
    /// Windows text-file origin.
    pub const WINDOWS: Self = Self(2);
    /// MS-DOS text-file origin.
    pub const MSDOS: Self = Self(3);

    /// Returns the numeric Excel enumeration value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Numeric value from Excel's `XlCorruptLoad` enumeration.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct XlCorruptLoad(i32);

impl XlCorruptLoad {
    /// Opens normally.
    pub const NORMAL: Self = Self(0);
    /// Asks Excel to repair a damaged file.
    pub const REPAIR: Self = Self(1);
    /// Asks Excel to extract data from a damaged file.
    pub const EXTRACT_DATA: Self = Self(2);

    /// Returns the numeric Excel enumeration value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Typed optional positions accepted by [`Workbooks::open`](super::Workbooks::open).
#[derive(Clone, Default, PartialEq)]
pub struct WorkbookOpenOptions<'a> {
    /// Controls external-link updates.
    pub update_links: Option<XlUpdateLinks>,
    /// Opens the workbook read-only when `true`.
    pub read_only: Option<bool>,
    /// Selects Excel's input-format interpretation.
    pub format: Option<WorkbookOpenFormat>,
    /// Optional open password. Debug output always redacts it.
    pub password: Option<&'a str>,
    /// Optional write-reservation password. Debug output always redacts it.
    pub write_res_password: Option<&'a str>,
    /// Ignores a workbook's read-only recommendation when `true`.
    pub ignore_read_only_recommended: Option<bool>,
    /// Selects text-file platform origin.
    pub origin: Option<XlPlatform>,
    /// Supplies a one-character text-file delimiter.
    pub delimiter: Option<char>,
    /// Opens an editable template when `true`.
    pub editable: Option<bool>,
    /// Requests notification if the file is unavailable when `true`.
    pub notify: Option<bool>,
    /// Selects an installed file converter by index.
    pub converter: Option<i32>,
    /// Adds the workbook to Excel's most-recently-used list when `true`.
    pub add_to_mru: Option<bool>,
    /// Uses local language settings when `true`.
    pub local: Option<bool>,
    /// Selects corrupt-file handling.
    pub corrupt_load: Option<XlCorruptLoad>,
}

impl<'a> WorkbookOpenOptions<'a> {
    /// Starts an all-`Missing` Open call after its required filename.
    pub const fn new() -> Self {
        Self {
            update_links: None,
            read_only: None,
            format: None,
            password: None,
            write_res_password: None,
            ignore_read_only_recommended: None,
            origin: None,
            delimiter: None,
            editable: None,
            notify: None,
            converter: None,
            add_to_mru: None,
            local: None,
            corrupt_load: None,
        }
    }
}

impl Debug for WorkbookOpenOptions<'_> {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("WorkbookOpenOptions")
            .field("update_links", &self.update_links)
            .field("read_only", &self.read_only)
            .field("format", &self.format)
            .field("password", &self.password.map(|_| "REDACTED"))
            .field(
                "write_res_password",
                &self.write_res_password.map(|_| "REDACTED"),
            )
            .field(
                "ignore_read_only_recommended",
                &self.ignore_read_only_recommended,
            )
            .field("origin", &self.origin)
            .field("delimiter", &self.delimiter)
            .field("editable", &self.editable)
            .field("notify", &self.notify)
            .field("converter", &self.converter)
            .field("add_to_mru", &self.add_to_mru)
            .field("local", &self.local)
            .field("corrupt_load", &self.corrupt_load)
            .finish()
    }
}

/// Typed optional positions accepted by [`Workbook::save_as`](super::Workbook::save_as).
#[derive(Clone, Default, PartialEq)]
pub struct WorkbookSaveAsOptions<'a> {
    /// Selects the destination workbook file format.
    pub file_format: Option<XlFileFormat>,
    /// Optional open password. Debug output always redacts it.
    pub password: Option<&'a str>,
    /// Optional write-reservation password. Debug output always redacts it.
    pub write_res_password: Option<&'a str>,
    /// Recommends read-only opening when `true`.
    pub read_only_recommended: Option<bool>,
    /// Creates a backup when `true`.
    pub create_backup: Option<bool>,
    /// Selects shared or exclusive workbook access.
    pub access_mode: Option<XlSaveAsAccessMode>,
    /// Selects save-conflict behavior.
    pub conflict_resolution: Option<XlSaveConflictResolution>,
    /// Adds the file to Excel's most-recently-used list when `true`.
    pub add_to_mru: Option<bool>,
    /// Text-codepage value for text formats.
    pub text_codepage: Option<i32>,
    /// Text visual-layout value for text formats.
    pub text_visual_layout: Option<i32>,
    /// Uses local language settings when `true`.
    pub local: Option<bool>,
}

impl<'a> WorkbookSaveAsOptions<'a> {
    /// Starts an all-`Missing` SaveAs call after its required filename.
    pub const fn new() -> Self {
        Self {
            file_format: None,
            password: None,
            write_res_password: None,
            read_only_recommended: None,
            create_backup: None,
            access_mode: None,
            conflict_resolution: None,
            add_to_mru: None,
            text_codepage: None,
            text_visual_layout: None,
            local: None,
        }
    }
}

impl Debug for WorkbookSaveAsOptions<'_> {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("WorkbookSaveAsOptions")
            .field("file_format", &self.file_format)
            .field("password", &self.password.map(|_| "REDACTED"))
            .field(
                "write_res_password",
                &self.write_res_password.map(|_| "REDACTED"),
            )
            .field("read_only_recommended", &self.read_only_recommended)
            .field("create_backup", &self.create_backup)
            .field("access_mode", &self.access_mode)
            .field("conflict_resolution", &self.conflict_resolution)
            .field("add_to_mru", &self.add_to_mru)
            .field("text_codepage", &self.text_codepage)
            .field("text_visual_layout", &self.text_visual_layout)
            .field("local", &self.local)
            .finish()
    }
}

/// Specifies whether [`Workbook::close`](super::Workbook::close) saves first.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum SaveChanges {
    /// Sends an explicit `Missing`, preserving Excel's prompt behavior.
    #[default]
    Prompt,
    /// Sends `true` and asks Excel to save before closing.
    Save,
    /// Sends `false` and discards unsaved changes.
    Discard,
}

/// Typed optional positions accepted by [`Workbook::close`](super::Workbook::close).
#[derive(Clone, Debug, Default, PartialEq)]
pub struct WorkbookCloseOptions<'a> {
    /// Whether Excel should prompt, save, or discard changes.
    pub save_changes: SaveChanges,
    /// Optional replacement filename, preserved as caller-supplied Windows path units.
    pub filename: Option<&'a Path>,
    /// Optional routing-workbook flag.
    pub route_workbook: Option<bool>,
}

impl<'a> WorkbookCloseOptions<'a> {
    /// Starts a close request that preserves Excel's default prompt behavior.
    pub const fn new() -> Self {
        Self {
            save_changes: SaveChanges::Prompt,
            filename: None,
            route_workbook: None,
        }
    }
}
