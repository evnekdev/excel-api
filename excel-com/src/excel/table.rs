//! Typed Excel table (`ListObject`) wrappers and their row/column collections.

use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;

use crate::ExcelComError;
use crate::automation::{
    EnumVariant, OwnedVariant, PositionalArguments, enumerated_dispatch, invoke, property_get,
    property_put,
};
use crate::excel::collection::{
    CollectionDescriptor, count as collection_count, enumerator, item_by_index,
};
use crate::excel::formula::FormulaMember;
use crate::excel::text::text_bstr;
use crate::excel::{AutoFilter, DispatchObject, Range, Sort, Worksheet};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

const LIST_OBJECTS: CollectionDescriptor = CollectionDescriptor {
    name: "ListObjects",
    count: MemberId::new("excel.listobjects.count"),
    item: MemberId::new("excel.listobjects.item"),
    new_enum: MemberId::new("excel.listobjects.newenum"),
};
const LIST_COLUMNS: CollectionDescriptor = CollectionDescriptor {
    name: "ListColumns",
    count: MemberId::new("excel.listcolumns.count"),
    item: MemberId::new("excel.listcolumns.item"),
    new_enum: MemberId::new("excel.listcolumns.newenum"),
};
const LIST_ROWS: CollectionDescriptor = CollectionDescriptor {
    name: "ListRows",
    count: MemberId::new("excel.listrows.count"),
    item: MemberId::new("excel.listrows.item"),
    new_enum: MemberId::new("excel.listrows.newenum"),
};

macro_rules! raw_excel_type {
    ($(#[$type_docs:meta])* $name:ident { $($constant:ident = $value:expr => $constant_docs:literal;)* }) => {
        $(#[$type_docs])*
        #[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
        #[repr(transparent)]
        pub struct $name(i32);

        impl $name {
            $(#[doc = $constant_docs] pub const $constant: Self = Self($value);)*

            /// Creates this value from an installed Excel type-library integer.
            pub const fn from_raw(value: i32) -> Self { Self(value) }

            /// Returns the type-library integer, including unknown future values.
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}

raw_excel_type! {
    /// A forward-compatible `XlListObjectSourceType` value.
    ListObjectSourceType {
        EXTERNAL = 0 => "`xlSrcExternal`. External sources are not created by this crate slice.";
        RANGE = 1 => "`xlSrcRange`, the supported source for [`ListObjects::add_from_range`].";
        XML = 2 => "`xlSrcXml`. XML-backed table creation is not implemented.";
        QUERY = 3 => "`xlSrcQuery`. Query-backed table creation is not implemented.";
        MODEL = 4 => "`xlSrcModel`. Data-model table creation is not implemented.";
    }
}

raw_excel_type! {
    /// A forward-compatible `XlYesNoGuess` header declaration for tables and sorting.
    TableHeaderMode {
        GUESS = 0 => "`xlGuess`: Excel decides whether the first row is a header.";
        YES = 1 => "`xlYes`: the first row is a header.";
        NO = 2 => "`xlNo`: the first row is ordinary data.";
    }
}

raw_excel_type! {
    /// A forward-compatible `XlTotalsCalculation` value for a table column.
    TotalsCalculation {
        NONE = 0 => "`xlTotalsCalculationNone`.";
        SUM = 1 => "`xlTotalsCalculationSum`.";
        AVERAGE = 2 => "`xlTotalsCalculationAverage`.";
        COUNT = 3 => "`xlTotalsCalculationCount`.";
        COUNT_NUMBERS = 4 => "`xlTotalsCalculationCountNums`.";
        MIN = 5 => "`xlTotalsCalculationMin`.";
        MAX = 6 => "`xlTotalsCalculationMax`.";
        STANDARD_DEVIATION = 7 => "`xlTotalsCalculationStdDev`.";
        VARIANCE = 8 => "`xlTotalsCalculationVar`.";
        CUSTOM = 9 => "`xlTotalsCalculationCustom`.";
    }
}

/// Narrow Range-backed input for [`ListObjects::add_from_range`].
///
/// Excel receives `xlSrcRange`, the source Range, a missing `LinkSource`, the
/// supplied header mode, the optional destination Range, and an optional style
/// name in their documented logical positions.  Multi-area sources are
/// rejected locally because Excel tables are rectangular.
#[derive(Debug)]
pub struct ListObjectAddOptions<'a> {
    /// The single-area source rectangle for the new table.
    pub source: &'a Range,
    /// Declares whether the first source row is a header.
    pub has_headers: TableHeaderMode,
    /// An optional worksheet destination retained in Excel's positional call.
    pub destination: Option<&'a Range>,
    /// An optional Excel table style name.
    pub table_style_name: Option<&'a str>,
}

/// Apartment-bound typed collection of worksheet tables.
pub struct ListObjects {
    inner: DispatchObject,
}
impl Debug for ListObjects {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("ListObjects")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for ListObjects {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ListObjects {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ListObjects",
            },
        }
    }

    /// Returns the current number of worksheet tables.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, LIST_OBJECTS)
    }

    /// Returns the one-based table at `index`.
    pub fn item_by_index(&self, index: usize) -> Result<ListObject, ExcelComError> {
        Ok(ListObject::from_dispatch(item_by_index(
            &self.inner,
            LIST_OBJECTS,
            index,
        )?))
    }

    /// Returns a table by its Excel-visible name.
    pub fn item_by_name(&self, name: &str) -> Result<ListObject, ExcelComError> {
        let mut value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.listobjects.item"), false),
            vec![text_bstr(name)?],
        )?;
        Ok(ListObject::from_dispatch(value.take_dispatch()?))
    }

    /// Creates one Range-backed Excel table.
    pub fn add_from_range(
        &self,
        options: &ListObjectAddOptions<'_>,
    ) -> Result<ListObject, ExcelComError> {
        if options.source.areas()?.count()? != 1 {
            return Err(ExcelComError::Unsupported {
                detail: "ListObjects.Add requires a single-area source Range",
            });
        }
        let mut arguments = PositionalArguments::new();
        arguments.push_optional(Some(OwnedVariant::i32(ListObjectSourceType::RANGE.raw())));
        arguments.push_object(options.source.dispatch_object());
        arguments.push_optional(None); // LinkSource
        arguments.push_optional(Some(OwnedVariant::i32(options.has_headers.raw())));
        arguments.push_optional_object(options.destination.map(Range::dispatch_object));
        match options.table_style_name {
            Some(name) => arguments.push_result(text_bstr(name))?,
            None => arguments.push_optional(None),
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.listobjects.add"), false),
            arguments.into_inner(),
            false,
        )?;
        Ok(ListObject::from_dispatch(value.take_dispatch()?))
    }

    /// Iterates tables in Excel's `_NewEnum` order.
    pub fn iter(&self) -> Result<ListObjectsIter, ExcelComError> {
        Ok(ListObjectsIter {
            enumerator: enumerator(&self.inner, LIST_OBJECTS)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Fallible, single-pass, apartment-bound iterator over [`ListObjects`].
pub struct ListObjectsIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for ListObjectsIter {
    type Item = Result<ListObject, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(
                    enumerated_dispatch(&mut value, "ListObjects", index)
                        .map(ListObject::from_dispatch),
                )
            }
            Ok(None) => {
                self.terminal = true;
                None
            }
            Err(error) => {
                self.terminal = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for ListObjectsIter {}

/// Apartment-bound wrapper for one Excel table (`ListObject`).
pub struct ListObject {
    inner: DispatchObject,
}
impl Debug for ListObject {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ListObject").field(&self.inner).finish()
    }
}
impl Clone for ListObject {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ListObject {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ListObject",
            },
        }
    }

    /// Returns the internal Excel table name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.listobject.name")
    }
    /// Changes the internal Excel table name after rejecting embedded NUL.
    pub fn set_name(&self, name: &str) -> Result<(), ExcelComError> {
        text_put(&self.inner, "excel.listobject.name", name)
    }
    /// Returns Excel's display name for the table.
    pub fn display_name(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.listobject.displayname")
    }
    /// Changes Excel's display name after rejecting embedded NUL.
    pub fn set_display_name(&self, name: &str) -> Result<(), ExcelComError> {
        text_put(&self.inner, "excel.listobject.displayname", name)
    }
    /// Returns the full table Range, including headers and a visible totals row.
    pub fn range(&self) -> Result<Range, ExcelComError> {
        range_get(&self.inner, "excel.listobject.range")
    }
    /// Returns the header Range, or no object when Excel omits headers.
    pub fn header_row_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range_get(&self.inner, "excel.listobject.headerrowrange")
    }
    /// Returns the data-body Range, or `None` for a header-only empty table.
    pub fn data_body_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range_get(&self.inner, "excel.listobject.databodyrange")
    }
    /// Returns the totals-row Range, or `None` while totals are hidden.
    pub fn totals_row_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range_get(&self.inner, "excel.listobject.totalsrowrange")
    }
    /// Returns Excel's insert-row Range when it is available.
    pub fn insert_row_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range_get(&self.inner, "excel.listobject.insertrowrange")
    }
    /// Returns whether table headers are displayed.
    pub fn show_headers(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.listobject.showheaders")
    }
    /// Shows or hides table headers.
    pub fn set_show_headers(&self, visible: bool) -> Result<(), ExcelComError> {
        bool_put(&self.inner, "excel.listobject.showheaders", visible)
    }
    /// Returns whether the totals row is displayed.
    pub fn show_totals(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.listobject.showtotals")
    }
    /// Shows or hides the totals row.
    pub fn set_show_totals(&self, visible: bool) -> Result<(), ExcelComError> {
        bool_put(&self.inner, "excel.listobject.showtotals", visible)
    }
    /// Returns whether the table's AutoFilter arrows are displayed.
    pub fn show_autofilter(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.listobject.showautofilter")
    }
    /// Shows or hides the table's AutoFilter arrows.
    pub fn set_show_autofilter(&self, visible: bool) -> Result<(), ExcelComError> {
        bool_put(&self.inner, "excel.listobject.showautofilter", visible)
    }
    /// Returns the applied Excel table-style name.
    pub fn table_style(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.listobject.tablestyle")
    }
    /// Applies an Excel table style by name.
    pub fn set_table_style(&self, style_name: &str) -> Result<(), ExcelComError> {
        text_put(&self.inner, "excel.listobject.tablestyle", style_name)
    }
    /// Returns the typed table-column collection.
    pub fn list_columns(&self) -> Result<ListColumns, ExcelComError> {
        object_get(
            &self.inner,
            "excel.listobject.listcolumns",
            ListColumns::from_dispatch,
        )
    }
    /// Returns the typed table-row collection.
    pub fn list_rows(&self) -> Result<ListRows, ExcelComError> {
        object_get(
            &self.inner,
            "excel.listobject.listrows",
            ListRows::from_dispatch,
        )
    }
    /// Returns this table's stateful AutoFilter object.
    pub fn auto_filter(&self) -> Result<AutoFilter, ExcelComError> {
        object_get(
            &self.inner,
            "excel.listobject.autofilter-3289",
            AutoFilter::from_dispatch,
        )
    }
    /// Returns this table's persistent Sort object.
    pub fn sort(&self) -> Result<Sort, ExcelComError> {
        object_get(
            &self.inner,
            "excel.listobject.sort-3288",
            Sort::from_dispatch,
        )
    }
    /// Resizes the table to a single-area Range; Excel validates overlap and content rules.
    pub fn resize(&self, new_range: &Range) -> Result<(), ExcelComError> {
        if new_range.areas()?.count()? != 1 {
            return Err(ExcelComError::Unsupported {
                detail: "ListObject.Resize requires a single-area Range",
            });
        }
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.listobject.resize"), false),
            vec![OwnedVariant::dispatch_borrowed(
                &new_range.dispatch_object().dispatch,
            )],
            false,
        )?;
        Ok(())
    }
    /// Deletes this table and consumes the wrapper because its target no longer exists.
    pub fn delete(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.listobject.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Converts this table to ordinary worksheet cells and consumes the table wrapper.
    ///
    /// The installed type library declares `Unlist` with no Range return; keep
    /// a separate [`Self::range`] handle before calling it when that Range is needed.
    pub fn unlist(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.listobject.unlist"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Returns whether Excel currently reports no data rows.
    pub fn is_empty(&self) -> Result<bool, ExcelComError> {
        Ok(self.list_rows()?.count()? == 0)
    }
    /// Returns the current data-row count.
    pub fn data_row_count(&self) -> Result<usize, ExcelComError> {
        self.list_rows()?.count()
    }
    /// Returns the current table-column count.
    pub fn column_count(&self) -> Result<usize, ExcelComError> {
        self.list_columns()?.count()
    }
}

/// Apartment-bound typed collection of columns in one Excel table.
pub struct ListColumns {
    inner: DispatchObject,
}
impl Debug for ListColumns {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ListColumns").field(&self.inner).finish()
    }
}
impl Clone for ListColumns {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ListColumns {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ListColumns",
            },
        }
    }
    /// Returns the number of columns.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, LIST_COLUMNS)
    }
    /// Returns the one-based column at `index`.
    pub fn item_by_index(&self, index: usize) -> Result<ListColumn, ExcelComError> {
        Ok(ListColumn::from_dispatch(item_by_index(
            &self.inner,
            LIST_COLUMNS,
            index,
        )?))
    }
    /// Returns a column by its Excel-visible header name.
    pub fn item_by_name(&self, name: &str) -> Result<ListColumn, ExcelComError> {
        let mut value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.listcolumns.item"), false),
            vec![text_bstr(name)?],
        )?;
        Ok(ListColumn::from_dispatch(value.take_dispatch()?))
    }
    /// Appends a column when `position` is `None`, otherwise inserts it at a one-based position.
    pub fn add(&self, position: Option<usize>) -> Result<ListColumn, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_optional(
            position
                .map(|value| one_based(value, "ListColumns.Add position"))
                .transpose()?
                .map(OwnedVariant::i32),
        );
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.listcolumns.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(ListColumn::from_dispatch(value.take_dispatch()?))
    }
    /// Iterates columns in Excel's `_NewEnum` order.
    pub fn iter(&self) -> Result<ListColumnsIter, ExcelComError> {
        Ok(ListColumnsIter {
            enumerator: enumerator(&self.inner, LIST_COLUMNS)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Fallible, single-pass, apartment-bound iterator over [`ListColumns`].
pub struct ListColumnsIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for ListColumnsIter {
    type Item = Result<ListColumn, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        next_typed(
            &mut self.enumerator,
            &mut self.next_index,
            &mut self.terminal,
            "ListColumns",
            ListColumn::from_dispatch,
        )
    }
}
impl FusedIterator for ListColumnsIter {}

/// Apartment-bound wrapper for one Excel table column.
pub struct ListColumn {
    inner: DispatchObject,
}
impl Debug for ListColumn {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ListColumn").field(&self.inner).finish()
    }
}
impl Clone for ListColumn {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ListColumn {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ListColumn",
            },
        }
    }
    /// Returns the one-based position in its table.
    pub fn index(&self) -> Result<usize, ExcelComError> {
        usize_get(&self.inner, "excel.listcolumn.index", "ListColumn.Index")
    }
    /// Returns the column's header text.
    pub fn name(&self) -> Result<String, ExcelComError> {
        string_get(&self.inner, "excel.listcolumn.name")
    }
    /// Changes the column's header text.
    pub fn set_name(&self, name: &str) -> Result<(), ExcelComError> {
        text_put(&self.inner, "excel.listcolumn.name", name)
    }
    /// Returns the full column Range, including its header and total cell when visible.
    pub fn range(&self) -> Result<Range, ExcelComError> {
        range_get(&self.inner, "excel.listcolumn.range")
    }
    /// Returns the data-body Range, or `None` for an empty table.
    pub fn data_body_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range_get(&self.inner, "excel.listcolumn.databodyrange")
    }
    /// Returns this column's totals cell Range when one exists.
    pub fn totals_row_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range_get(&self.inner, "excel.listcolumn.total")
    }
    /// Applies a calculated-column Formula through Excel's data-body Range.
    pub fn set_calculated_column_formula(&self, formula: &str) -> Result<(), ExcelComError> {
        self.data_body_range()?
            .ok_or(ExcelComError::Unsupported {
                detail: "calculated columns require a table data body",
            })?
            .set_table_calculated_column_formula(FormulaMember::Formula, formula)
    }
    /// Applies a dynamic-array-aware calculated-column Formula2 through Excel.
    pub fn set_calculated_column_formula2(&self, formula: &str) -> Result<(), ExcelComError> {
        self.data_body_range()?
            .ok_or(ExcelComError::Unsupported {
                detail: "calculated columns require a table data body",
            })?
            .set_table_calculated_column_formula(FormulaMember::Formula2, formula)
    }
    /// Returns Excel's totals calculation value, preserving unknown values.
    pub fn totals_calculation(&self) -> Result<TotalsCalculation, ExcelComError> {
        Ok(TotalsCalculation::from_raw(i32_get(
            &self.inner,
            "excel.listcolumn.totalscalculation",
            "ListColumn.TotalsCalculation",
        )?))
    }
    /// Sets Excel's totals calculation value.
    pub fn set_totals_calculation(
        &self,
        calculation: TotalsCalculation,
    ) -> Result<(), ExcelComError> {
        i32_put(
            &self.inner,
            "excel.listcolumn.totalscalculation",
            calculation.raw(),
        )
    }
    /// Deletes this table column and consumes the invalidated wrapper.
    pub fn delete(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.listcolumn.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

/// Apartment-bound typed collection of data rows in one Excel table.
pub struct ListRows {
    inner: DispatchObject,
}
impl Debug for ListRows {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ListRows").field(&self.inner).finish()
    }
}
impl Clone for ListRows {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ListRows {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ListRows",
            },
        }
    }
    /// Returns the number of table data rows.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, LIST_ROWS)
    }
    /// Returns the one-based data row at `index`.
    pub fn item(&self, index: usize) -> Result<ListRow, ExcelComError> {
        Ok(ListRow::from_dispatch(item_by_index(
            &self.inner,
            LIST_ROWS,
            index,
        )?))
    }
    /// Adds a row at an optional one-based position.
    ///
    /// `always_insert` maps to Excel's `AlwaysInsert`; it controls whether
    /// worksheet cells below the table are shifted instead of re-used.
    pub fn add(
        &self,
        position: Option<usize>,
        always_insert: Option<bool>,
    ) -> Result<ListRow, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_optional(
            position
                .map(|value| one_based(value, "ListRows.Add position"))
                .transpose()?
                .map(OwnedVariant::i32),
        );
        args.push_optional(always_insert.map(OwnedVariant::bool));
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.listrows.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(ListRow::from_dispatch(value.take_dispatch()?))
    }
    /// Iterates data rows in Excel's `_NewEnum` order.
    pub fn iter(&self) -> Result<ListRowsIter, ExcelComError> {
        Ok(ListRowsIter {
            enumerator: enumerator(&self.inner, LIST_ROWS)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Fallible, single-pass, apartment-bound iterator over [`ListRows`].
pub struct ListRowsIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for ListRowsIter {
    type Item = Result<ListRow, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        next_typed(
            &mut self.enumerator,
            &mut self.next_index,
            &mut self.terminal,
            "ListRows",
            ListRow::from_dispatch,
        )
    }
}
impl FusedIterator for ListRowsIter {}

/// Apartment-bound wrapper for one Excel table data row.
pub struct ListRow {
    inner: DispatchObject,
}
impl Debug for ListRow {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ListRow").field(&self.inner).finish()
    }
}
impl Clone for ListRow {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ListRow {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ListRow",
            },
        }
    }
    /// Returns the one-based data-row position in its table.
    pub fn index(&self) -> Result<usize, ExcelComError> {
        usize_get(&self.inner, "excel.listrow.index", "ListRow.Index")
    }
    /// Returns the table row Range, excluding worksheet cells outside the table.
    pub fn range(&self) -> Result<Range, ExcelComError> {
        range_get(&self.inner, "excel.listrow.range")
    }
    /// Deletes this table data row and consumes the invalidated wrapper.
    pub fn delete(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.listrow.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

impl Worksheet {
    /// Returns this worksheet's typed `ListObjects` collection.
    pub fn list_objects(&self) -> Result<ListObjects, ExcelComError> {
        object_get(
            self.dispatch_object(),
            "excel.worksheet.listobjects",
            ListObjects::from_dispatch,
        )
    }
}

pub(crate) fn object_get<T>(
    target: &DispatchObject,
    id: &'static str,
    make: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    Ok(make(value.take_dispatch()?))
}
pub(crate) fn range_get(target: &DispatchObject, id: &'static str) -> Result<Range, ExcelComError> {
    object_get(target, id, Range::from_dispatch)
}
pub(crate) fn optional_range_get(
    target: &DispatchObject,
    id: &'static str,
) -> Result<Option<Range>, ExcelComError> {
    let mut value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    value
        .take_optional_dispatch()
        .map(|value| value.map(Range::from_dispatch))
}
pub(crate) fn string_get(
    target: &DispatchObject,
    id: &'static str,
) -> Result<String, ExcelComError> {
    property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?.as_string()
}
pub(crate) fn bool_get(target: &DispatchObject, id: &'static str) -> Result<bool, ExcelComError> {
    property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?
        .as_bool()
        .ok_or(ExcelComError::Unsupported {
            detail: "expected Boolean Excel property result",
        })
}
pub(crate) fn i32_get(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<i32, ExcelComError> {
    property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?
        .as_i32()
        .ok_or(ExcelComError::Unsupported { detail })
}
pub(crate) fn usize_get(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<usize, ExcelComError> {
    usize::try_from(i32_get(target, id, detail)?).map_err(|_| ExcelComError::Unsupported { detail })
}
pub(crate) fn text_put(
    target: &DispatchObject,
    id: &'static str,
    text: &str,
) -> Result<(), ExcelComError> {
    let _ = property_put(
        &target.dispatch,
        member(MemberId::new(id), true),
        text_bstr(text)?,
    )?;
    Ok(())
}
pub(crate) fn bool_put(
    target: &DispatchObject,
    id: &'static str,
    value: bool,
) -> Result<(), ExcelComError> {
    let _ = property_put(
        &target.dispatch,
        member(MemberId::new(id), true),
        OwnedVariant::bool(value),
    )?;
    Ok(())
}
pub(crate) fn i32_put(
    target: &DispatchObject,
    id: &'static str,
    value: i32,
) -> Result<(), ExcelComError> {
    let _ = property_put(
        &target.dispatch,
        member(MemberId::new(id), true),
        OwnedVariant::i32(value),
    )?;
    Ok(())
}
pub(crate) fn one_based(value: usize, detail: &'static str) -> Result<i32, ExcelComError> {
    if value == 0 {
        return Err(ExcelComError::Unsupported {
            detail: "Excel table indexes and fields are one-based and nonzero",
        });
    }
    i32::try_from(value).map_err(|_| ExcelComError::Unsupported { detail })
}

fn next_typed<T>(
    enumerator: &mut EnumVariant,
    next_index: &mut usize,
    terminal: &mut bool,
    collection: &'static str,
    make: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Option<Result<T, ExcelComError>> {
    if *terminal {
        return None;
    }
    match enumerator.next() {
        Ok(Some(mut value)) => {
            let index = *next_index;
            *next_index += 1;
            Some(enumerated_dispatch(&mut value, collection, index).map(make))
        }
        Ok(None) => {
            *terminal = true;
            None
        }
        Err(error) => {
            *terminal = true;
            Some(Err(error))
        }
    }
}
