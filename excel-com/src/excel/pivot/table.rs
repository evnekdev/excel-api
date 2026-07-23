//! PivotTable report operations and guarded layout application.

use super::helpers::{boolean, object, optional_range, put, range, text};
use super::{PivotCache, PivotDataField, PivotFields, PivotLayoutOptions};
use crate::ExcelComError;
use crate::automation::{OwnedVariant, invoke};
use crate::excel::{DispatchObject, Range};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use std::fmt::{Debug, Formatter};

/// Apartment-bound PivotTable report. A PivotCache owns source data; this
/// object owns report layout, field placement, formatting, and filters.
pub struct PivotTable {
    inner: DispatchObject,
}
impl Debug for PivotTable {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotTable").field(&self.inner).finish()
    }
}
impl Clone for PivotTable {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotTable {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotTable",
            },
        }
    }
    /// Returns the report name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivottable.name")
    }
    /// Renames the report after rejecting embedded NUL.
    pub fn set_name(&self, name: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivottable.name",
            OwnedVariant::bstr(name)?,
        )
    }
    /// Returns the complete PivotTable range excluding page fields.
    pub fn table_range1(&self) -> Result<Range, ExcelComError> {
        range(&self.inner, "excel.pivottable.tablerange1")
    }
    /// Returns the complete PivotTable range including page fields.
    pub fn table_range2(&self) -> Result<Range, ExcelComError> {
        range(&self.inner, "excel.pivottable.tablerange2")
    }
    /// Returns the data body range or `None` when Excel has no data body.
    pub fn data_body_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range(&self.inner, "excel.pivottable.databodyrange")
    }
    /// Returns the row-label range when it exists.
    pub fn row_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range(&self.inner, "excel.pivottable.rowrange")
    }
    /// Returns the column-label range when it exists.
    pub fn column_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range(&self.inner, "excel.pivottable.columnrange")
    }
    /// Returns the page-field range when it exists.
    pub fn page_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range(&self.inner, "excel.pivottable.pagerange")
    }
    /// Returns the cache that owns this report's source data.
    pub fn pivot_cache(&self) -> Result<PivotCache, ExcelComError> {
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivottable.pivotcache"), false),
            vec![],
            false,
        )?;
        Ok(PivotCache::from_dispatch(value.take_dispatch()?))
    }
    /// Returns all report fields.
    pub fn pivot_fields(&self) -> Result<PivotFields, ExcelComError> {
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivottable.pivotfields"), false),
            vec![],
            false,
        )?;
        Ok(PivotFields::from_dispatch(value.take_dispatch()?))
    }
    /// Returns current row fields.
    pub fn row_fields(&self) -> Result<PivotFields, ExcelComError> {
        object(
            &self.inner,
            "excel.pivottable.rowfields",
            PivotFields::from_dispatch,
        )
    }
    /// Returns current column fields.
    pub fn column_fields(&self) -> Result<PivotFields, ExcelComError> {
        object(
            &self.inner,
            "excel.pivottable.columnfields",
            PivotFields::from_dispatch,
        )
    }
    /// Returns current page fields.
    pub fn page_fields(&self) -> Result<PivotFields, ExcelComError> {
        object(
            &self.inner,
            "excel.pivottable.pagefields",
            PivotFields::from_dispatch,
        )
    }
    /// Returns current data fields.
    pub fn data_fields(&self) -> Result<PivotFields, ExcelComError> {
        object(
            &self.inner,
            "excel.pivottable.datafields",
            PivotFields::from_dispatch,
        )
    }
    /// Refreshes the report and preserves Excel's Boolean result.
    pub fn refresh_table(&self) -> Result<bool, ExcelComError> {
        let value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivottable.refreshtable"), false),
            vec![],
            false,
        )?;
        value.as_bool().ok_or(ExcelComError::Unsupported {
            detail: "PivotTable.RefreshTable did not return VT_BOOL",
        })
    }
    /// Asks Excel to update the report after property changes.
    pub fn update(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivottable.update"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Returns whether Excel defers automatic report updates.
    pub fn manual_update(&self) -> Result<bool, ExcelComError> {
        boolean(&self.inner, "excel.pivottable.manualupdate")
    }
    /// Enables or disables deferred report updates.
    pub fn set_manual_update(&self, enabled: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivottable.manualupdate",
            OwnedVariant::bool(enabled),
        )
    }
    /// Returns whether row stripes are shown by the active TableStyle2.
    pub fn show_table_style_row_stripes(&self) -> Result<bool, ExcelComError> {
        boolean(&self.inner, "excel.pivottable.showtablestylerowstripes")
    }
    /// Enables or disables TableStyle2 row stripes.
    pub fn set_show_table_style_row_stripes(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivottable.showtablestylerowstripes",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the current Excel table style name.
    pub fn table_style2(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivottable.tablestyle2")
    }
    /// Sets the Excel table style name after rejecting embedded NUL.
    pub fn set_table_style2(&self, name: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivottable.tablestyle2",
            OwnedVariant::bstr(name)?,
        )
    }
    /// Clears the report and consumes the wrapper because Excel may invalidate it.
    pub fn clear_table(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivottable.cleartable"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Applies placements and data fields while restoring `ManualUpdate` explicitly and on drop.
    pub fn apply_layout(&self, options: &PivotLayoutOptions<'_>) -> Result<(), ExcelComError> {
        let mut guard = PivotTableManualUpdateGuard::new(self)?;
        let result = self.apply_layout_inner(options);
        let restore = guard.restore();
        match (result, restore) {
            (Err(operation), Err(restoration)) => Err(ExcelComError::StateRestoration {
                operation: Box::new(operation),
                restoration: Box::new(restoration),
            }),
            (Err(error), Ok(())) => Err(error),
            (Ok(()), Err(error)) => Err(error),
            (Ok(()), Ok(())) => self.update(),
        }
    }
    fn apply_layout_inner(&self, options: &PivotLayoutOptions<'_>) -> Result<(), ExcelComError> {
        for placement in &options.fields {
            let field = self.pivot_fields()?.item_by_name(placement.field_name)?;
            field.set_orientation(placement.orientation)?;
            if let Some(position) = placement.position {
                field.set_position(position)?;
            }
        }
        for data in &options.data_fields {
            self.add_data_field(data)?;
        }
        Ok(())
    }
    fn add_data_field(&self, data: &PivotDataField<'_>) -> Result<(), ExcelComError> {
        let field = self.pivot_fields()?.item_by_name(data.field_name)?;
        let mut args = crate::automation::PositionalArguments::new();
        args.push_object(field.dispatch_object());
        match data.caption {
            Some(caption) => args.push_required(OwnedVariant::bstr(caption)?),
            None => args.push_optional(None),
        };
        args.push_required(OwnedVariant::i32(data.function.raw()));
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivottable.adddatafield"), false),
            args.into_inner(),
            false,
        )?;
        let data_field = super::PivotField::from_dispatch(value.take_dispatch()?);
        if let Some(format) = data.number_format {
            data_field.set_number_format(format)?;
        }
        Ok(())
    }
    /// Changes the report to an existing compatible PivotCache.
    pub fn change_pivot_cache(&self, cache: &PivotCache) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivottable.changepivotcache"), false),
            vec![OwnedVariant::dispatch_borrowed(
                &cache.dispatch_object().dispatch,
            )],
            false,
        )?;
        Ok(())
    }
    /// Changes source data using Excel's writable `PivotTable.SourceData` property.
    pub fn change_source_data(&self, source: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivottable.sourcedata",
            OwnedVariant::bstr(&super::helpers::source_reference(source)?)?,
        )
    }
}

struct PivotTableManualUpdateGuard<'a> {
    table: &'a PivotTable,
    previous: bool,
    active: bool,
}
impl<'a> PivotTableManualUpdateGuard<'a> {
    fn new(table: &'a PivotTable) -> Result<Self, ExcelComError> {
        let previous = table.manual_update()?;
        table.set_manual_update(true)?;
        Ok(Self {
            table,
            previous,
            active: true,
        })
    }
    fn restore(&mut self) -> Result<(), ExcelComError> {
        if self.active {
            self.table.set_manual_update(self.previous)?;
            self.active = false;
        }
        Ok(())
    }
}
impl Drop for PivotTableManualUpdateGuard<'_> {
    fn drop(&mut self) {
        if self.active {
            let _ = self.table.set_manual_update(self.previous);
            self.active = false;
        }
    }
}
