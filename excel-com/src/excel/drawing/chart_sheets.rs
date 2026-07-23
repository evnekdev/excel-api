//! Focused drawing implementation component.
#![allow(missing_docs)]
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
/// A workbook's collection of chart sheets.
pub struct Charts {
    inner: DispatchObject,
}
impl Debug for Charts {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Charts").field(&self.inner).finish()
    }
}
impl Clone for Charts {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Charts {
    pub(crate) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Charts", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, CHARTS)
    }
    pub fn item_by_index(&self, index: usize) -> Result<ChartSheet, ExcelComError> {
        collection_item(&self.inner, CHARTS, index, ChartSheet::from_dispatch)
    }
    pub fn item_by_name(&self, name: &str) -> Result<ChartSheet, ExcelComError> {
        collection_named(&self.inner, CHARTS, name, ChartSheet::from_dispatch)
    }
    pub fn iter(&self) -> Result<ChartsIter, ExcelComError> {
        Ok(ChartsIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(CHARTS.new_enum),
                CHARTS.name,
            )?,
            index: 0,
            done: false,
        })
    }
    pub fn add(
        &self,
        destination: &ChartSheetDestination<'_>,
    ) -> Result<ChartSheet, ExcelComError> {
        let mut args = PositionalArguments::new();
        match destination {
            ChartSheetDestination::Before(sheet) => {
                args.push_object(sheet.dispatch_object());
                args.push_optional(None);
                args.push_optional(None);
            }
            ChartSheetDestination::After(sheet) => {
                args.push_optional(None);
                args.push_object(sheet.dispatch_object());
                args.push_optional(None);
            }
            ChartSheetDestination::End => {
                args.push_optional(None);
                args.push_optional(None);
                args.push_optional(None);
            }
        };
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.charts.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(ChartSheet::from_dispatch(value.take_dispatch()?))
    }
}
pub struct ChartsIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for ChartsIter {
    type Item = Result<ChartSheet, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "Charts", index).map(ChartSheet::from_dispatch),
                )
            }
            Ok(None) => {
                self.done = true;
                None
            }
            Err(error) => {
                self.done = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for ChartsIter {}

/// A separate chart sheet in Excel's heterogeneous Sheets collection.
pub struct ChartSheet {
    inner: DispatchObject,
}
impl Debug for ChartSheet {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ChartSheet").field(&self.inner).finish()
    }
}
impl Clone for ChartSheet {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ChartSheet {
    pub(crate) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Chart", value),
        }
    }
    pub(crate) fn dispatch_object(&self) -> &DispatchObject {
        &self.inner
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.chart.name")
    }
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.chart.name", text_bstr(value)?)
    }
    pub fn index(&self) -> Result<usize, ExcelComError> {
        usize::try_from(get_i32(
            &self.inner,
            "excel.chart.index",
            "Chart.Index was not positive",
        )?)
        .ok()
        .filter(|value| *value > 0)
        .ok_or(ExcelComError::Unsupported {
            detail: "Chart.Index was not positive",
        })
    }
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.chart.activate", vec![])
    }
    pub fn chart(&self) -> Result<Chart, ExcelComError> {
        Ok(Chart {
            inner: self.inner.clone(),
        })
    }
    pub fn move_to(&self, destination: &SheetDestination<'_>) -> Result<(), ExcelComError> {
        sheet_copy_move(&self.inner, "excel.chart.move", destination)
    }
    pub fn copy(&self, destination: &SheetDestination<'_>) -> Result<ChartSheet, ExcelComError> {
        let args = sheet_destination_arguments(destination);
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.chart.copy"), false),
            args,
            false,
        )?;
        Ok(ChartSheet::from_dispatch(value.take_dispatch()?))
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.chart.delete", vec![])
    }
}
fn sheet_destination_arguments(destination: &SheetDestination<'_>) -> Vec<OwnedVariant> {
    let mut args = PositionalArguments::new();
    match destination {
        SheetDestination::Before(value) => {
            args.push_object(value.dispatch_object());
            args.push_optional(None);
        }
        SheetDestination::After(value) => {
            args.push_optional(None);
            args.push_object(value.dispatch_object());
        }
        SheetDestination::NewWorkbook => {
            args.push_optional(None);
            args.push_optional(None);
        }
    }
    args.into_inner()
}
fn sheet_copy_move(
    target: &DispatchObject,
    id: &'static str,
    destination: &SheetDestination<'_>,
) -> Result<(), ExcelComError> {
    call(target, id, sheet_destination_arguments(destination))
}
impl Workbook {
    /// Returns the workbook's chart-sheet collection.
    pub fn charts(&self) -> Result<Charts, ExcelComError> {
        get_dispatch(
            self.dispatch_object(),
            "excel.workbook.charts",
            Charts::from_dispatch,
        )
    }
}
