use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;

use crate::ExcelComError;
use crate::automation::{
    AutomationValue, ConversionPolicy, EnumVariant, OwnedVariant, PositionalArguments,
    decode_variant, enumerated_dispatch, invoke, property_get,
};
use crate::excel::collection::{
    CollectionDescriptor, count as collection_count, enumerator, item_by_index, item_by_name,
};
use crate::excel::{DispatchObject, Range, Worksheet};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

use super::helpers::{array_argument, text_argument};
use super::{ScenarioAddOptions, ScenarioSummaryOptions};

const DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Scenarios",
    count: MemberId::new("excel.scenarios.count"),
    item: MemberId::new("excel.scenarios.item"),
    new_enum: MemberId::new("excel.scenarios.newenum"),
};

/// A typed, apartment-bound Excel `Scenarios` collection.
pub struct Scenarios {
    inner: DispatchObject,
}
impl Debug for Scenarios {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Scenarios").field(&self.inner).finish()
    }
}
impl Clone for Scenarios {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Scenarios {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Scenarios",
            },
        }
    }
    /// Returns the number of scenarios.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, DESCRIPTOR)
    }
    /// Returns the one-based scenario at `index`.
    pub fn item_by_index(&self, index: usize) -> Result<Scenario, ExcelComError> {
        Ok(Scenario::from_dispatch(item_by_index(
            &self.inner,
            DESCRIPTOR,
            index,
        )?))
    }
    /// Returns a scenario by its Excel name.
    pub fn item_by_name(&self, name: &str) -> Result<Scenario, ExcelComError> {
        Ok(Scenario::from_dispatch(item_by_name(
            &self.inner,
            DESCRIPTOR,
            name,
        )?))
    }
    /// Iterates scenarios in Excel's `_NewEnum` order.
    pub fn iter(&self) -> Result<ScenariosIter, ExcelComError> {
        Ok(ScenariosIter {
            enumerator: enumerator(&self.inner, DESCRIPTOR)?,
            next_index: 0,
            terminal: false,
        })
    }
    /// Adds an Excel-owned scenario; values are passed as an Automation array without a detached scenario model.
    pub fn add(&self, options: &ScenarioAddOptions<'_>) -> Result<Scenario, ExcelComError> {
        validate_scenario_values(options)?;
        let mut args = PositionalArguments::new();
        args.push_result(text_argument(options.name))?;
        args.push_object(options.changing_cells.dispatch_object());
        args.push_optional(options.values.map(array_argument).transpose()?);
        args.push_optional(options.comment.map(text_argument).transpose()?);
        args.push_optional(options.locked.map(OwnedVariant::bool));
        args.push_optional(options.hidden.map(OwnedVariant::bool));
        let mut result = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.scenarios.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Scenario::from_dispatch(result.take_dispatch()?))
    }
    /// Creates a Scenario Summary worksheet and returns the worksheet reported by Excel.
    pub fn create_summary(
        &self,
        options: &ScenarioSummaryOptions<'_>,
    ) -> Result<Worksheet, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(options.report_type.raw()));
        args.push_optional_object(options.result_cells.map(Range::dispatch_object));
        let mut result = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.scenarios.createsummary"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Worksheet::from_dispatch(result.take_dispatch()?))
    }
}

/// A fallible, single-pass iterator over Excel scenarios.
pub struct ScenariosIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for ScenariosIter {
    type Item = Result<Scenario, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(
                    enumerated_dispatch(&mut value, "Scenarios", index)
                        .map(Scenario::from_dispatch),
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
impl FusedIterator for ScenariosIter {}

/// An apartment-bound Excel Scenario object.
pub struct Scenario {
    inner: DispatchObject,
}
impl Debug for Scenario {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Scenario").field(&self.inner).finish()
    }
}
impl Scenario {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Scenario",
            },
        }
    }
    /// Returns Excel's current scenario name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.scenario.name"), false),
            vec![],
        )?
        .as_string()
    }
    /// Returns Excel's current scenario comment.
    pub fn comment(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.scenario.comment"), false),
            vec![],
        )?
        .as_string()
    }
    /// Returns the changing cells as an apartment-bound Range.
    pub fn changing_cells(&self) -> Result<Range, ExcelComError> {
        let mut value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.scenario.changingcells"), false),
            vec![],
        )?;
        Ok(Range::from_dispatch(value.take_dispatch()?))
    }
    /// Returns the scenario values as Excel Automation data.
    pub fn values(&self) -> Result<AutomationValue, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.scenario.values"), false),
            vec![],
        )?;
        decode_variant(&value, ConversionPolicy::default())
    }
    /// Shows this Scenario's values in its worksheet.
    pub fn show(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.scenario.show"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Changes selected scenario properties; unspecified positions remain `Missing`.
    pub fn change(
        &self,
        changing_cells: Option<&Range>,
        values: Option<&crate::AutomationArray>,
    ) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_optional_object(changing_cells.map(Range::dispatch_object));
        args.push_optional(values.map(array_argument).transpose()?);
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.scenario.changescenario"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }
    /// Deletes this Excel Scenario and consumes its wrapper.
    pub fn delete(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.scenario.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

impl Worksheet {
    /// Returns this worksheet's apartment-bound Excel Scenarios collection.
    pub fn scenarios(&self) -> Result<Scenarios, ExcelComError> {
        let mut value = property_get(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.worksheet.scenarios"), false),
            vec![],
        )?;
        Ok(Scenarios::from_dispatch(value.take_dispatch()?))
    }
}

fn validate_scenario_values(options: &ScenarioAddOptions<'_>) -> Result<(), ExcelComError> {
    if options.name.contains('\0') || options.comment.is_some_and(|value| value.contains('\0')) {
        return Err(ExcelComError::Unsupported {
            detail: "Scenario text cannot contain embedded NUL",
        });
    }
    if let Some(values) = options.values {
        if values.rows() != 1 && values.columns() != 1 {
            return Err(ExcelComError::Unsupported {
                detail: "Scenario values must be one-dimensional",
            });
        }
    }
    Ok(())
}
