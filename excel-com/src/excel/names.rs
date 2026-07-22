use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;

use crate::ExcelComError;
use crate::automation::{
    EnumVariant, OwnedVariant, PositionalArguments, enumerated_dispatch, invoke, property_get,
};
use crate::excel::collection::{
    CollectionDescriptor, count as collection_count, enumerator, item_by_index,
};
use crate::excel::text::text_bstr;
use crate::excel::{DispatchObject, Name, Range, ReferenceStyle};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

const DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Names",
    count: MemberId::new("excel.names.count"),
    item: MemberId::new("excel.names.item"),
    new_enum: MemberId::new("excel.names.newenum"),
};

/// Describes the target supplied to [`Names::add`].
///
/// Range targets are converted to Excel's own external A1 address before
/// calling `Names.Add`; a Range dispatch object is not passed as `RefersTo`.
/// A1, R1C1, and formula strings are validated by Excel after embedded NUL is
/// rejected locally.
#[derive(Debug)]
pub enum NameRefersTo<'a> {
    /// An existing Range, encoded as Excel's qualified external A1 expression.
    Range(&'a Range),
    /// An A1-style reference expression, sent through `RefersTo`.
    A1(&'a str),
    /// An R1C1-style reference expression, sent through `RefersToR1C1`.
    R1C1(&'a str),
    /// A formula-oriented expression, sent unchanged through `RefersTo`.
    Formula(&'a str),
}

/// Narrow input for [`Names::add`].
///
/// Workbook or worksheet scope is determined by the collection used to create
/// the Name. This first slice intentionally leaves Excel's macro, category,
/// localization, and shortcut arguments as `Missing`.
#[derive(Debug)]
pub struct NameAddOptions<'a> {
    /// The simple name requested from Excel.
    pub name: &'a str,
    /// The target expression or Range.
    pub refers_to: NameRefersTo<'a>,
    /// Whether Excel should show the Name, or `None` to send `Missing`.
    pub visible: Option<bool>,
}

/// Experimental typed wrapper for an Excel Names collection.
pub struct Names {
    inner: DispatchObject,
}

impl Debug for Names {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Names").field(&self.inner).finish()
    }
}

impl Clone for Names {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl Names {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Names",
            },
        }
    }

    /// Returns the collection's current number of Names.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, DESCRIPTOR)
    }

    /// Returns the one-based Name at `index`.
    pub fn item_by_index(&self, index: usize) -> Result<Name, ExcelComError> {
        Ok(Name::from_dispatch(item_by_index(
            &self.inner,
            DESCRIPTOR,
            index,
        )?))
    }

    /// Returns the Name selected by its Excel-visible string key.
    ///
    /// Workbook and worksheet-local collections may expose different
    /// qualification strings. Excel supplies lookup and collision semantics.
    pub fn item_by_name(&self, name: &str) -> Result<Name, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.names.item"), false),
            vec![text_bstr(name)?],
        )?;
        Ok(Name::from_dispatch(result.take_dispatch()?))
    }

    /// Adds a Name using the narrow, position-preserving Prompt 11 input model.
    ///
    /// ```no_run
    /// # fn example(names: &excel_com::Names, range: &excel_com::Range) -> Result<(), excel_com::ExcelComError> {
    /// use excel_com::{NameAddOptions, NameRefersTo};
    /// let name = names.add(&NameAddOptions {
    ///     name: "InputRange",
    ///     refers_to: NameRefersTo::Range(range),
    ///     visible: Some(true),
    /// })?;
    /// # drop(name);
    /// # Ok(())
    /// # }
    /// ```
    pub fn add(&self, options: &NameAddOptions<'_>) -> Result<Name, ExcelComError> {
        if options.name.is_empty() {
            return Err(ExcelComError::Unsupported {
                detail: "Names.Add requires a nonempty name",
            });
        }
        let (refers_to, refers_to_r1c1) = match &options.refers_to {
            NameRefersTo::Range(range) => (
                Some(ensure_formula(&range.external_address(ReferenceStyle::A1)?)),
                None,
            ),
            NameRefersTo::A1(value) => (Some(ensure_formula(value)), None),
            NameRefersTo::R1C1(value) => (None, Some(ensure_formula(value))),
            NameRefersTo::Formula(value) => (Some((*value).to_owned()), None),
        };
        let arguments = name_add_arguments(
            options.name,
            refers_to.as_deref(),
            refers_to_r1c1.as_deref(),
            options.visible,
        )?;
        let mut result = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.names.add"), false),
            arguments,
            false,
        )?;
        Ok(Name::from_dispatch(result.take_dispatch()?))
    }

    /// Iterates Names through Excel's `_NewEnum` in Excel-defined order.
    ///
    /// The iterator is apartment-bound, single-pass, and fallible. A terminal
    /// enumeration error fuses it; dropping it early releases the enumerator.
    pub fn iter(&self) -> Result<NamesIter, ExcelComError> {
        Ok(NamesIter {
            enumerator: enumerator(&self.inner, DESCRIPTOR)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Typed, single-pass iterator over Excel Names.
pub struct NamesIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}

impl Iterator for NamesIter {
    type Item = Result<Name, ExcelComError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(enumerated_dispatch(&mut value, "Names", index).map(Name::from_dispatch))
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

impl FusedIterator for NamesIter {}

fn ensure_formula(value: &str) -> String {
    if value.starts_with('=') {
        value.to_owned()
    } else {
        format!("={value}")
    }
}

fn name_add_arguments(
    name: &str,
    refers_to: Option<&str>,
    refers_to_r1c1: Option<&str>,
    visible: Option<bool>,
) -> Result<Vec<OwnedVariant>, ExcelComError> {
    let mut arguments = PositionalArguments::new();
    arguments.push_result(text_bstr(name))?;
    push_optional_text(&mut arguments, refers_to)?;
    arguments.push_optional(visible.map(OwnedVariant::bool));
    // MacroType, ShortcutKey, Category, NameLocal, RefersToLocal, CategoryLocal.
    for _ in 0..6 {
        arguments.push_optional(None);
    }
    // RefersToR1C1 and RefersToR1C1Local.
    push_optional_text(&mut arguments, refers_to_r1c1)?;
    arguments.push_optional(None);
    Ok(arguments.into_inner())
}

fn push_optional_text(
    arguments: &mut PositionalArguments,
    value: Option<&str>,
) -> Result<(), ExcelComError> {
    match value {
        Some(value) => arguments.push_result(text_bstr(value)),
        None => {
            arguments.push_optional(None);
            Ok(())
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn name_add_preserves_all_eleven_positions_for_a1() {
        let values =
            name_add_arguments("Input", Some("=Sheet1!$A$1"), None, Some(true)).expect("arguments");
        assert_eq!(values.len(), 11);
        assert_eq!(values[0].as_string().expect("name"), "Input");
        assert_eq!(values[1].as_string().expect("formula"), "=Sheet1!$A$1");
        assert_eq!(values[2].as_bool(), Some(true));
        assert_eq!(values[9].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[10].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }

    #[test]
    fn name_add_uses_only_r1c1_slot_for_r1c1_input() {
        let values = name_add_arguments("Input", None, Some("=R1C1"), None).expect("arguments");
        assert_eq!(values[1].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[9].as_string().expect("r1c1"), "=R1C1");
        assert_eq!(values[10].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }

    #[test]
    fn text_and_formula_encoding_reject_nul_without_com() {
        assert!(name_add_arguments("bad\0name", Some("=A1"), None, None).is_err());
        assert!(name_add_arguments("name", Some("=A\0"), None, None).is_err());
        assert_eq!(ensure_formula("A1"), "=A1");
        assert_eq!(ensure_formula("=SUM(A1:A2)"), "=SUM(A1:A2)");
    }
}
