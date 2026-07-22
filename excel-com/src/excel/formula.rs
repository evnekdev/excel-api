//! Formula, array-formula, spill, and dependency Range operations.

use crate::ExcelComError;
use crate::automation::{
    AutomationArray, AutomationValue, ConversionPolicy, OwnedVariant, decode_variant,
    encode_variant, invoke, property_get, property_put, validate_range_shape,
};
use crate::excel::Range;
use crate::excel::formatting::{MixedValue, mixed_bool, property_mixed_get};
use crate::excel::text::text_bstr;
use crate::object_model::{MemberId, member};
use windows_sys::Win32::System::Variant::{
    VT_ARRAY, VT_BSTR, VT_DISPATCH, VT_EMPTY, VT_NULL, VT_VARIANT,
};

/// Excel formula text returned by a Range formula property.
///
/// Formula getters preserve a scalar formula string separately from a
/// rectangular `AutomationArray`. `Mixed` represents Excel `VT_NULL`, while
/// `Empty` represents `VT_EMPTY`. An empty formula string is deliberately
/// represented as `Text(String::new())`, not as `Empty`.
#[derive(Clone, Debug, PartialEq)]
pub enum FormulaValue {
    /// A scalar formula string exactly as returned by Excel.
    Text(String),
    /// A zero-based row-major rectangular array returned by Excel.
    ///
    /// Excel owns formula grammar and can return constants in an array result;
    /// their semantic `AutomationValue` forms are retained rather than parsed.
    Array(AutomationArray),
    /// The source cells have no common formula value (`VT_NULL`).
    Mixed,
    /// Excel returned `VT_EMPTY`.
    Empty,
}

#[derive(Clone, Copy)]
pub(crate) enum FormulaMember {
    Formula,
    Formula2,
    FormulaR1C1,
    Formula2R1C1,
    FormulaLocal,
    FormulaR1C1Local,
}

impl FormulaMember {
    const fn id(self) -> &'static str {
        match self {
            Self::Formula => "excel.range.formula",
            Self::Formula2 => "excel.range.formula2",
            Self::FormulaR1C1 => "excel.range.formular1c1",
            Self::Formula2R1C1 => "excel.range.formula2r1c1",
            Self::FormulaLocal => "excel.range.formulalocal",
            Self::FormulaR1C1Local => "excel.range.formular1c1local",
        }
    }
}

impl Range {
    pub(crate) fn set_table_calculated_column_formula(
        &self,
        formula_member: FormulaMember,
        formula: &str,
    ) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.dispatch_object().dispatch,
            member(MemberId::new(formula_member.id()), true),
            text_bstr(formula)?,
        )?;
        Ok(())
    }

    /// Returns invariant A1-style `Range.Formula` text or a formula array.
    ///
    /// `Formula` retains Excel's legacy implicit-intersection behavior. Use
    /// [`Self::formula2`] for dynamic-array-aware formula semantics.
    pub fn formula(&self) -> Result<FormulaValue, ExcelComError> {
        self.formula_get(FormulaMember::Formula)
    }

    /// Sets an invariant A1-style `Range.Formula` on a single cell.
    ///
    /// This crate deliberately rejects scalar assignment to a multi-cell Range
    /// before COM so callers cannot accidentally rely on Excel's fill policy.
    /// Use [`Self::set_formula_array_values`] for a same-shape rectangle.
    pub fn set_formula(&self, formula: &str) -> Result<(), ExcelComError> {
        self.set_formula_text(FormulaMember::Formula, formula)
    }

    /// Sets invariant A1-style Formula values from an exact-shape rectangle.
    ///
    /// Every element must be formula text and the array dimensions must exactly
    /// equal the receiver Range. This is distinct from legacy
    /// [`Self::set_formula_array`].
    pub fn set_formula_array_values(
        &self,
        formulas: &AutomationArray,
    ) -> Result<(), ExcelComError> {
        self.set_formula_values(FormulaMember::Formula, formulas)
    }

    /// Returns dynamic-array-aware A1-style `Range.Formula2` text or an array.
    ///
    /// In dynamic-array-enabled Excel, Formula2 formulas are evaluated as
    /// arrays and may spill. Version availability and formula interpretation
    /// remain controlled by Excel.
    pub fn formula2(&self) -> Result<FormulaValue, ExcelComError> {
        self.formula_get(FormulaMember::Formula2)
    }

    /// Sets a dynamic-array-aware A1-style formula on a single cell.
    pub fn set_formula2(&self, formula: &str) -> Result<(), ExcelComError> {
        self.set_formula_text(FormulaMember::Formula2, formula)
    }

    /// Sets dynamic-array-aware A1-style formulas from an exact-shape rectangle.
    pub fn set_formula2_array(&self, formulas: &AutomationArray) -> Result<(), ExcelComError> {
        self.set_formula_values(FormulaMember::Formula2, formulas)
    }

    /// Returns invariant R1C1-style `Range.FormulaR1C1` text or an array.
    ///
    /// Relative references are interpreted by Excel; this crate never
    /// translates R1C1 text to A1.
    pub fn formula_r1c1(&self) -> Result<FormulaValue, ExcelComError> {
        self.formula_get(FormulaMember::FormulaR1C1)
    }

    /// Sets an invariant R1C1-style formula on a single cell.
    pub fn set_formula_r1c1(&self, formula: &str) -> Result<(), ExcelComError> {
        self.set_formula_text(FormulaMember::FormulaR1C1, formula)
    }

    /// Sets invariant R1C1-style formulas from an exact-shape rectangle.
    pub fn set_formula_r1c1_array(&self, formulas: &AutomationArray) -> Result<(), ExcelComError> {
        self.set_formula_values(FormulaMember::FormulaR1C1, formulas)
    }

    /// Returns dynamic-array-aware R1C1 `Range.Formula2R1C1` text or an array.
    ///
    /// The installed Excel type library declares this version-sensitive member.
    /// Its array evaluation and spill semantics remain Excel-defined.
    pub fn formula2_r1c1(&self) -> Result<FormulaValue, ExcelComError> {
        self.formula_get(FormulaMember::Formula2R1C1)
    }

    /// Sets a dynamic-array-aware R1C1 formula on a single cell.
    pub fn set_formula2_r1c1(&self, formula: &str) -> Result<(), ExcelComError> {
        self.set_formula_text(FormulaMember::Formula2R1C1, formula)
    }

    /// Sets dynamic-array-aware R1C1 formulas from an exact-shape rectangle.
    pub fn set_formula2_r1c1_array(&self, formulas: &AutomationArray) -> Result<(), ExcelComError> {
        self.set_formula_values(FormulaMember::Formula2R1C1, formulas)
    }

    /// Returns locale-dependent A1 `Range.FormulaLocal` text or an array.
    ///
    /// Excel supplies function names and separators in the installed locale.
    /// The crate preserves that text and never translates it.
    pub fn formula_local(&self) -> Result<FormulaValue, ExcelComError> {
        self.formula_get(FormulaMember::FormulaLocal)
    }

    /// Sets locale-dependent A1 formula text on a single cell.
    ///
    /// Callers are responsible for supplying syntax accepted by the active
    /// Excel locale.
    pub fn set_formula_local(&self, formula: &str) -> Result<(), ExcelComError> {
        self.set_formula_text(FormulaMember::FormulaLocal, formula)
    }

    /// Returns locale-dependent R1C1 `Range.FormulaR1C1Local` text or an array.
    pub fn formula_r1c1_local(&self) -> Result<FormulaValue, ExcelComError> {
        self.formula_get(FormulaMember::FormulaR1C1Local)
    }

    /// Sets locale-dependent R1C1 formula text on a single cell.
    pub fn set_formula_r1c1_local(&self, formula: &str) -> Result<(), ExcelComError> {
        self.set_formula_text(FormulaMember::FormulaR1C1Local, formula)
    }

    /// Returns formula presence while retaining mixed and empty states.
    ///
    /// `Uniform(true)` means all selected cells have formulas,
    /// `Uniform(false)` means none do, and `Mixed` means Excel reported a
    /// heterogeneous selection.
    pub fn has_formula(&self) -> Result<MixedValue<bool>, ExcelComError> {
        self.mixed_bool_property("excel.range.hasformula")
    }

    /// Returns legacy array-formula presence while retaining mixed and empty states.
    pub fn has_array(&self) -> Result<MixedValue<bool>, ExcelComError> {
        self.mixed_bool_property("excel.range.hasarray")
    }

    /// Returns the entire legacy array-formula area containing this Range.
    ///
    /// Excel reports its structured invocation error when the Range is not part
    /// of a legacy array formula.
    pub fn current_array(&self) -> Result<Range, ExcelComError> {
        self.required_range_property("excel.range.currentarray")
    }

    /// Returns `Range.FormulaArray`, the legacy array-formula member.
    ///
    /// This is intentionally separate from rectangular ordinary formula arrays.
    /// Excel reports `Mixed` when no single legacy-array formula is available.
    pub fn formula_array(&self) -> Result<FormulaValue, ExcelComError> {
        self.formula_property_get("excel.range.formulaarray")
    }

    /// Assigns one legacy array formula across the receiver Range.
    ///
    /// Excel, rather than this crate, enforces legacy-array rules (including
    /// the documented 255-character FormulaArray limit). This method does not
    /// emulate Ctrl+Shift+Enter and preserves any Excel rejection, including a
    /// partial edit inside an existing array formula.
    pub fn set_formula_array(&self, formula: &str) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.formulaarray"), true),
            text_bstr(formula)?,
        )?;
        Ok(())
    }

    /// Returns dynamic-spill membership while retaining mixed and empty states.
    pub fn has_spill(&self) -> Result<MixedValue<bool>, ExcelComError> {
        self.mixed_bool_property("excel.range.hasspill")
    }

    /// Returns the Range currently produced by a dynamic-array spill origin.
    ///
    /// Excel returns an error when no spill Range exists; this method does not
    /// infer a spill rectangle from addresses.
    pub fn spilling_to_range(&self) -> Result<Range, ExcelComError> {
        self.required_range_property("excel.range.spillingtorange")
    }

    /// Returns the dynamic-array spill origin for a spill member.
    ///
    /// Excel returns an error if the receiver is not in a spill range.
    pub fn spill_parent(&self) -> Result<Range, ExcelComError> {
        self.required_range_property("excel.range.spillparent")
    }

    /// Calculates only this Range according to Excel's calculation engine.
    ///
    /// Completion and dependency behavior are Excel-defined; callers can read
    /// [`crate::Application::calculation_state`] after this call when needed.
    pub fn calculate(&self) -> Result<(), ExcelComError> {
        self.range_method("excel.range.calculate")
    }

    /// Marks this Range dirty for Excel recalculation.
    ///
    /// This does not itself guarantee calculation. It is version-sensitive and
    /// preserves Excel's structured error if the host rejects the operation.
    pub fn mark_dirty(&self) -> Result<(), ExcelComError> {
        self.range_method("excel.range.dirty")
    }

    /// Returns Excel's direct precedent Range, including any Excel-defined areas.
    pub fn direct_precedents(&self) -> Result<Range, ExcelComError> {
        self.required_range_property("excel.range.directprecedents")
    }

    /// Returns Excel's direct dependent Range, including any Excel-defined areas.
    pub fn direct_dependents(&self) -> Result<Range, ExcelComError> {
        self.required_range_property("excel.range.directdependents")
    }

    /// Returns Excel's transitive precedent Range.
    ///
    /// Cross-sheet and cross-workbook behavior remains subject to Excel's
    /// auditing limitations and is not expanded into a Rust dependency graph.
    pub fn precedents(&self) -> Result<Range, ExcelComError> {
        self.required_range_property("excel.range.precedents")
    }

    /// Returns Excel's transitive dependent Range.
    ///
    /// Cross-sheet and cross-workbook behavior remains subject to Excel's
    /// auditing limitations and is not expanded into a Rust dependency graph.
    pub fn dependents(&self) -> Result<Range, ExcelComError> {
        self.required_range_property("excel.range.dependents")
    }

    fn formula_get(&self, member: FormulaMember) -> Result<FormulaValue, ExcelComError> {
        self.formula_property_get(member.id())
    }

    fn formula_property_get(&self, id: &'static str) -> Result<FormulaValue, ExcelComError> {
        let value = property_get(
            &self.dispatch_object().dispatch,
            member(MemberId::new(id), false),
            vec![],
        )?;
        decode_formula_value(&value)
    }

    fn set_formula_text(
        &self,
        formula_member: FormulaMember,
        formula: &str,
    ) -> Result<(), ExcelComError> {
        let (rows, columns) = self.dimensions()?;
        validate_range_shape(&AutomationValue::Text(formula.to_owned()), rows, columns)?;
        let _ = property_put(
            &self.dispatch_object().dispatch,
            member(MemberId::new(formula_member.id()), true),
            text_bstr(formula)?,
        )?;
        Ok(())
    }

    fn set_formula_values(
        &self,
        formula_member: FormulaMember,
        formulas: &AutomationArray,
    ) -> Result<(), ExcelComError> {
        let (rows, columns) = self.dimensions()?;
        validate_range_shape(&AutomationValue::Array(formulas.clone()), rows, columns)?;
        for value in formulas.values() {
            if !matches!(value, AutomationValue::Text(_)) {
                return Err(ExcelComError::Unsupported {
                    detail: "formula array values must contain only formula text",
                });
            }
        }
        let encoded = encode_variant(
            &AutomationValue::Array(formulas.clone()),
            ConversionPolicy::default(),
        )?;
        let _ = property_put(
            &self.dispatch_object().dispatch,
            member(MemberId::new(formula_member.id()), true),
            encoded,
        )?;
        Ok(())
    }

    fn mixed_bool_property(&self, id: &'static str) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), id, mixed_bool)
    }

    fn required_range_property(&self, id: &'static str) -> Result<Range, ExcelComError> {
        let mut value = property_get(
            &self.dispatch_object().dispatch,
            member(MemberId::new(id), false),
            vec![],
        )?;
        Ok(Range::from_dispatch(value.take_dispatch()?))
    }

    fn range_method(&self, id: &'static str) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new(id), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

fn decode_formula_value(value: &OwnedVariant) -> Result<FormulaValue, ExcelComError> {
    if value.vt() == VT_ARRAY | VT_VARIANT {
        return match decode_variant(value, ConversionPolicy::default())? {
            AutomationValue::Array(array) => Ok(FormulaValue::Array(array)),
            _ => Err(ExcelComError::Unsupported {
                detail: "formula array result did not decode as an Automation array",
            }),
        };
    }
    match value.vt() {
        VT_BSTR => value.as_string().map(FormulaValue::Text),
        VT_NULL => Ok(FormulaValue::Mixed),
        VT_EMPTY => Ok(FormulaValue::Empty),
        VT_DISPATCH => Err(ExcelComError::Unsupported {
            detail: "formula property returned an object",
        }),
        _ => Err(ExcelComError::Unsupported {
            detail: "formula property returned an unsupported scalar value",
        }),
    }
}

#[cfg(test)]
mod tests {
    use super::{FormulaMember, FormulaValue, decode_formula_value};
    use crate::automation::{AutomationArray, AutomationValue, OwnedVariant};

    #[test]
    fn formula_result_decoding_preserves_text_array_and_mixed_states() {
        assert_eq!(
            decode_formula_value(&OwnedVariant::bstr("=A1").expect("BSTR")),
            Ok(FormulaValue::Text("=A1".to_owned()))
        );
        assert_eq!(
            decode_formula_value(&OwnedVariant::null()),
            Ok(FormulaValue::Mixed)
        );
        assert_eq!(
            decode_formula_value(&OwnedVariant::empty()),
            Ok(FormulaValue::Empty)
        );
        assert!(decode_formula_value(&OwnedVariant::i32(1)).is_err());
        let array =
            AutomationArray::row(vec![AutomationValue::Text("=A1".to_owned())]).expect("array");
        let variant = crate::automation::encode_variant(
            &AutomationValue::Array(array.clone()),
            crate::ConversionPolicy::default(),
        )
        .expect("variant");
        assert_eq!(
            decode_formula_value(&variant),
            Ok(FormulaValue::Array(array))
        );
    }

    #[test]
    fn formula_member_ids_are_exact() {
        assert_eq!(FormulaMember::Formula.id(), "excel.range.formula");
        assert_eq!(FormulaMember::Formula2.id(), "excel.range.formula2");
        assert_eq!(FormulaMember::FormulaR1C1.id(), "excel.range.formular1c1");
        assert_eq!(FormulaMember::Formula2R1C1.id(), "excel.range.formula2r1c1");
        assert_eq!(FormulaMember::FormulaLocal.id(), "excel.range.formulalocal");
        assert_eq!(
            FormulaMember::FormulaR1C1Local.id(),
            "excel.range.formular1c1local"
        );
    }
}
