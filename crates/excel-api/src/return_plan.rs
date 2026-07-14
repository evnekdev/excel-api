//! Pure, deterministic planning for future Excel ABI return storage.
//!
//! This module deliberately creates no raw pointers, backing buffers,
//! ownership bits, or FFI values. A successful [`ReturnPlan`] owns the exact
//! logical payload and enough validated metadata for Prompt 05 to materialize
//! it without revisiting return policy.

use core::mem::size_of;

use crate::{ExcelArray, ExcelError, ExcelString, ExcelValue, ReturnError};

const XLOPER12_BYTES: usize = size_of::<excel_api_sys::XLOPER12>();
const XCHAR_BYTES: usize = size_of::<excel_api_sys::XCHAR>();

/// Owned text source whose representation choice is retained by the plan.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ReturnText {
    Utf8(String),
    Utf16(ExcelString),
}

impl From<String> for ReturnText {
    fn from(value: String) -> Self {
        Self::Utf8(value)
    }
}

impl From<&str> for ReturnText {
    fn from(value: &str) -> Self {
        Self::Utf8(value.to_owned())
    }
}

impl From<ExcelString> for ReturnText {
    fn from(value: ExcelString) -> Self {
        Self::Utf16(value)
    }
}

/// Fully owned logical value offered to the return planner.
#[derive(Clone, Debug, PartialEq)]
pub enum ExcelReturnValue {
    Number(f64),
    Integer(i32),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text(ReturnText),
    Array(ExcelReturnArray),
}

impl ExcelReturnValue {
    pub fn plan(self) -> Result<ReturnPlan, ReturnError> {
        self.plan_with_limits(&ReturnLimits::default())
    }

    pub fn plan_with_limits(self, limits: &ReturnLimits) -> Result<ReturnPlan, ReturnError> {
        Planner::new(limits).plan(self)
    }
}

impl From<ExcelValue> for ExcelReturnValue {
    fn from(value: ExcelValue) -> Self {
        match value {
            ExcelValue::Number(value) => Self::Number(value),
            ExcelValue::Integer(value) => Self::Integer(value),
            ExcelValue::Boolean(value) => Self::Boolean(value),
            ExcelValue::Error(value) => Self::Error(value),
            ExcelValue::Missing => Self::Missing,
            ExcelValue::Empty => Self::Empty,
            ExcelValue::Text(value) => Self::Text(ReturnText::Utf16(value)),
            ExcelValue::Array(value) => Self::Array(value.into()),
        }
    }
}

impl From<ExcelArray> for ExcelReturnValue {
    fn from(value: ExcelArray) -> Self {
        Self::Array(value.into())
    }
}

impl From<ExcelString> for ExcelReturnValue {
    fn from(value: ExcelString) -> Self {
        Self::Text(value.into())
    }
}

impl From<ReturnText> for ExcelReturnValue {
    fn from(value: ReturnText) -> Self {
        Self::Text(value)
    }
}

impl From<String> for ExcelReturnValue {
    fn from(value: String) -> Self {
        Self::Text(value.into())
    }
}

impl From<&str> for ExcelReturnValue {
    fn from(value: &str) -> Self {
        Self::Text(value.into())
    }
}

impl From<ExcelArray> for ExcelReturnArray {
    fn from(value: ExcelArray) -> Self {
        let (rows, columns, values) = value.into_parts();
        Self {
            rows,
            columns,
            values: values
                .into_vec()
                .into_iter()
                .map(ExcelReturnValue::from)
                .collect(),
        }
    }
}

/// Owned flat rectangular logical return array.
///
/// Construction validates shape but deliberately leaves return-specific ABI,
/// nesting, and resource checks to [`ExcelReturnValue::plan_with_limits`].
#[derive(Clone, Debug, PartialEq)]
pub struct ExcelReturnArray {
    rows: usize,
    columns: usize,
    values: Box<[ExcelReturnValue]>,
}

impl ExcelReturnArray {
    pub fn new(
        rows: usize,
        columns: usize,
        values: impl Into<Box<[ExcelReturnValue]>>,
    ) -> Result<Self, ReturnError> {
        let values = values.into();
        let expected = rows
            .checked_mul(columns)
            .ok_or(ReturnError::ArrayElementCountOverflow { rows, columns })?;
        if expected != values.len() {
            return Err(ReturnError::InvalidArrayShape {
                rows,
                columns,
                elements: values.len(),
            });
        }
        Ok(Self {
            rows,
            columns,
            values,
        })
    }

    pub const fn rows(&self) -> usize {
        self.rows
    }

    pub const fn columns(&self) -> usize {
        self.columns
    }

    pub const fn len(&self) -> usize {
        self.values.len()
    }

    pub const fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    pub fn values(&self) -> &[ExcelReturnValue] {
        &self.values
    }
}

/// Independent project limits for constructing ordinary Rust-created returns.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ReturnLimits {
    pub max_string_code_units: usize,
    pub max_array_elements: usize,
    pub max_total_bytes: usize,
    pub max_allocations: usize,
    pub max_depth: usize,
}

impl ReturnLimits {
    pub const DEFAULT: Self = Self {
        max_string_code_units: excel_api_sys::EXCEL12_MAX_STRING_CODE_UNITS,
        max_array_elements: 65_536,
        max_total_bytes: 16 * 1024 * 1024,
        max_allocations: 65_538,
        max_depth: 8,
    };
}

impl Default for ReturnLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Future ownership mechanism selected for the planned return.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ReturnOwnershipStrategy {
    DllOwnedXloper12,
}

/// Exact future ABI storage totals for the chosen Prompt 05 layout.
///
/// `total_bytes` is exactly the sum of initialized ABI root, array-element,
/// and counted-string storage. It excludes Rust container headers and allocator
/// bookkeeping, which are allocator-dependent and are not claimed as heap cost.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ReturnStorageTotals {
    pub root_bytes: usize,
    pub array_element_bytes: usize,
    pub string_payload_code_units: usize,
    pub string_storage_code_units: usize,
    pub string_storage_bytes: usize,
    pub total_bytes: usize,
    pub allocation_count: usize,
}

/// Immutable, fully validated description of a future return allocation.
#[derive(Clone, Debug, PartialEq)]
pub struct ReturnPlan {
    root: PlannedValue,
    totals: ReturnStorageTotals,
    strategy: ReturnOwnershipStrategy,
}

impl ReturnPlan {
    pub const fn root(&self) -> &PlannedValue {
        &self.root
    }

    pub const fn totals(&self) -> &ReturnStorageTotals {
        &self.totals
    }

    pub const fn strategy(&self) -> ReturnOwnershipStrategy {
        self.strategy
    }
}

/// Validated logical root retained by a [`ReturnPlan`].
#[derive(Clone, Debug, PartialEq)]
pub enum PlannedValue {
    Number(f64),
    Integer(i32),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text(PlannedText),
    Array(PlannedArray),
}

/// Original text plus its exact future counted-string sizes.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PlannedText {
    source: ReturnText,
    payload_code_units: usize,
    storage_code_units: usize,
}

impl PlannedText {
    pub const fn source(&self) -> &ReturnText {
        &self.source
    }

    pub const fn payload_code_units(&self) -> usize {
        self.payload_code_units
    }

    pub const fn storage_code_units(&self) -> usize {
        self.storage_code_units
    }
}

/// Validated flat array metadata and elements in row-major order.
#[derive(Clone, Debug, PartialEq)]
pub struct PlannedArray {
    rows: usize,
    columns: usize,
    elements: Box<[PlannedArrayElement]>,
}

impl PlannedArray {
    pub const fn rows(&self) -> usize {
        self.rows
    }

    pub const fn columns(&self) -> usize {
        self.columns
    }

    pub const fn len(&self) -> usize {
        self.elements.len()
    }

    pub const fn is_empty(&self) -> bool {
        self.elements.is_empty()
    }

    pub fn elements(&self) -> &[PlannedArrayElement] {
        &self.elements
    }
}

/// Supported non-array value within a planned `xltypeMulti`.
#[derive(Clone, Debug, PartialEq)]
pub enum PlannedArrayElement {
    Number(f64),
    Integer(i32),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text(PlannedText),
}

struct Planner<'limits> {
    limits: &'limits ReturnLimits,
    array_element_bytes: usize,
    string_payload_code_units: usize,
    string_storage_code_units: usize,
    string_storage_bytes: usize,
    allocation_count: usize,
}

impl<'limits> Planner<'limits> {
    fn new(limits: &'limits ReturnLimits) -> Self {
        Self {
            limits,
            array_element_bytes: 0,
            string_payload_code_units: 0,
            string_storage_code_units: 0,
            string_storage_bytes: 0,
            allocation_count: 1,
        }
    }

    fn plan(mut self, value: ExcelReturnValue) -> Result<ReturnPlan, ReturnError> {
        self.enforce_depth(0)?;
        let root = self.plan_root(value)?;
        let total_bytes = XLOPER12_BYTES
            .checked_add(self.array_element_bytes)
            .and_then(|total| total.checked_add(self.string_storage_bytes))
            .ok_or(ReturnError::TotalByteOverflow)?;
        if total_bytes > self.limits.max_total_bytes {
            return Err(ReturnError::TotalByteLimitExceeded {
                required: total_bytes,
                maximum: self.limits.max_total_bytes,
            });
        }
        if self.allocation_count > self.limits.max_allocations {
            return Err(ReturnError::AllocationCountLimitExceeded {
                required: self.allocation_count,
                maximum: self.limits.max_allocations,
            });
        }
        Ok(ReturnPlan {
            root,
            totals: ReturnStorageTotals {
                root_bytes: XLOPER12_BYTES,
                array_element_bytes: self.array_element_bytes,
                string_payload_code_units: self.string_payload_code_units,
                string_storage_code_units: self.string_storage_code_units,
                string_storage_bytes: self.string_storage_bytes,
                total_bytes,
                allocation_count: self.allocation_count,
            },
            strategy: ReturnOwnershipStrategy::DllOwnedXloper12,
        })
    }

    fn plan_root(&mut self, value: ExcelReturnValue) -> Result<PlannedValue, ReturnError> {
        Ok(match value {
            ExcelReturnValue::Number(value) => PlannedValue::Number(value),
            ExcelReturnValue::Integer(value) => PlannedValue::Integer(value),
            ExcelReturnValue::Boolean(value) => PlannedValue::Boolean(value),
            ExcelReturnValue::Error(value) => PlannedValue::Error(value),
            ExcelReturnValue::Missing => PlannedValue::Missing,
            ExcelReturnValue::Empty => PlannedValue::Empty,
            ExcelReturnValue::Text(value) => PlannedValue::Text(self.plan_text(value)?),
            ExcelReturnValue::Array(value) => PlannedValue::Array(self.plan_array(value)?),
        })
    }

    fn plan_array(&mut self, array: ExcelReturnArray) -> Result<PlannedArray, ReturnError> {
        self.enforce_depth(1)?;
        let count = array.rows.checked_mul(array.columns).ok_or(
            ReturnError::ArrayElementCountOverflow {
                rows: array.rows,
                columns: array.columns,
            },
        )?;
        if count != array.values.len() {
            return Err(ReturnError::InvalidArrayShape {
                rows: array.rows,
                columns: array.columns,
                elements: array.values.len(),
            });
        }
        if array.rows > excel_api_sys::EXCEL12_MAX_ROWS as usize
            || array.columns > excel_api_sys::EXCEL12_MAX_COLUMNS as usize
        {
            return Err(ReturnError::ArrayDimensionExceedsAbi {
                rows: array.rows,
                columns: array.columns,
            });
        }
        if array.rows == 0 || array.columns == 0 {
            return Err(ReturnError::EmptyArrayUnsupported);
        }
        if count > self.limits.max_array_elements {
            return Err(ReturnError::ArrayElementLimitExceeded {
                actual: count,
                maximum: self.limits.max_array_elements,
            });
        }
        self.array_element_bytes = count
            .checked_mul(XLOPER12_BYTES)
            .ok_or(ReturnError::TotalByteOverflow)?;
        self.add_allocation()?;

        let mut elements = Vec::with_capacity(count);
        for value in array.values.into_vec() {
            elements.push(self.plan_array_element(value)?);
        }
        Ok(PlannedArray {
            rows: array.rows,
            columns: array.columns,
            elements: elements.into_boxed_slice(),
        })
    }

    fn plan_array_element(
        &mut self,
        value: ExcelReturnValue,
    ) -> Result<PlannedArrayElement, ReturnError> {
        Ok(match value {
            ExcelReturnValue::Number(value) => PlannedArrayElement::Number(value),
            ExcelReturnValue::Integer(value) => PlannedArrayElement::Integer(value),
            ExcelReturnValue::Boolean(value) => PlannedArrayElement::Boolean(value),
            ExcelReturnValue::Error(value) => PlannedArrayElement::Error(value),
            ExcelReturnValue::Missing => PlannedArrayElement::Missing,
            ExcelReturnValue::Empty => PlannedArrayElement::Empty,
            ExcelReturnValue::Text(value) => PlannedArrayElement::Text(self.plan_text(value)?),
            ExcelReturnValue::Array(_) => return Err(ReturnError::NestedArrayUnsupported),
        })
    }

    fn plan_text(&mut self, source: ReturnText) -> Result<PlannedText, ReturnError> {
        let payload_code_units = match &source {
            ReturnText::Utf8(value) => value.encode_utf16().count(),
            ReturnText::Utf16(value) => value.len_utf16(),
        };
        if payload_code_units > excel_api_sys::EXCEL12_MAX_STRING_CODE_UNITS {
            return Err(ReturnError::StringTooLongForExcel {
                actual: payload_code_units,
                maximum: excel_api_sys::EXCEL12_MAX_STRING_CODE_UNITS,
            });
        }
        if payload_code_units > self.limits.max_string_code_units {
            return Err(ReturnError::StringLimitExceeded {
                actual: payload_code_units,
                maximum: self.limits.max_string_code_units,
            });
        }
        let storage_code_units = payload_code_units
            .checked_add(1)
            .ok_or(ReturnError::TotalByteOverflow)?;
        let storage_bytes = storage_code_units
            .checked_mul(XCHAR_BYTES)
            .ok_or(ReturnError::TotalByteOverflow)?;
        self.string_payload_code_units = self
            .string_payload_code_units
            .checked_add(payload_code_units)
            .ok_or(ReturnError::TotalByteOverflow)?;
        self.string_storage_code_units = self
            .string_storage_code_units
            .checked_add(storage_code_units)
            .ok_or(ReturnError::TotalByteOverflow)?;
        self.string_storage_bytes = self
            .string_storage_bytes
            .checked_add(storage_bytes)
            .ok_or(ReturnError::TotalByteOverflow)?;
        self.add_allocation()?;
        Ok(PlannedText {
            source,
            payload_code_units,
            storage_code_units,
        })
    }

    fn add_allocation(&mut self) -> Result<(), ReturnError> {
        self.allocation_count = self
            .allocation_count
            .checked_add(1)
            .ok_or(ReturnError::AllocationCountOverflow)?;
        Ok(())
    }

    fn enforce_depth(&self, depth: usize) -> Result<(), ReturnError> {
        if depth > self.limits.max_depth {
            return Err(ReturnError::PlanningDepthExceeded {
                depth,
                maximum: self.limits.max_depth,
            });
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn scalar_totals() -> ReturnStorageTotals {
        ReturnStorageTotals {
            root_bytes: XLOPER12_BYTES,
            array_element_bytes: 0,
            string_payload_code_units: 0,
            string_storage_code_units: 0,
            string_storage_bytes: 0,
            total_bytes: XLOPER12_BYTES,
            allocation_count: 1,
        }
    }

    #[test]
    fn every_scalar_and_excel_error_plans_as_one_root() {
        let scalars = [
            ExcelReturnValue::Number(1.25),
            ExcelReturnValue::Integer(-7),
            ExcelReturnValue::Boolean(true),
            ExcelReturnValue::Missing,
            ExcelReturnValue::Empty,
        ];
        for value in scalars {
            let plan = value.plan().unwrap();
            assert_eq!(*plan.totals(), scalar_totals());
            assert_eq!(plan.strategy(), ReturnOwnershipStrategy::DllOwnedXloper12);
        }

        for error in [
            ExcelError::Null,
            ExcelError::Div0,
            ExcelError::Value,
            ExcelError::Ref,
            ExcelError::Name,
            ExcelError::Num,
            ExcelError::Na,
            ExcelError::GettingData,
        ] {
            let plan = ExcelReturnValue::Error(error).plan().unwrap();
            assert_eq!(plan.root(), &PlannedValue::Error(error));
            assert_eq!(*plan.totals(), scalar_totals());
        }
    }

    #[test]
    fn strings_preserve_sources_and_account_for_prefix_exactly() {
        let cases = [
            ReturnText::from(""),
            ReturnText::from("ASCII"),
            ReturnText::from("é水"),
            ReturnText::from("😃"),
            ReturnText::from("A\0B"),
            ReturnText::Utf16(ExcelString::from_utf16_units([0xD83D, 0xDE03])),
            ReturnText::Utf16(ExcelString::from_utf16_units([0xD800])),
            ReturnText::Utf16(ExcelString::from_utf16_units([0xDC00])),
        ];
        for source in cases {
            let expected_units = match &source {
                ReturnText::Utf8(value) => value.encode_utf16().count(),
                ReturnText::Utf16(value) => value.len_utf16(),
            };
            let plan = ExcelReturnValue::Text(source.clone()).plan().unwrap();
            let PlannedValue::Text(text) = plan.root() else {
                panic!("expected planned text");
            };
            assert_eq!(text.source(), &source);
            assert_eq!(text.payload_code_units(), expected_units);
            assert_eq!(text.storage_code_units(), expected_units + 1);
            assert_eq!(plan.totals().string_payload_code_units, expected_units);
            assert_eq!(plan.totals().string_storage_code_units, expected_units + 1);
            assert_eq!(plan.totals().string_storage_bytes, (expected_units + 1) * 2);
            assert_eq!(
                plan.totals().total_bytes,
                XLOPER12_BYTES + (expected_units + 1) * 2
            );
            assert_eq!(plan.totals().allocation_count, 2);
        }

        let utf8 = "A😃";
        assert_ne!(utf8.len(), utf8.encode_utf16().count());
        let plan = ExcelReturnValue::Text(utf8.into()).plan().unwrap();
        assert_eq!(plan.totals().string_payload_code_units, 3);
    }

    #[test]
    fn excel_string_maximum_is_a_hard_boundary() {
        let maximum = excel_api_sys::EXCEL12_MAX_STRING_CODE_UNITS;
        let accepted = ExcelString::from_utf16_units(vec![0xD800; maximum]);
        let plan = ExcelReturnValue::Text(accepted.into()).plan().unwrap();
        assert_eq!(plan.totals().string_payload_code_units, maximum);

        let rejected = ExcelString::from_utf16_units(vec![0_u16; maximum + 1]);
        assert_eq!(
            ExcelReturnValue::Text(rejected.into()).plan(),
            Err(ReturnError::StringTooLongForExcel {
                actual: maximum + 1,
                maximum,
            })
        );
    }

    #[test]
    fn rectangular_array_preserves_order_variants_and_exact_storage() {
        let single = ExcelReturnArray::new(1, 1, vec![ExcelReturnValue::Integer(9)]).unwrap();
        let single = ExcelReturnValue::Array(single).plan().unwrap();
        let PlannedValue::Array(single) = single.root() else {
            panic!("expected planned array");
        };
        assert_eq!((single.rows(), single.columns(), single.len()), (1, 1, 1));

        let values = vec![
            ExcelReturnValue::Number(1.0),
            ExcelReturnValue::Integer(2),
            ExcelReturnValue::Text(ReturnText::from("A")),
            ExcelReturnValue::Error(ExcelError::Na),
            ExcelReturnValue::Boolean(false),
            ExcelReturnValue::Empty,
        ];
        let array = ExcelReturnArray::new(2, 3, values).unwrap();
        let plan = ExcelReturnValue::Array(array).plan().unwrap();
        let PlannedValue::Array(array) = plan.root() else {
            panic!("expected planned array");
        };
        assert_eq!((array.rows(), array.columns(), array.len()), (2, 3, 6));
        assert!(matches!(
            array.elements()[0],
            PlannedArrayElement::Number(1.0)
        ));
        assert!(matches!(
            array.elements()[1],
            PlannedArrayElement::Integer(2)
        ));
        assert!(matches!(array.elements()[2], PlannedArrayElement::Text(_)));
        assert!(matches!(
            array.elements()[3],
            PlannedArrayElement::Error(ExcelError::Na)
        ));
        assert!(matches!(
            array.elements()[4],
            PlannedArrayElement::Boolean(false)
        ));
        assert!(matches!(array.elements()[5], PlannedArrayElement::Empty));
        assert_eq!(plan.totals().array_element_bytes, 6 * XLOPER12_BYTES);
        assert_eq!(plan.totals().string_payload_code_units, 1);
        assert_eq!(plan.totals().string_storage_bytes, 4);
        assert_eq!(
            plan.totals().total_bytes,
            XLOPER12_BYTES + 6 * XLOPER12_BYTES + 4
        );
        assert_eq!(plan.totals().allocation_count, 3);
    }

    #[test]
    fn array_shape_dimension_empty_and_nesting_fail_precisely() {
        assert!(matches!(
            ExcelReturnArray::new(usize::MAX, 2, Vec::<ExcelReturnValue>::new()),
            Err(ReturnError::ArrayElementCountOverflow { .. })
        ));
        assert!(matches!(
            ExcelReturnArray::new(2, 2, vec![ExcelReturnValue::Empty]),
            Err(ReturnError::InvalidArrayShape { .. })
        ));

        let empty = ExcelReturnArray::new(0, 0, Vec::new()).unwrap();
        assert_eq!(
            ExcelReturnValue::Array(empty).plan(),
            Err(ReturnError::EmptyArrayUnsupported)
        );

        let dimension_overflow = ExcelReturnArray {
            rows: excel_api_sys::EXCEL12_MAX_ROWS as usize + 1,
            columns: 1,
            values: Vec::new().into_boxed_slice(),
        };
        assert!(matches!(
            ExcelReturnValue::Array(dimension_overflow).plan(),
            Err(ReturnError::InvalidArrayShape { .. })
        ));

        let inner = ExcelReturnArray::new(1, 1, vec![ExcelReturnValue::Empty]).unwrap();
        let outer = ExcelReturnArray::new(1, 1, vec![ExcelReturnValue::Array(inner)]).unwrap();
        assert_eq!(
            ExcelReturnValue::Array(outer).plan(),
            Err(ReturnError::NestedArrayUnsupported)
        );
    }

    #[test]
    fn dimension_overflow_is_checked_before_project_element_limit() {
        let rows = excel_api_sys::EXCEL12_MAX_ROWS as usize + 1;
        let array = ExcelReturnArray {
            rows,
            columns: 0,
            values: Vec::new().into_boxed_slice(),
        };
        assert_eq!(
            ExcelReturnValue::Array(array).plan(),
            Err(ReturnError::ArrayDimensionExceedsAbi { rows, columns: 0 })
        );
    }

    #[test]
    fn return_limits_accept_boundaries_and_reject_one_over() {
        let limits = ReturnLimits {
            max_string_code_units: 2,
            max_array_elements: 2,
            max_total_bytes: usize::MAX,
            max_allocations: 3,
            max_depth: 8,
        };
        ExcelReturnValue::Text(ReturnText::from("ab"))
            .plan_with_limits(&limits)
            .unwrap();
        assert_eq!(
            ExcelReturnValue::Text(ReturnText::from("abc")).plan_with_limits(&limits),
            Err(ReturnError::StringLimitExceeded {
                actual: 3,
                maximum: 2,
            })
        );

        let two =
            ExcelReturnArray::new(1, 2, vec![ExcelReturnValue::Empty, ExcelReturnValue::Empty])
                .unwrap();
        ExcelReturnValue::Array(two)
            .plan_with_limits(&limits)
            .unwrap();
        let three = ExcelReturnArray::new(1, 3, vec![ExcelReturnValue::Empty; 3]).unwrap();
        assert_eq!(
            ExcelReturnValue::Array(three).plan_with_limits(&limits),
            Err(ReturnError::ArrayElementLimitExceeded {
                actual: 3,
                maximum: 2,
            })
        );

        let exact_bytes = ReturnLimits {
            max_total_bytes: XLOPER12_BYTES,
            ..ReturnLimits::default()
        };
        ExcelReturnValue::Empty
            .plan_with_limits(&exact_bytes)
            .unwrap();
        let one_under = ReturnLimits {
            max_total_bytes: XLOPER12_BYTES - 1,
            ..ReturnLimits::default()
        };
        assert_eq!(
            ExcelReturnValue::Empty.plan_with_limits(&one_under),
            Err(ReturnError::TotalByteLimitExceeded {
                required: XLOPER12_BYTES,
                maximum: XLOPER12_BYTES - 1,
            })
        );
    }

    #[test]
    fn aggregate_strings_allocation_and_depth_limits_are_enforced() {
        let array = ExcelReturnArray::new(
            1,
            2,
            vec![
                ExcelReturnValue::Text(ReturnText::from("A")),
                ExcelReturnValue::Text(ReturnText::Utf16(ExcelString::from_utf16_units([0xD800]))),
            ],
        )
        .unwrap();
        let plan = ExcelReturnValue::Array(array.clone()).plan().unwrap();
        assert_eq!(plan.totals().string_payload_code_units, 2);
        assert_eq!(plan.totals().string_storage_code_units, 4);
        assert_eq!(plan.totals().allocation_count, 4);

        let exact_allocation_limit = ReturnLimits {
            max_allocations: 4,
            ..ReturnLimits::default()
        };
        ExcelReturnValue::Array(array.clone())
            .plan_with_limits(&exact_allocation_limit)
            .unwrap();

        let allocation_limit = ReturnLimits {
            max_allocations: 3,
            ..ReturnLimits::default()
        };
        assert_eq!(
            ExcelReturnValue::Array(array.clone()).plan_with_limits(&allocation_limit),
            Err(ReturnError::AllocationCountLimitExceeded {
                required: 4,
                maximum: 3,
            })
        );

        let byte_limit = ReturnLimits {
            max_total_bytes: plan.totals().total_bytes - 1,
            ..ReturnLimits::default()
        };
        assert!(matches!(
            ExcelReturnValue::Array(array.clone()).plan_with_limits(&byte_limit),
            Err(ReturnError::TotalByteLimitExceeded { .. })
        ));

        let depth_limit = ReturnLimits {
            max_depth: 0,
            ..ReturnLimits::default()
        };
        assert_eq!(
            ExcelReturnValue::Array(array).plan_with_limits(&depth_limit),
            Err(ReturnError::PlanningDepthExceeded {
                depth: 1,
                maximum: 0,
            })
        );
    }

    #[test]
    fn accounting_overflow_paths_are_explicit() {
        let limits = ReturnLimits {
            max_total_bytes: usize::MAX,
            max_allocations: usize::MAX,
            ..ReturnLimits::default()
        };
        let planner = Planner {
            limits: &limits,
            array_element_bytes: usize::MAX,
            string_payload_code_units: 0,
            string_storage_code_units: 0,
            string_storage_bytes: 0,
            allocation_count: 1,
        };
        assert_eq!(
            planner.plan(ExcelReturnValue::Empty),
            Err(ReturnError::TotalByteOverflow)
        );

        let mut planner = Planner {
            limits: &limits,
            array_element_bytes: 0,
            string_payload_code_units: 0,
            string_storage_code_units: 0,
            string_storage_bytes: 0,
            allocation_count: usize::MAX,
        };
        assert_eq!(
            planner.add_allocation(),
            Err(ReturnError::AllocationCountOverflow)
        );
    }

    #[test]
    fn planning_is_deterministic_owned_send_sync_and_pointer_free() {
        fn assert_owned<T: Send + Sync + 'static>() {}
        assert_owned::<ReturnText>();
        assert_owned::<ExcelReturnValue>();
        assert_owned::<ReturnPlan>();

        let value = ExcelReturnValue::Array(
            ExcelReturnArray::new(
                1,
                2,
                vec![
                    ExcelReturnValue::Integer(7),
                    ExcelReturnValue::Text(ReturnText::Utf16(ExcelString::from_utf16_units([
                        0xD800, 0, 0xDC00,
                    ]))),
                ],
            )
            .unwrap(),
        );
        let first = value.clone().plan().unwrap();
        let second = value.plan().unwrap();
        assert_eq!(first, second);

        let PlannedValue::Array(array) = first.root() else {
            panic!("expected array");
        };
        let PlannedArrayElement::Text(text) = &array.elements()[1] else {
            panic!("expected text");
        };
        let ReturnText::Utf16(text) = text.source() else {
            panic!("expected original UTF-16 source");
        };
        assert_eq!(text.as_utf16(), &[0xD800, 0, 0xDC00]);

        // Public planned fields are owned enums, integers, and slices; no API
        // exposes an Excel ABI pointer or a callback lifetime.
        let moved = std::thread::spawn(move || first.totals().total_bytes)
            .join()
            .unwrap();
        assert!(moved >= XLOPER12_BYTES);
    }

    #[test]
    fn semantic_values_convert_without_losing_integer_or_utf16_identity() {
        let semantic = ExcelValue::Array(
            ExcelArray::new(
                1,
                2,
                vec![
                    ExcelValue::Integer(12),
                    ExcelValue::Text(ExcelString::from_utf16_units([0xD800])),
                ],
            )
            .unwrap(),
        );
        let plan = ExcelReturnValue::from(semantic).plan().unwrap();
        let PlannedValue::Array(array) = plan.root() else {
            panic!("expected array");
        };
        assert!(matches!(
            array.elements()[0],
            PlannedArrayElement::Integer(12)
        ));
        let PlannedArrayElement::Text(text) = &array.elements()[1] else {
            panic!("expected text");
        };
        assert!(matches!(text.source(), ReturnText::Utf16(_)));
    }
}
