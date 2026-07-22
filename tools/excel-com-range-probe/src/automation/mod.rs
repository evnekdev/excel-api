//! Crate-internal semantic Automation value model.
//!
//! This layer owns no COM pointers. It consumes the raw ownership layer and
//! deliberately makes no stable public API commitment.

#![allow(dead_code)] // The transport-facing entry points land in Prompt 07.

use std::fs;
use std::path::Path;
use std::slice;

use serde_json::json;
use windows_sys::Win32::Foundation::{SysStringLen, DISP_E_PARAMNOTFOUND};
use windows_sys::Win32::System::Com::SAFEARRAYBOUND;
use windows_sys::Win32::System::Variant::{
    VT_ARRAY, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_EMPTY, VT_ERROR, VT_I2, VT_I4, VT_I8,
    VT_NULL, VT_R4, VT_R8, VT_VARIANT,
};

use crate::raw::safearray::{
    checked_element_count, get_variant_borrowed, row_column_indices, ObservedSafeArray,
    OwnedSafeArray,
};
use crate::raw::variant::OwnedVariant;

const EXCEL_CELL_STRING_LIMIT: usize = 32_767;

/// A lossless physical Excel error SCODE, not an Excel worksheet error number.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub(crate) struct ExcelError(i32);

#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub(crate) enum ExcelErrorKind {
    Null,
    Div0,
    Value,
    Ref,
    Name,
    Num,
    NotAvailable,
    Other(i32),
}

impl ExcelError {
    pub(crate) const NULL: Self = Self(0x800A_07D0_u32 as i32);
    pub(crate) const DIV0: Self = Self(0x800A_07D7_u32 as i32);
    pub(crate) const VALUE: Self = Self(0x800A_07DF_u32 as i32);
    pub(crate) const REF: Self = Self(0x800A_07E7_u32 as i32);
    pub(crate) const NAME: Self = Self(0x800A_07ED_u32 as i32);
    pub(crate) const NUM: Self = Self(0x800A_07F4_u32 as i32);
    pub(crate) const NOT_AVAILABLE: Self = Self(0x800A_07FA_u32 as i32);

    pub(crate) const fn from_scode(scode: i32) -> Self {
        Self(scode)
    }

    pub(crate) const fn scode(self) -> i32 {
        self.0
    }

    pub(crate) fn kind(self) -> ExcelErrorKind {
        match self {
            Self::NULL => ExcelErrorKind::Null,
            Self::DIV0 => ExcelErrorKind::Div0,
            Self::VALUE => ExcelErrorKind::Value,
            Self::REF => ExcelErrorKind::Ref,
            Self::NAME => ExcelErrorKind::Name,
            Self::NUM => ExcelErrorKind::Num,
            Self::NOT_AVAILABLE => ExcelErrorKind::NotAvailable,
            Self(value) => ExcelErrorKind::Other(value),
        }
    }

    /// Returns the familiar Excel error number only for an `0x800Axxxx` SCODE.
    pub(crate) fn excel_number(self) -> Option<u16> {
        ((self.0 as u32 & 0xffff_0000) == 0x800A_0000)
            .then_some((self.0 as u32 & 0xffff) as u16)
    }

    fn valid_for_direct_write(self) -> bool {
        self.0 < 0
    }
}

/// A finite OLE Automation date serial. It intentionally is not a calendar
/// type: Excel date-system interpretation belongs at a higher layer.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct OaDate(f64);

impl OaDate {
    pub(crate) fn new(serial: f64) -> Result<Self, ConversionError> {
        serial
            .is_finite()
            .then_some(Self(serial))
            .ok_or(ConversionError::NonFiniteNumber)
    }

    pub(crate) const fn serial(self) -> f64 {
        self.0
    }
}

/// Exact COM `CY` storage, whose fixed scale is 10,000.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd, Hash)]
pub(crate) struct Currency(i64);

impl Currency {
    pub(crate) const SCALE: i64 = 10_000;

    pub(crate) const fn from_scaled(scaled: i64) -> Self {
        Self(scaled)
    }

    pub(crate) const fn scaled(self) -> i64 {
        self.0
    }

    pub(crate) fn from_decimal_parts(
        whole: i64,
        fractional_ten_thousandths: u16,
    ) -> Result<Self, ConversionError> {
        if fractional_ten_thousandths >= Self::SCALE as u16 {
            return Err(ConversionError::CurrencyOverflow);
        }
        whole
            .checked_mul(Self::SCALE)
            .and_then(|scaled| {
                scaled.checked_add(if whole < 0 {
                    -(fractional_ten_thousandths as i64)
                } else {
                    fractional_ten_thousandths as i64
                })
            })
            .map(Self)
            .ok_or(ConversionError::CurrencyOverflow)
    }
}

/// Worksheet values. `Missing` is intentionally absent: it belongs only to
/// invocation arguments, never a cell or SAFEARRAY element.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum AutomationValue {
    Empty,
    Null,
    Bool(bool),
    Number(f64),
    Text(String),
    Error(ExcelError),
    Date(OaDate),
    Currency(Currency),
    Array(AutomationArray),
}

/// A zero-based, row-major semantic rectangle. COM bounds stay at the codec
/// boundary; observed SAFEARRAY dimension one maps to rows, two to columns.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct AutomationArray {
    rows: usize,
    columns: usize,
    values: Vec<AutomationValue>,
}

impl AutomationArray {
    pub(crate) fn new(
        rows: usize,
        columns: usize,
        values: Vec<AutomationValue>,
    ) -> Result<Self, ConversionError> {
        (rows.checked_mul(columns) == Some(values.len()))
            .then_some(Self {
                rows,
                columns,
                values,
            })
            .ok_or(ConversionError::InvalidElementCount)
    }

    pub(crate) fn row(values: Vec<AutomationValue>) -> Result<Self, ConversionError> {
        Self::new(1, values.len(), values)
    }

    pub(crate) fn column(values: Vec<AutomationValue>) -> Result<Self, ConversionError> {
        let rows = values.len();
        Self::new(rows, 1, values)
    }

    pub(crate) fn from_rows(rows: Vec<Vec<AutomationValue>>) -> Result<Self, ConversionError> {
        let row_count = rows.len();
        let columns = rows.first().map_or(0, Vec::len);
        if rows.iter().any(|row| row.len() != columns) {
            return Err(ConversionError::InvalidElementCount);
        }
        let values = rows.into_iter().flatten().collect();
        Self::new(row_count, columns, values)
    }

    pub(crate) const fn rows(&self) -> usize {
        self.rows
    }

    pub(crate) const fn columns(&self) -> usize {
        self.columns
    }

    pub(crate) const fn len(&self) -> usize {
        self.values.len()
    }

    pub(crate) const fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    pub(crate) fn get(&self, row: usize, column: usize) -> Option<&AutomationValue> {
        self.index(row, column).and_then(|index| self.values.get(index))
    }

    pub(crate) fn get_mut(
        &mut self,
        row: usize,
        column: usize,
    ) -> Option<&mut AutomationValue> {
        self.index(row, column)
            .and_then(|index| self.values.get_mut(index))
    }

    pub(crate) fn values(&self) -> &[AutomationValue] {
        &self.values
    }

    pub(crate) fn into_values(self) -> Vec<AutomationValue> {
        self.values
    }

    fn index(&self, row: usize, column: usize) -> Option<usize> {
        (row < self.rows && column < self.columns)
            .then(|| row.checked_mul(self.columns)?.checked_add(column))
            .flatten()
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum DateWritePolicy {
    DateVariant,
    Value2Serial,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum ShapePolicy {
    Exact,
}

impl ShapePolicy {
    pub(crate) fn validate(
        self,
        source: &AutomationArray,
        target_rows: usize,
        target_columns: usize,
    ) -> Result<(), ConversionError> {
        match self {
            Self::Exact
                if source.rows == target_rows && source.columns == target_columns => Ok(()),
            Self::Exact => Err(ConversionError::ShapeMismatch {
                source_rows: source.rows,
                source_columns: source.columns,
                target_rows,
                target_columns,
            }),
        }
    }
}

/// Explicit pre-COM conversion behavior. The strict default encodes only
/// deterministic, non-lossy representations proven by the Prompt 05 matrix.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) struct ConversionPolicy {
    pub(crate) date_write: DateWritePolicy,
    pub(crate) reject_non_finite_numbers: bool,
    pub(crate) reject_embedded_nul: bool,
    pub(crate) require_exact_integer_conversion: bool,
    pub(crate) shape: ShapePolicy,
}

impl Default for ConversionPolicy {
    fn default() -> Self {
        Self {
            date_write: DateWritePolicy::DateVariant,
            reject_non_finite_numbers: true,
            reject_embedded_nul: true,
            require_exact_integer_conversion: true,
            shape: ShapePolicy::Exact,
        }
    }
}

/// Pre-COM failures. COM `HRESULT`, `EXCEPINFO`, and Excel application errors
/// deliberately remain in the raw transport layer.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) enum ConversionError {
    UnsupportedVariantType { vartype: u16 },
    UnsupportedSafeArrayRank { rank: u32 },
    UnsupportedSafeArrayElementType { vartype: u16 },
    ShapeMismatch {
        source_rows: usize,
        source_columns: usize,
        target_rows: usize,
        target_columns: usize,
    },
    InvalidElementCount,
    NumericPrecisionLoss,
    NonFiniteNumber,
    EmbeddedNul,
    InvalidDateForPolicy,
    CurrencyOverflow,
    StringTooLong,
    InvalidExcelErrorScode,
    InvalidUtf16String,
    SafeArrayConstructionFailed,
    SafeArrayElementFailed { row: usize, column: usize },
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) enum AutomationArgument {
    Value(AutomationValue),
    Missing,
}

pub(crate) fn encode_argument(
    argument: &AutomationArgument,
    policy: &ConversionPolicy,
) -> Result<OwnedVariant, ConversionError> {
    match argument {
        AutomationArgument::Value(value) => encode_variant(value, policy),
        AutomationArgument::Missing => Ok(OwnedVariant::error(DISP_E_PARAMNOTFOUND)),
    }
}

/// Decodes a raw `VARIANT`, preserving only pointer-free semantic values.
pub(crate) fn decode_variant(
    raw: &OwnedVariant,
    policy: &ConversionPolicy,
) -> Result<AutomationValue, ConversionError> {
    let vartype = raw.vt();
    if vartype == VT_ARRAY | VT_VARIANT {
        return decode_safearray(raw, policy).map(AutomationValue::Array);
    }
    unsafe {
        match vartype {
            VT_EMPTY => Ok(AutomationValue::Empty),
            VT_NULL => Ok(AutomationValue::Null),
            VT_BOOL => Ok(AutomationValue::Bool(
                raw.0.Anonymous.Anonymous.Anonymous.boolVal != 0,
            )),
            VT_I2 => decode_integer(raw.0.Anonymous.Anonymous.Anonymous.iVal as i64, policy),
            VT_I4 => decode_integer(raw.0.Anonymous.Anonymous.Anonymous.lVal as i64, policy),
            VT_I8 => decode_integer(raw.0.Anonymous.Anonymous.Anonymous.llVal, policy),
            VT_R4 => decode_number(raw.0.Anonymous.Anonymous.Anonymous.fltVal as f64, policy),
            VT_R8 => decode_number(raw.0.Anonymous.Anonymous.Anonymous.dblVal, policy),
            VT_BSTR => decode_bstr(raw.0.Anonymous.Anonymous.Anonymous.bstrVal),
            VT_ERROR => Ok(AutomationValue::Error(ExcelError::from_scode(
                raw.0.Anonymous.Anonymous.Anonymous.scode,
            ))),
            VT_DATE => OaDate::new(raw.0.Anonymous.Anonymous.Anonymous.date)
                .map(AutomationValue::Date),
            VT_CY => Ok(AutomationValue::Currency(Currency::from_scaled(
                raw.0.Anonymous.Anonymous.Anonymous.cyVal.int64,
            ))),
            _ => Err(ConversionError::UnsupportedVariantType { vartype }),
        }
    }
}

/// Decodes only rank-two `SAFEARRAY(VARIANT)` values. Lower bounds are
/// normalized away; dimensions one and two become semantic rows and columns.
pub(crate) fn decode_safearray(
    raw: &OwnedVariant,
    policy: &ConversionPolicy,
) -> Result<AutomationArray, ConversionError> {
    if raw.vt() != VT_ARRAY | VT_VARIANT {
        return Err(ConversionError::UnsupportedVariantType { vartype: raw.vt() });
    }
    let pointer = unsafe { raw.0.Anonymous.Anonymous.Anonymous.parray };
    let metadata = unsafe { ObservedSafeArray::inspect(pointer) }
        .ok_or(ConversionError::SafeArrayConstructionFailed)?;
    if metadata.rank != 2 {
        return Err(ConversionError::UnsupportedSafeArrayRank {
            rank: metadata.rank,
        });
    }
    if metadata.element_vartype != Some(VT_VARIANT) {
        return Err(ConversionError::UnsupportedSafeArrayElementType {
            vartype: metadata.element_vartype.unwrap_or_default(),
        });
    }
    let [rows, columns]: &[_; 2] = metadata
        .dimensions
        .as_slice()
        .try_into()
        .map_err(|_| ConversionError::UnsupportedSafeArrayRank {
            rank: metadata.rank,
        })?;
    let rows = usize::try_from(rows.element_count).map_err(|_| ConversionError::InvalidElementCount)?;
    let columns =
        usize::try_from(columns.element_count).map_err(|_| ConversionError::InvalidElementCount)?;
    checked_element_count(&metadata.dimensions).ok_or(ConversionError::InvalidElementCount)?;
    let mut values = Vec::with_capacity(
        rows.checked_mul(columns)
            .ok_or(ConversionError::InvalidElementCount)?,
    );
    for row in 0..rows {
        for column in 0..columns {
            let indices = row_column_indices(
                &metadata.dimensions,
                u32::try_from(row).map_err(|_| ConversionError::InvalidElementCount)?,
                u32::try_from(column).map_err(|_| ConversionError::InvalidElementCount)?,
            )
            .ok_or(ConversionError::InvalidElementCount)?;
            let element = get_variant_borrowed(pointer, &indices).map_err(|_| {
                ConversionError::SafeArrayElementFailed { row, column }
            })?;
            values.push(decode_variant(&element, policy).map_err(|_| {
                ConversionError::SafeArrayElementFailed { row, column }
            })?);
        }
    }
    AutomationArray::new(rows, columns, values)
}

/// Encodes a semantic value with no user-controlled pointer ownership.
pub(crate) fn encode_variant(
    value: &AutomationValue,
    policy: &ConversionPolicy,
) -> Result<OwnedVariant, ConversionError> {
    match value {
        AutomationValue::Empty => Ok(OwnedVariant::empty()),
        AutomationValue::Null => Ok(OwnedVariant::null()),
        AutomationValue::Bool(value) => Ok(OwnedVariant::boolean(*value)),
        AutomationValue::Number(value) => {
            require_finite(*value, policy)?;
            Ok(OwnedVariant::r8(*value))
        }
        AutomationValue::Text(value) => encode_text(value, policy),
        AutomationValue::Error(value) if value.valid_for_direct_write() => {
            Ok(OwnedVariant::error(value.scode()))
        }
        AutomationValue::Error(_) => Err(ConversionError::InvalidExcelErrorScode),
        AutomationValue::Date(value) => match policy.date_write {
            DateWritePolicy::DateVariant if value.serial() < 0.0 => {
                Err(ConversionError::InvalidDateForPolicy)
            }
            DateWritePolicy::DateVariant => Ok(OwnedVariant::date(value.serial())),
            DateWritePolicy::Value2Serial => Ok(OwnedVariant::r8(value.serial())),
        },
        AutomationValue::Currency(value) => Ok(OwnedVariant::currency(value.scaled())),
        AutomationValue::Array(array) => encode_safearray(array, policy),
    }
}

/// Encodes a deterministic one-based rank-two `SAFEARRAY(VARIANT)`.
pub(crate) fn encode_safearray(
    array: &AutomationArray,
    policy: &ConversionPolicy,
) -> Result<OwnedVariant, ConversionError> {
    let rows = u32::try_from(array.rows).map_err(|_| ConversionError::InvalidElementCount)?;
    let columns =
        u32::try_from(array.columns).map_err(|_| ConversionError::InvalidElementCount)?;
    if rows == 0 || columns == 0 {
        return Err(ConversionError::InvalidElementCount);
    }
    let owner = OwnedSafeArray::create_variant(&[
        SAFEARRAYBOUND {
            cElements: rows,
            lLbound: 1,
        },
        SAFEARRAYBOUND {
            cElements: columns,
            lLbound: 1,
        },
    ])
    .map_err(|_| ConversionError::SafeArrayConstructionFailed)?;
    for row in 0..array.rows {
        for column in 0..array.columns {
            let value = array
                .get(row, column)
                .ok_or(ConversionError::InvalidElementCount)?;
            let encoded = encode_scalar(value, policy).map_err(|_| {
                ConversionError::SafeArrayElementFailed { row, column }
            })?;
            let indices = [
                i32::try_from(row + 1).map_err(|_| ConversionError::InvalidElementCount)?,
                i32::try_from(column + 1).map_err(|_| ConversionError::InvalidElementCount)?,
            ];
            owner.put_variant(&indices, &encoded).map_err(|_| {
                ConversionError::SafeArrayElementFailed { row, column }
            })?;
        }
    }
    Ok(OwnedVariant::array(owner))
}

fn encode_scalar(
    value: &AutomationValue,
    policy: &ConversionPolicy,
) -> Result<OwnedVariant, ConversionError> {
    if matches!(value, AutomationValue::Array(_)) {
        return Err(ConversionError::UnsupportedVariantType {
            vartype: VT_ARRAY | VT_VARIANT,
        });
    }
    encode_variant(value, policy)
}

fn decode_integer(value: i64, policy: &ConversionPolicy) -> Result<AutomationValue, ConversionError> {
    let number = value as f64;
    if policy.require_exact_integer_conversion && (number as i128 != value as i128) {
        Err(ConversionError::NumericPrecisionLoss)
    } else {
        Ok(AutomationValue::Number(number))
    }
}

fn decode_number(value: f64, policy: &ConversionPolicy) -> Result<AutomationValue, ConversionError> {
    require_finite(value, policy)?;
    Ok(AutomationValue::Number(value))
}

fn require_finite(value: f64, policy: &ConversionPolicy) -> Result<(), ConversionError> {
    if policy.reject_non_finite_numbers && !value.is_finite() {
        Err(ConversionError::NonFiniteNumber)
    } else {
        Ok(())
    }
}

fn decode_bstr(pointer: *const u16) -> Result<AutomationValue, ConversionError> {
    let length = usize::try_from(unsafe { SysStringLen(pointer) })
        .map_err(|_| ConversionError::StringTooLong)?;
    let units = if pointer.is_null() {
        &[]
    } else {
        unsafe { slice::from_raw_parts(pointer, length) }
    };
    String::from_utf16(units)
        .map(AutomationValue::Text)
        .map_err(|_| ConversionError::InvalidUtf16String)
}

fn encode_text(value: &str, policy: &ConversionPolicy) -> Result<OwnedVariant, ConversionError> {
    if value.chars().count() > EXCEL_CELL_STRING_LIMIT {
        return Err(ConversionError::StringTooLong);
    }
    if value.contains('\0') && policy.reject_embedded_nul {
        return Err(ConversionError::EmbeddedNul);
    }
    if value.contains('\0') {
        return Err(ConversionError::EmbeddedNul);
    }
    OwnedVariant::bstr(value).map_err(|_| ConversionError::SafeArrayConstructionFailed)
}

struct LiveCase {
    id: &'static str,
    member: &'static str,
    address: &'static str,
    input: AutomationValue,
    expected: AutomationValue,
    date_write: DateWritePolicy,
    optional: bool,
}

/// Executes the bounded Prompt 06 compatibility suite. It requires a quiet
/// desktop, uses only L-mode raw transport, and records no process, HWND,
/// pointer, workbook path, or user-session information.
pub(crate) fn live_compatibility(root: &Path, only_case: Option<&str>) -> Result<String, String> {
    if crate::raw::excel_process_count()? != 0 {
        return Err("live compatibility requires pre-existing EXCEL.EXE count = 0".to_owned());
    }
    let cases = live_cases()
        .into_iter()
        .filter(|case| only_case.is_none_or(|id| id == case.id))
        .collect::<Vec<_>>();
    if cases.is_empty() {
        return Err(format!(
            "unknown live compatibility case: {}",
            only_case.unwrap_or("--")
        ));
    }
    let mut rows = Vec::with_capacity(cases.len());
    let mut required_failures = Vec::new();
    for case in cases {
        let policy = ConversionPolicy {
            date_write: case.date_write,
            ..ConversionPolicy::default()
        };
        let row = match encode_variant(&case.input, &policy) {
            Err(error) => json!({
                "schema_version": 1,
                "id": case.id,
                "classification": "Conversion-failed",
                "detail": format!("pre-COM conversion error: {error:?}"),
                "raw_pointer_values_recorded": false,
            }),
            Ok(encoded) => match crate::raw::matrix::semantic_round_trip(
                "L",
                case.member,
                case.address,
                encoded,
            ) {
                Err(error) => json!({
                    "schema_version": 1,
                    "id": case.id,
                    "classification": if case.optional { "Optional-not-accepted" } else { "Runtime-failed" },
                    "detail": error,
                    "raw_pointer_values_recorded": false,
                }),
                Ok(result) => {
                    let decoded = decode_variant(&result.value, &policy);
                    let compatible = result.write_hresult == 0
                        && result.read_hresult == 0
                        && result.clear_hresult == 0
                        && result.owned_process_exit_verified
                        && decoded
                            .as_ref()
                            .is_ok_and(|value| semantic_equivalent(&case.expected, value));
                    json!({
                        "schema_version": 1,
                        "id": case.id,
                        "classification": if compatible { "Runtime-observed" } else if case.optional { "Optional-not-accepted" } else { "Runtime-failed" },
                        "member": case.member,
                        "shape": case.address,
                        "write_hresult": result.write_hresult,
                        "read_hresult": result.read_hresult,
                        "clear_hresult": result.clear_hresult,
                        "read_vartype": result.value.vt(),
                        "decoded": decoded.as_ref().map(semantic_label).unwrap_or_else(|error| format!("decode-error:{error:?}")),
                        "semantic_comparison": if compatible { "compatible-under-documented-normalization" } else { "not-compatible" },
                        "owned_process_exit_verified": result.owned_process_exit_verified,
                        "raw_pointer_values_recorded": false,
                    })
                }
            },
        };
        let completed = row.get("classification").and_then(|value| value.as_str()) == Some("Runtime-observed");
        if !case.optional && !completed {
            required_failures.push(case.id);
        }
        rows.push(row);
        if crate::raw::excel_process_count()? != 0 {
            return Err("an owned live-compatibility Excel child did not exit".to_owned());
        }
    }
    write_jsonl(
        &root.join("automation-value-layer/live-compatibility-observations.jsonl"),
        &rows,
    )?;
    if required_failures.is_empty() {
        Ok(format!(
            "completed {} required live compatibility cases; optional unknown-SCODE result recorded",
            rows.iter()
                .filter(|row| row.get("classification").and_then(|value| value.as_str()) != Some("Optional-not-accepted"))
                .count()
        ))
    } else {
        Err(format!(
            "live compatibility required failures: {}",
            required_failures.join(", ")
        ))
    }
}

fn live_cases() -> Vec<LiveCase> {
    let date = OaDate::new(45_292.5).expect("finite date literal");
    let mixed = AutomationArray::from_rows(vec![
        vec![
            AutomationValue::Number(42.0),
            AutomationValue::Text("Grüße Ω".to_owned()),
            AutomationValue::Bool(true),
        ],
        vec![
            AutomationValue::Empty,
            AutomationValue::Error(ExcelError::NOT_AVAILABLE),
            AutomationValue::Currency(Currency::from_scaled(12_345)),
        ],
    ])
    .expect("literal rectangle");
    let errors = AutomationArray::from_rows(vec![
        vec![
            AutomationValue::Error(ExcelError::NULL),
            AutomationValue::Error(ExcelError::DIV0),
        ],
        vec![
            AutomationValue::Error(ExcelError::VALUE),
            AutomationValue::Error(ExcelError::NOT_AVAILABLE),
        ],
    ])
    .expect("literal rectangle");
    let row = AutomationArray::row(vec![
        AutomationValue::Number(1.0),
        AutomationValue::Number(2.0),
        AutomationValue::Number(3.0),
    ])
    .expect("literal row");
    let column = AutomationArray::column(vec![
        AutomationValue::Number(1.0),
        AutomationValue::Number(2.0),
        AutomationValue::Number(3.0),
    ])
    .expect("literal column");
    vec![
        LiveCase { id: "live.scalar-number", member: "Value2", address: "A1", input: AutomationValue::Number(42.0), expected: AutomationValue::Number(42.0), date_write: DateWritePolicy::Value2Serial, optional: false },
        LiveCase { id: "live.scalar-unicode", member: "Value2", address: "A1", input: AutomationValue::Text("Grüße Ω 😀".to_owned()), expected: AutomationValue::Text("Grüße Ω 😀".to_owned()), date_write: DateWritePolicy::Value2Serial, optional: false },
        LiveCase { id: "live.scalar-na", member: "Value2", address: "A1", input: AutomationValue::Error(ExcelError::NOT_AVAILABLE), expected: AutomationValue::Error(ExcelError::NOT_AVAILABLE), date_write: DateWritePolicy::Value2Serial, optional: false },
        LiveCase { id: "live.scalar-unknown-error", member: "Value2", address: "A1", input: AutomationValue::Error(ExcelError::from_scode(0x8123_4567_u32 as i32)), expected: AutomationValue::Error(ExcelError::from_scode(0x8123_4567_u32 as i32)), date_write: DateWritePolicy::Value2Serial, optional: true },
        LiveCase { id: "live.scalar-date-value2", member: "Value2", address: "A1", input: AutomationValue::Date(date), expected: AutomationValue::Number(date.serial()), date_write: DateWritePolicy::Value2Serial, optional: false },
        LiveCase { id: "live.scalar-date-value", member: "Value", address: "A1", input: AutomationValue::Date(date), expected: AutomationValue::Date(date), date_write: DateWritePolicy::DateVariant, optional: false },
        LiveCase { id: "live.scalar-currency", member: "Value", address: "A1", input: AutomationValue::Currency(Currency::from_scaled(1_234_500)), expected: AutomationValue::Currency(Currency::from_scaled(1_234_500)), date_write: DateWritePolicy::DateVariant, optional: false },
        LiveCase { id: "live.array-mixed-2x3", member: "Value2", address: "A1:C2", input: AutomationValue::Array(mixed.clone()), expected: AutomationValue::Array(mixed), date_write: DateWritePolicy::Value2Serial, optional: false },
        LiveCase { id: "live.array-errors-2x2", member: "Value2", address: "A1:B2", input: AutomationValue::Array(errors.clone()), expected: AutomationValue::Array(errors), date_write: DateWritePolicy::Value2Serial, optional: false },
        LiveCase { id: "live.array-row-1x3", member: "Value2", address: "A1:C1", input: AutomationValue::Array(row.clone()), expected: AutomationValue::Array(row), date_write: DateWritePolicy::Value2Serial, optional: false },
        LiveCase { id: "live.array-column-3x1", member: "Value2", address: "A1:A3", input: AutomationValue::Array(column.clone()), expected: AutomationValue::Array(column), date_write: DateWritePolicy::Value2Serial, optional: false },
    ]
}

fn semantic_equivalent(expected: &AutomationValue, actual: &AutomationValue) -> bool {
    match (expected, actual) {
        (AutomationValue::Date(expected), AutomationValue::Number(actual)) => expected.serial() == *actual,
        (AutomationValue::Currency(expected), AutomationValue::Number(actual)) => {
            expected.scaled() as f64 / Currency::SCALE as f64 == *actual
        }
        (AutomationValue::Array(expected), AutomationValue::Array(actual)) => {
            expected.rows == actual.rows
                && expected.columns == actual.columns
                && expected
                    .values
                    .iter()
                    .zip(&actual.values)
                    .all(|(expected, actual)| semantic_equivalent(expected, actual))
        }
        _ => expected == actual,
    }
}

fn semantic_label(value: &AutomationValue) -> String {
    match value {
        AutomationValue::Empty => "Empty".to_owned(),
        AutomationValue::Null => "Null".to_owned(),
        AutomationValue::Bool(value) => format!("Bool({value})"),
        AutomationValue::Number(value) => format!("Number({value})"),
        AutomationValue::Text(value) => format!("Text(chars={})", value.chars().count()),
        AutomationValue::Error(value) => format!("Error(0x{:08X})", value.scode() as u32),
        AutomationValue::Date(value) => format!("Date({})", value.serial()),
        AutomationValue::Currency(value) => format!("Currency({})", value.scaled()),
        AutomationValue::Array(value) => format!("Array({}x{})", value.rows, value.columns),
    }
}

fn write_jsonl(path: &Path, rows: &[serde_json::Value]) -> Result<(), String> {
    let mut rows = rows.to_vec();
    rows.sort_by_key(|row| row.get("id").and_then(|value| value.as_str()).unwrap_or("").to_owned());
    let text = rows
        .into_iter()
        .map(|row| serde_json::to_string(&row).map_err(|error| error.to_string()))
        .collect::<Result<Vec<_>, _>>()?
        .join("\n");
    fs::write(path, format!("{text}\n")).map_err(|error| error.to_string())
}

/// Checks static Prompt 06 evidence without opening Excel or mutating files.
pub(crate) fn check_evidence(root: &Path) -> Result<(), String> {
    let source = root.join("automation-value-layer");
    let generated = root.join("generated/automation-value-layer");
    for file in [
        "SOURCE_MANIFEST.toml",
        "design-decisions.jsonl",
        "conversion-policies.jsonl",
        "codec-cases.jsonl",
        "live-compatibility-observations.jsonl",
        "unresolved.jsonl",
    ] {
        let text = fs::read_to_string(source.join(file)).map_err(|error| error.to_string())?;
        if text.contains('\r') || !text.ends_with('\n') {
            return Err(format!("automation-value-layer/{file} must use LF and a final newline"));
        }
        if file.ends_with(".jsonl") {
            for line in text.lines() {
                serde_json::from_str::<serde_json::Value>(line)
                    .map_err(|error| format!("automation-value-layer/{file}: {error}"))?;
            }
        }
    }
    for file in [
        "value-model.md",
        "excel-errors.md",
        "dates.md",
        "currency.md",
        "arrays.md",
        "conversion-policies.md",
        "conversion-errors.md",
        "codec-test-matrix.md",
        "live-compatibility.md",
        "remaining-blockers.md",
    ] {
        let text = fs::read_to_string(generated.join(file)).map_err(|error| error.to_string())?;
        if text.contains('\r') || !text.ends_with('\n') || !text.starts_with("# ") {
            return Err(format!("generated/automation-value-layer/{file} is not a valid deterministic report"));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::System::Com::SAFEARRAYBOUND;

    fn strict() -> ConversionPolicy {
        ConversionPolicy::default()
    }

    fn round_trip(value: AutomationValue) -> AutomationValue {
        let policy = strict();
        let raw = encode_variant(&value, &policy).expect("encode");
        decode_variant(&raw, &policy).expect("decode")
    }

    #[test]
    fn standard_errors_are_exact_signed_scodes() {
        let expected = [
            (ExcelError::NULL, 0x800A_07D0, ExcelErrorKind::Null),
            (ExcelError::DIV0, 0x800A_07D7, ExcelErrorKind::Div0),
            (ExcelError::VALUE, 0x800A_07DF, ExcelErrorKind::Value),
            (ExcelError::REF, 0x800A_07E7, ExcelErrorKind::Ref),
            (ExcelError::NAME, 0x800A_07ED, ExcelErrorKind::Name),
            (ExcelError::NUM, 0x800A_07F4, ExcelErrorKind::Num),
            (ExcelError::NOT_AVAILABLE, 0x800A_07FA, ExcelErrorKind::NotAvailable),
        ];
        for (error, bits, kind) in expected {
            assert_eq!(error.scode() as u32, bits);
            assert_eq!(error.kind(), kind);
            assert_eq!(round_trip(AutomationValue::Error(error)), AutomationValue::Error(error));
        }
    }

    #[test]
    fn unknown_negative_scode_round_trips_but_short_excel_number_cannot_write() {
        let unknown = ExcelError::from_scode(0x8123_4567_u32 as i32);
        assert_eq!(unknown.kind(), ExcelErrorKind::Other(unknown.scode()));
        assert_eq!(round_trip(AutomationValue::Error(unknown)), AutomationValue::Error(unknown));
        assert!(matches!(
            encode_variant(&AutomationValue::Error(ExcelError::from_scode(2042)), &strict()),
            Err(ConversionError::InvalidExcelErrorScode)
        ));
    }

    #[test]
    fn scalar_round_trips_preserve_distinctions_and_unicode() {
        let cases = vec![
            AutomationValue::Empty,
            AutomationValue::Null,
            AutomationValue::Bool(false),
            AutomationValue::Bool(true),
            AutomationValue::Number(0.0),
            AutomationValue::Number(-0.0),
            AutomationValue::Number(-42.5),
            AutomationValue::Text(String::new()),
            AutomationValue::Text("Grüße Ω 😀".to_owned()),
            AutomationValue::Date(OaDate::new(45_292.5).unwrap()),
            AutomationValue::Currency(Currency::from_scaled(0)),
            AutomationValue::Currency(Currency::from_scaled(-1_234_500)),
        ];
        for value in cases {
            assert_eq!(round_trip(value.clone()), value);
        }
    }

    #[test]
    fn finite_number_property_samples_round_trip_bitwise() {
        for bits in [0_u64, 1, 0x8000_0000_0000_0000, 0x3ff0_0000_0000_0000, 0xc05e_d000_0000_0000] {
            let value = f64::from_bits(bits);
            if value.is_finite() {
                assert_eq!(round_trip(AutomationValue::Number(value)), AutomationValue::Number(value));
            }
        }
    }

    #[test]
    fn strict_policy_rejects_lossy_and_unsafe_scalars() {
        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert!(matches!(
                encode_variant(&AutomationValue::Number(value), &strict()),
                Err(ConversionError::NonFiniteNumber)
            ));
        }
        assert!(matches!(
            encode_variant(&AutomationValue::Text("left\0right".to_owned()), &strict()),
            Err(ConversionError::EmbeddedNul)
        ));
        let too_long = "x".repeat(EXCEL_CELL_STRING_LIMIT + 1);
        assert!(matches!(
            encode_variant(&AutomationValue::Text(too_long), &strict()),
            Err(ConversionError::StringTooLong)
        ));
        assert_eq!(
            decode_variant(&OwnedVariant::i8(9_007_199_254_740_993), &strict()),
            Err(ConversionError::NumericPrecisionLoss)
        );
    }

    #[test]
    fn dates_have_explicit_negative_date_policy() {
        let negative = AutomationValue::Date(OaDate::new(-1.25).unwrap());
        assert!(matches!(
            encode_variant(&negative, &strict()),
            Err(ConversionError::InvalidDateForPolicy)
        ));
        let policy = ConversionPolicy {
            date_write: DateWritePolicy::Value2Serial,
            ..strict()
        };
        let encoded = encode_variant(&negative, &policy).unwrap();
        assert_eq!(encoded.vt(), VT_R8);
        assert_eq!(decode_variant(&encoded, &policy), Ok(AutomationValue::Number(-1.25)));
    }

    #[test]
    fn currency_is_exact_and_checked() {
        assert_eq!(Currency::from_decimal_parts(123, 4500).unwrap().scaled(), 1_234_500);
        assert_eq!(Currency::from_decimal_parts(-123, 4500).unwrap().scaled(), -1_234_500);
        assert_eq!(
            Currency::from_decimal_parts(i64::MAX, 1),
            Err(ConversionError::CurrencyOverflow)
        );
        for scaled in [-1_i64, 0, 1, i64::MAX, i64::MIN] {
            assert_eq!(round_trip(AutomationValue::Currency(Currency::from_scaled(scaled))), AutomationValue::Currency(Currency::from_scaled(scaled)));
        }
    }

    #[test]
    fn rectangular_array_accessors_and_shape_policy_are_exact() {
        let mut values = AutomationArray::from_rows(vec![
            vec![AutomationValue::Number(1.0), AutomationValue::Number(2.0)],
            vec![AutomationValue::Number(3.0), AutomationValue::Number(4.0)],
        ])
        .unwrap();
        assert_eq!(values.rows(), 2);
        assert_eq!(values.columns(), 2);
        assert_eq!(values.len(), 4);
        assert!(!values.is_empty());
        assert_eq!(values.get(1, 0), Some(&AutomationValue::Number(3.0)));
        *values.get_mut(0, 1).unwrap() = AutomationValue::Number(20.0);
        assert_eq!(values.values()[1], AutomationValue::Number(20.0));
        assert!(matches!(
            strict().shape.validate(&values, 1, 4),
            Err(ConversionError::ShapeMismatch { .. })
        ));
        assert_eq!(AutomationArray::from_rows(vec![vec![AutomationValue::Empty], vec![]]), Err(ConversionError::InvalidElementCount));
    }

    #[test]
    fn safearray_round_trips_mixed_values_in_row_major_order() {
        let values = AutomationArray::from_rows(vec![
            vec![
                AutomationValue::Empty,
                AutomationValue::Null,
                AutomationValue::Bool(true),
            ],
            vec![
                AutomationValue::Text("Ω".to_owned()),
                AutomationValue::Error(ExcelError::NOT_AVAILABLE),
                AutomationValue::Currency(Currency::from_scaled(12_345)),
            ],
        ])
        .unwrap();
        let decoded = round_trip(AutomationValue::Array(values.clone()));
        assert_eq!(decoded, AutomationValue::Array(values));
    }

    #[test]
    fn rank_one_array_is_rejected_and_non_one_bounds_decode() {
        let vector = OwnedSafeArray::create_variant_vector(-4, 2).unwrap();
        vector.put_variant(&[-4], &OwnedVariant::i4(1)).unwrap();
        vector.put_variant(&[-3], &OwnedVariant::i4(2)).unwrap();
        let raw = OwnedVariant::array(vector);
        assert_eq!(
            decode_safearray(&raw, &strict()),
            Err(ConversionError::UnsupportedSafeArrayRank { rank: 1 })
        );

        let owner = OwnedSafeArray::create_variant(&[
            SAFEARRAYBOUND { cElements: 2, lLbound: 7 },
            SAFEARRAYBOUND { cElements: 2, lLbound: -3 },
        ])
        .unwrap();
        for (indices, value) in [([7, -3], 1), ([7, -2], 2), ([8, -3], 3), ([8, -2], 4)] {
            owner.put_variant(&indices, &OwnedVariant::i4(value)).unwrap();
        }
        let raw = OwnedVariant::array(owner);
        assert_eq!(
            decode_safearray(&raw, &strict()),
            Ok(AutomationArray::from_rows(vec![
                vec![AutomationValue::Number(1.0), AutomationValue::Number(2.0)],
                vec![AutomationValue::Number(3.0), AutomationValue::Number(4.0)],
            ])
            .unwrap())
        );
    }

    #[test]
    fn array_failure_identifies_element_and_nested_arrays_are_not_allowed() {
        let nested = AutomationArray::row(vec![AutomationValue::Empty]).unwrap();
        let array = AutomationArray::row(vec![AutomationValue::Array(nested)]).unwrap();
        assert!(matches!(
            encode_safearray(&array, &strict()),
            Err(ConversionError::SafeArrayElementFailed { row: 0, column: 0 })
        ));
    }

    #[test]
    fn missing_is_not_an_excel_error() {
        let missing = encode_argument(&AutomationArgument::Missing, &strict()).unwrap();
        let error = encode_argument(
            &AutomationArgument::Value(AutomationValue::Error(ExcelError::NOT_AVAILABLE)),
            &strict(),
        )
        .unwrap();
        assert_eq!(missing.vt(), VT_ERROR);
        assert_eq!(missing.error_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(error.vt(), VT_ERROR);
        assert_eq!(error.error_scode(), Some(ExcelError::NOT_AVAILABLE.scode()));
    }
}
