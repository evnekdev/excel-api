use super::helpers::{chart_bounds, series_data, shape_bounds};
use super::*;

#[test]
fn drawing_values_preserve_unknown_values() {
    assert_eq!(ChartType::from_raw(999).raw(), 999);
    assert_eq!(ShapePlacement::from_raw(-7).raw(), -7);
}

#[test]
fn geometry_validation_rejects_invalid_values() {
    assert!(
        chart_bounds(ChartBounds {
            left: 0.0,
            top: 0.0,
            width: 1.0,
            height: 1.0
        })
        .is_ok()
    );
    assert!(
        chart_bounds(ChartBounds {
            left: f64::NAN,
            top: 0.0,
            width: 1.0,
            height: 1.0
        })
        .is_err()
    );
    assert!(
        shape_bounds(ShapeBounds {
            left: 0.0,
            top: 0.0,
            width: 0.0,
            height: 1.0
        })
        .is_err()
    );
}

#[test]
fn series_formula_rejects_nul_before_com() {
    assert!(matches!(
        series_data(SeriesData::Formula("=A\0")),
        Err(ExcelComError::Unsupported { .. })
    ));
}
