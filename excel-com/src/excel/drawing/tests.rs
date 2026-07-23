use super::helpers::{chart_bounds, series_data, shape_bounds};
use super::*;

#[test]
fn drawing_values_preserve_unknown_values() {
    assert_eq!(ChartType::from_raw(999).raw(), 999);
    assert_eq!(ShapePlacement::from_raw(-7).raw(), -7);
}

#[test]
fn chart_type_constants_match_the_installed_typelib_inventory() {
    assert_eq!(ChartType::LINE_3D.raw(), -4101);
    assert_eq!(ChartType::PIE_3D.raw(), -4102);
    assert_eq!(ChartType::PIE_3D_EXPLODED.raw(), 70);
    assert_eq!(ChartType::CYLINDER_COLUMN_CLUSTERED.raw(), 92);
    assert_eq!(ChartType::PYRAMID_COLUMN_3D.raw(), 112);
    assert_eq!(ChartType::FUNNEL.raw(), 123);
    assert_eq!(ChartType::XY_SCATTER_SMOOTH.raw(), 72);
    assert_eq!(ChartType::REGION_MAP.raw(), 140);
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
