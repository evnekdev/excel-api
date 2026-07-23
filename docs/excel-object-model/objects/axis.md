# Axis

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Axis` |
| GUID | `{00020848-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Axis` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Blocked |

## Capabilities

No capability metadata is recorded for this surface.


## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Border` | `excel.border` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| HasMajorGridlines | PROPERTYGET/PROPERTYPUT | bool | declared | 24 | Metadata Only | Reviewed | Not Tested | |
| HasMinorGridlines | PROPERTYGET/PROPERTYPUT | bool | declared | 25 | Metadata Only | Reviewed | Not Tested | |
| MajorTickMark | PROPERTYGET/PROPERTYPUT | XlTickMark | declared | 26 | Implemented | Reviewed | Live Tested | |
| MinorTickMark | PROPERTYGET/PROPERTYPUT | XlTickMark | declared | 27 | Implemented | Reviewed | Live Tested | |
| TickLabelPosition | PROPERTYGET/PROPERTYPUT | XlTickLabelPosition | declared | 28 | Implemented | Reviewed | Live Tested | |
| TickLabelSpacing | PROPERTYGET/PROPERTYPUT | i32 | declared | 29 | Metadata Only | Reviewed | Not Tested | |
| TickMarkSpacing | PROPERTYGET/PROPERTYPUT | i32 | declared | 31 | Metadata Only | Reviewed | Not Tested | |
| MinimumScale | PROPERTYGET/PROPERTYPUT | f64 | declared | 33 | Implemented | Reviewed | Live Tested | |
| MinimumScaleIsAuto | PROPERTYGET/PROPERTYPUT | bool | declared | 34 | Implemented | Reviewed | Live Tested | |
| MaximumScale | PROPERTYGET/PROPERTYPUT | f64 | declared | 35 | Implemented | Reviewed | Live Tested | |
| MaximumScaleIsAuto | PROPERTYGET/PROPERTYPUT | bool | declared | 36 | Implemented | Reviewed | Live Tested | |
| MajorUnit | PROPERTYGET/PROPERTYPUT | f64 | declared | 37 | Implemented | Reviewed | Live Tested | |
| MajorUnitIsAuto | PROPERTYGET/PROPERTYPUT | bool | declared | 38 | Implemented | Reviewed | Live Tested | |
| MinorUnit | PROPERTYGET/PROPERTYPUT | f64 | declared | 39 | Metadata Only | Reviewed | Not Tested | |
| MinorUnitIsAuto | PROPERTYGET/PROPERTYPUT | bool | declared | 40 | Metadata Only | Reviewed | Not Tested | |
| ScaleType | PROPERTYGET/PROPERTYPUT | XlScaleType | declared | 41 | Implemented | Reviewed | Live Tested | |
| Crosses | PROPERTYGET/PROPERTYPUT | XlAxisCrosses | declared | 42 | Metadata Only | Reviewed | Not Tested | |
| CrossesAt | PROPERTYGET/PROPERTYPUT | f64 | declared | 43 | Metadata Only | Reviewed | Not Tested | |
| ReversePlotOrder | PROPERTYGET/PROPERTYPUT | bool | declared | 44 | Metadata Only | Reviewed | Not Tested | |
| AxisBetweenCategories | PROPERTYGET/PROPERTYPUT | bool | declared | 45 | Metadata Only | Reviewed | Not Tested | |
| AxisGroup | PROPERTYGET | XlAxisGroup | declared | 47 | Metadata Only | Reviewed | Not Tested | |
| HasTitle | PROPERTYGET/PROPERTYPUT | bool | declared | 54 | Implemented | Reviewed | Live Tested | |
| AxisTitle | PROPERTYGET | AxisTitle | declared | 82 | Implemented | Reviewed | Live Tested | |
| MajorGridlines | PROPERTYGET | Gridlines | declared | 89 | Metadata Only | Reviewed | Not Tested | |
| MinorGridlines | PROPERTYGET | Gridlines | declared | 90 | Metadata Only | Reviewed | Not Tested | |
| TickLabels | PROPERTYGET | TickLabels | declared | 91 | Implemented | Reviewed | Live Tested | |
| Type | PROPERTYGET/PROPERTYPUT | XlAxisType | declared | 108 | Metadata Only | Reviewed | Not Tested | |
| Format | PROPERTYGET | ChartFormat | declared | 116 | Metadata Only | Reviewed | Not Tested | |
| Width | PROPERTYGET | f64 | declared | 122 | Metadata Only | Reviewed | Not Tested | |
| Height | PROPERTYGET | f64 | declared | 123 | Metadata Only | Reviewed | Not Tested | |
| Top | PROPERTYGET | f64 | declared | 126 | Metadata Only | Reviewed | Not Tested | |
| Left | PROPERTYGET | f64 | declared | 127 | Metadata Only | Reviewed | Not Tested | |
| Border | PROPERTYGET | Border | declared | 128 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| CategoryNames | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 156 | Metadata Only | Reviewed | Not Tested | |
| BaseUnit | PROPERTYGET/PROPERTYPUT | XlTimeUnit | declared | 1647 | Metadata Only | Reviewed | Not Tested | |
| BaseUnitIsAuto | PROPERTYGET/PROPERTYPUT | bool | declared | 1648 | Metadata Only | Reviewed | Not Tested | |
| MajorUnitScale | PROPERTYGET/PROPERTYPUT | XlTimeUnit | declared | 1649 | Metadata Only | Reviewed | Not Tested | |
| MinorUnitScale | PROPERTYGET/PROPERTYPUT | XlTimeUnit | declared | 1650 | Metadata Only | Reviewed | Not Tested | |
| CategoryType | PROPERTYGET/PROPERTYPUT | XlCategoryType | declared | 1651 | Metadata Only | Reviewed | Not Tested | |
| DisplayUnit | PROPERTYGET/PROPERTYPUT | XlDisplayUnit | declared | 1886 | Metadata Only | Reviewed | Not Tested | |
| DisplayUnitCustom | PROPERTYGET/PROPERTYPUT | f64 | declared | 1887 | Metadata Only | Reviewed | Not Tested | |
| HasDisplayUnitLabel | PROPERTYGET/PROPERTYPUT | bool | declared | 1888 | Metadata Only | Reviewed | Not Tested | |
| DisplayUnitLabel | PROPERTYGET | DisplayUnitLabel | declared | 1889 | Metadata Only | Reviewed | Not Tested | |
| LogBase | PROPERTYGET/PROPERTYPUT | f64 | declared | 2646 | Implemented | Reviewed | Live Tested | |
| TickLabelSpacingIsAuto | PROPERTYGET/PROPERTYPUT | bool | declared | 2647 | Metadata Only | Reviewed | Not Tested | |
| CategorySortOrder | PROPERTYGET/PROPERTYPUT | XlCategorySortOrder | declared | 3228 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Delete | AutomationValue | 0 | declared | 117 | Metadata Only | Reviewed | Not Tested | |
| Select | AutomationValue | 0 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| SetProperty | Unknown | 2 | declared | 3253 | Metadata Only | Reviewed | Not Tested | |
| GetProperty | AutomationValue | 1 | declared | 3256 | Metadata Only | Reviewed | Not Tested | |
| QueryInterface | Unknown | 2 | inherited-iunknown | 1610612736 | Metadata Only | Reviewed | Not Tested | |
| AddRef | Unknown | 0 | inherited-iunknown | 1610612737 | Metadata Only | Reviewed | Not Tested | |
| Release | Unknown | 0 | inherited-iunknown | 1610612738 | Metadata Only | Reviewed | Not Tested | |
| GetTypeInfoCount | Unknown | 1 | inherited-idispatch | 1610678272 | Metadata Only | Reviewed | Not Tested | |
| GetTypeInfo | Unknown | 3 | inherited-idispatch | 1610678273 | Metadata Only | Reviewed | Not Tested | |
| GetIDsOfNames | Unknown | 5 | inherited-idispatch | 1610678274 | Metadata Only | Reviewed | Not Tested | |
| Invoke | Unknown | 8 | inherited-idispatch | 1610678275 | Metadata Only | Reviewed | Not Tested | |

## Events

| Event | Arguments | DISPID | Implementation | Docs | Tests |
|---|---:|---:|---|---|---|
| -- | -- | -- | Not started | Generated | Not tested |

## Unsupported or deferred behaviour

See the global unsupported index for unimplemented object-model areas.
<!-- END GENERATED MEMBERS -->
