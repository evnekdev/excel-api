# PlotArea

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `PlotArea` |
| GUID | `{000208cb-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::PlotArea` |
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
| `Interior` | `excel.interior` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Name | PROPERTYGET | String | declared | 110 | Metadata Only | Reviewed | Not Tested | |
| Format | PROPERTYGET | ChartFormat | declared | 116 | Implemented | Reviewed | Live Tested | |
| Width | PROPERTYGET/PROPERTYPUT | f64 | declared | 122 | Metadata Only | Reviewed | Not Tested | |
| Height | PROPERTYGET/PROPERTYPUT | f64 | declared | 123 | Metadata Only | Reviewed | Not Tested | |
| Top | PROPERTYGET/PROPERTYPUT | f64 | declared | 126 | Metadata Only | Reviewed | Not Tested | |
| Left | PROPERTYGET/PROPERTYPUT | f64 | declared | 127 | Metadata Only | Reviewed | Not Tested | |
| Border | PROPERTYGET | Border | declared | 128 | Metadata Only | Reviewed | Not Tested | |
| Interior | PROPERTYGET | Interior | declared | 129 | Metadata Only | Reviewed | Not Tested | |
| Position | PROPERTYGET/PROPERTYPUT | XlChartElementPosition | declared | 133 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Fill | PROPERTYGET | ChartFillFormat | declared | 1663 | Metadata Only | Reviewed | Not Tested | |
| InsideLeft | PROPERTYGET/PROPERTYPUT | f64 | declared | 1667 | Metadata Only | Reviewed | Not Tested | |
| InsideTop | PROPERTYGET/PROPERTYPUT | f64 | declared | 1668 | Metadata Only | Reviewed | Not Tested | |
| InsideWidth | PROPERTYGET/PROPERTYPUT | f64 | declared | 1669 | Metadata Only | Reviewed | Not Tested | |
| InsideHeight | PROPERTYGET/PROPERTYPUT | f64 | declared | 1670 | Metadata Only | Reviewed | Not Tested | |
| _InsideLeft | PROPERTYGET | f64 | declared | 2654 | Metadata Only | Reviewed | Not Tested | |
| _InsideTop | PROPERTYGET | f64 | declared | 2655 | Metadata Only | Reviewed | Not Tested | |
| _InsideWidth | PROPERTYGET | f64 | declared | 2656 | Metadata Only | Reviewed | Not Tested | |
| _InsideHeight | PROPERTYGET | f64 | declared | 2657 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| ClearFormats | AutomationValue | 0 | declared | 112 | Metadata Only | Reviewed | Not Tested | |
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
