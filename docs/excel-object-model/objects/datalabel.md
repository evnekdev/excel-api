# DataLabel

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `DataLabel` |
| GUID | `{000208b2-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::DataLabel` |
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
| `Font` | `excel.font` | Implemented |
| `Interior` | `excel.interior` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Shadow | PROPERTYGET/PROPERTYPUT | bool | declared | 103 | Metadata Only | Reviewed | Not Tested | |
| Type | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 108 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET | String | declared | 110 | Metadata Only | Reviewed | Not Tested | |
| Format | PROPERTYGET | ChartFormat | declared | 116 | Implemented | Reviewed | Live Tested | |
| Width | PROPERTYGET/PROPERTYPUT | f64 | declared | 122 | Metadata Only | Reviewed | Not Tested | |
| Height | PROPERTYGET/PROPERTYPUT | f64 | declared | 123 | Metadata Only | Reviewed | Not Tested | |
| Top | PROPERTYGET/PROPERTYPUT | f64 | declared | 126 | Metadata Only | Reviewed | Not Tested | |
| Left | PROPERTYGET/PROPERTYPUT | f64 | declared | 127 | Metadata Only | Reviewed | Not Tested | |
| Border | PROPERTYGET | Border | declared | 128 | Metadata Only | Reviewed | Not Tested | |
| Interior | PROPERTYGET | Interior | declared | 129 | Metadata Only | Reviewed | Not Tested | |
| Position | PROPERTYGET/PROPERTYPUT | XlDataLabelPosition | declared | 133 | Implemented | Reviewed | Live Tested | |
| Orientation | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 134 | Metadata Only | Reviewed | Not Tested | |
| AutoText | PROPERTYGET/PROPERTYPUT | bool | declared | 135 | Metadata Only | Reviewed | Not Tested | |
| HorizontalAlignment | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 136 | Metadata Only | Reviewed | Not Tested | |
| VerticalAlignment | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 137 | Metadata Only | Reviewed | Not Tested | |
| Text | PROPERTYGET/PROPERTYPUT | String | declared | 138 | Implemented | Reviewed | Live Tested | |
| Caption | PROPERTYGET/PROPERTYPUT | String | declared | 139 | Metadata Only | Reviewed | Not Tested | |
| Font | PROPERTYGET | Font | declared | 146 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| ShowLegendKey | PROPERTYGET/PROPERTYPUT | bool | declared | 171 | Implemented | Reviewed | Live Tested | |
| NumberFormat | PROPERTYGET/PROPERTYPUT | String | declared | 193 | Implemented | Reviewed | Live Tested | |
| NumberFormatLinked | PROPERTYGET/PROPERTYPUT | bool | declared | 194 | Metadata Only | Reviewed | Not Tested | |
| Formula | PROPERTYGET/PROPERTYPUT | String | declared | 261 | Metadata Only | Reviewed | Not Tested | |
| FormulaLocal | PROPERTYGET/PROPERTYPUT | String | declared | 263 | Metadata Only | Reviewed | Not Tested | |
| FormulaR1C1 | PROPERTYGET/PROPERTYPUT | String | declared | 264 | Metadata Only | Reviewed | Not Tested | |
| FormulaR1C1Local | PROPERTYGET/PROPERTYPUT | String | declared | 265 | Metadata Only | Reviewed | Not Tested | |
| Characters | PROPERTYGET | Characters | declared | 603 | Metadata Only | Reviewed | Not Tested | |
| ReadingOrder | PROPERTYGET/PROPERTYPUT | i32 | declared | 975 | Metadata Only | Reviewed | Not Tested | |
| NumberFormatLocal | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1097 | Metadata Only | Reviewed | Not Tested | |
| AutoScaleFont | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1525 | Metadata Only | Reviewed | Not Tested | |
| Fill | PROPERTYGET | ChartFillFormat | declared | 1663 | Metadata Only | Reviewed | Not Tested | |
| ShowSeriesName | PROPERTYGET/PROPERTYPUT | bool | declared | 2022 | Implemented | Reviewed | Live Tested | |
| ShowCategoryName | PROPERTYGET/PROPERTYPUT | bool | declared | 2023 | Implemented | Reviewed | Live Tested | |
| ShowValue | PROPERTYGET/PROPERTYPUT | bool | declared | 2024 | Implemented | Reviewed | Live Tested | |
| ShowPercentage | PROPERTYGET/PROPERTYPUT | bool | declared | 2025 | Implemented | Reviewed | Live Tested | |
| ShowBubbleSize | PROPERTYGET/PROPERTYPUT | bool | declared | 2026 | Implemented | Reviewed | Live Tested | |
| Separator | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2027 | Implemented | Reviewed | Live Tested | |
| _Height | PROPERTYGET | f64 | declared | 3084 | Metadata Only | Reviewed | Not Tested | |
| _Width | PROPERTYGET | f64 | declared | 3085 | Metadata Only | Reviewed | Not Tested | |
| ShowRange | PROPERTYGET/PROPERTYPUT | bool | declared | 3086 | Metadata Only | Reviewed | Not Tested | |

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
