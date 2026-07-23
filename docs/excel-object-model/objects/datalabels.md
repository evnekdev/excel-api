# DataLabels

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `DataLabels` |
| GUID | `{000208b3-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::DataLabels` |
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
| `Font` | `excel.font` | Metadata Only |
| `Interior` | `excel.interior` | Metadata Only |
| `Item` | `excel.datalabel` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Shadow | PROPERTYGET/PROPERTYPUT | bool | declared | 103 | Metadata Only | Reviewed | Not Tested | |
| Type | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 108 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET | String | declared | 110 | Metadata Only | Reviewed | Not Tested | |
| Format | PROPERTYGET | ChartFormat | declared | 116 | Metadata Only | Reviewed | Not Tested | |
| Count | PROPERTYGET | i32 | declared | 118 | Implemented | Reviewed | Live Tested | |
| Border | PROPERTYGET | Border | declared | 128 | Metadata Only | Reviewed | Not Tested | |
| Interior | PROPERTYGET | Interior | declared | 129 | Metadata Only | Reviewed | Not Tested | |
| Position | PROPERTYGET/PROPERTYPUT | XlDataLabelPosition | declared | 133 | Metadata Only | Reviewed | Not Tested | |
| Orientation | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 134 | Metadata Only | Reviewed | Not Tested | |
| AutoText | PROPERTYGET/PROPERTYPUT | bool | declared | 135 | Metadata Only | Reviewed | Not Tested | |
| HorizontalAlignment | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 136 | Metadata Only | Reviewed | Not Tested | |
| VerticalAlignment | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 137 | Metadata Only | Reviewed | Not Tested | |
| Font | PROPERTYGET | Font | declared | 146 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| ShowLegendKey | PROPERTYGET/PROPERTYPUT | bool | declared | 171 | Metadata Only | Reviewed | Not Tested | |
| NumberFormat | PROPERTYGET/PROPERTYPUT | String | declared | 193 | Metadata Only | Reviewed | Not Tested | |
| NumberFormatLinked | PROPERTYGET/PROPERTYPUT | bool | declared | 194 | Metadata Only | Reviewed | Not Tested | |
| ReadingOrder | PROPERTYGET/PROPERTYPUT | i32 | declared | 975 | Metadata Only | Reviewed | Not Tested | |
| NumberFormatLocal | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1097 | Metadata Only | Reviewed | Not Tested | |
| AutoScaleFont | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1525 | Metadata Only | Reviewed | Not Tested | |
| Fill | PROPERTYGET | ChartFillFormat | declared | 1663 | Metadata Only | Reviewed | Not Tested | |
| ShowSeriesName | PROPERTYGET/PROPERTYPUT | bool | declared | 2022 | Metadata Only | Reviewed | Not Tested | |
| ShowCategoryName | PROPERTYGET/PROPERTYPUT | bool | declared | 2023 | Metadata Only | Reviewed | Not Tested | |
| ShowValue | PROPERTYGET/PROPERTYPUT | bool | declared | 2024 | Metadata Only | Reviewed | Not Tested | |
| ShowPercentage | PROPERTYGET/PROPERTYPUT | bool | declared | 2025 | Metadata Only | Reviewed | Not Tested | |
| ShowBubbleSize | PROPERTYGET/PROPERTYPUT | bool | declared | 2026 | Metadata Only | Reviewed | Not Tested | |
| Separator | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2027 | Metadata Only | Reviewed | Not Tested | |
| ShowRange | PROPERTYGET/PROPERTYPUT | bool | declared | 3086 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| _NewEnum | Unknown | 0 | declared | -4 | Metadata Only | Reviewed | Not Tested | |
| _Default | DataLabel | 1 | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| Delete | AutomationValue | 0 | declared | 117 | Metadata Only | Reviewed | Not Tested | |
| Item | DataLabel | 1 | declared | 170 | Implemented | Reviewed | Live Tested | |
| Select | AutomationValue | 0 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| Propagate | Unknown | 1 | declared | 3087 | Metadata Only | Reviewed | Not Tested | |
| SetProperty | Unknown | 2 | declared | 3253 | Metadata Only | Reviewed | Not Tested | |
| GetProperty | AutomationValue | 1 | declared | 3256 | Metadata Only | Reviewed | Not Tested | |
| _Dummy9 | Unknown | 0 | declared | 65545 | Metadata Only | Reviewed | Not Tested | |
| _Dummy10 | Unknown | 0 | declared | 65546 | Metadata Only | Reviewed | Not Tested | |
| _Dummy13 | Unknown | 0 | declared | 65549 | Metadata Only | Reviewed | Not Tested | |
| _Dummy16 | Unknown | 0 | declared | 65552 | Metadata Only | Reviewed | Not Tested | |
| _Dummy17 | Unknown | 0 | declared | 65553 | Metadata Only | Reviewed | Not Tested | |
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
