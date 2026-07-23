# ChartObject

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `ChartObject` |
| GUID | `{000208cf-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::ChartObject` |
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
| `BottomRightCell` | `excel.range` | Metadata Only |
| `Interior` | `excel.interior` | Metadata Only |
| `TopLeftCell` | `excel.range` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Chart | PROPERTYGET | Chart | declared | 7 | Implemented | Reviewed | Live Tested | |
| Shadow | PROPERTYGET/PROPERTYPUT | bool | declared | 103 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET/PROPERTYPUT | String | declared | 110 | Implemented | Reviewed | Live Tested | |
| Width | PROPERTYGET/PROPERTYPUT | f64 | declared | 122 | Implemented | Reviewed | Live Tested | |
| Height | PROPERTYGET/PROPERTYPUT | f64 | declared | 123 | Implemented | Reviewed | Live Tested | |
| Top | PROPERTYGET/PROPERTYPUT | f64 | declared | 126 | Implemented | Reviewed | Live Tested | |
| Left | PROPERTYGET/PROPERTYPUT | f64 | declared | 127 | Implemented | Reviewed | Live Tested | |
| Border | PROPERTYGET | Border | declared | 128 | Metadata Only | Reviewed | Not Tested | |
| Interior | PROPERTYGET | Interior | declared | 129 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Locked | PROPERTYGET/PROPERTYPUT | bool | declared | 269 | Metadata Only | Reviewed | Not Tested | |
| Index | PROPERTYGET | i32 | declared | 486 | Metadata Only | Reviewed | Not Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | bool | declared | 558 | Implemented | Reviewed | Live Tested | |
| OnAction | PROPERTYGET/PROPERTYPUT | String | declared | 596 | Metadata Only | Reviewed | Not Tested | |
| Enabled | PROPERTYGET/PROPERTYPUT | bool | declared | 600 | Metadata Only | Reviewed | Not Tested | |
| BottomRightCell | PROPERTYGET | Range | declared | 615 | Metadata Only | Reviewed | Not Tested | |
| Placement | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 617 | Implemented | Reviewed | Live Tested | |
| PrintObject | PROPERTYGET/PROPERTYPUT | bool | declared | 618 | Metadata Only | Reviewed | Not Tested | |
| RoundedCorners | PROPERTYGET/PROPERTYPUT | bool | declared | 619 | Metadata Only | Reviewed | Not Tested | |
| TopLeftCell | PROPERTYGET | Range | declared | 620 | Metadata Only | Reviewed | Not Tested | |
| ZOrder | PROPERTYGET | i32 | declared | 622 | Metadata Only | Reviewed | Not Tested | |
| ShapeRange | PROPERTYGET | ShapeRange | declared | 1528 | Metadata Only | Reviewed | Not Tested | |
| ProtectChartObject | PROPERTYGET/PROPERTYPUT | bool | declared | 1529 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Delete | AutomationValue | 0 | declared | 117 | Implemented | Reviewed | Live Tested | |
| CopyPicture | AutomationValue | 2 | declared | 213 | Implemented | Reviewed | Live Tested | |
| Select | AutomationValue | 1 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| Activate | AutomationValue | 0 | declared | 304 | Implemented | Reviewed | Live Tested | |
| Copy | AutomationValue | 0 | declared | 551 | Implemented | Reviewed | Live Tested | |
| Cut | AutomationValue | 0 | declared | 565 | Metadata Only | Reviewed | Not Tested | |
| BringToFront | AutomationValue | 0 | declared | 602 | Metadata Only | Reviewed | Not Tested | |
| SendToBack | AutomationValue | 0 | declared | 605 | Metadata Only | Reviewed | Not Tested | |
| Duplicate | Object | 0 | declared | 1039 | Metadata Only | Reviewed | Not Tested | |
| _Copy | AutomationValue | 0 | declared | 2609 | Metadata Only | Reviewed | Not Tested | |
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
