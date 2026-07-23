# ChartObjects

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `ChartObjects` |
| GUID | `{000208d0-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::ChartObjects` |
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
| `Item` | `excel.chartobject` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Shadow | PROPERTYGET/PROPERTYPUT | bool | declared | 103 | Metadata Only | Reviewed | Not Tested | |
| Count | PROPERTYGET | i32 | declared | 118 | Implemented | Reviewed | Live Tested | |
| Width | PROPERTYGET/PROPERTYPUT | f64 | declared | 122 | Metadata Only | Reviewed | Not Tested | |
| Height | PROPERTYGET/PROPERTYPUT | f64 | declared | 123 | Metadata Only | Reviewed | Not Tested | |
| Top | PROPERTYGET/PROPERTYPUT | f64 | declared | 126 | Metadata Only | Reviewed | Not Tested | |
| Left | PROPERTYGET/PROPERTYPUT | f64 | declared | 127 | Metadata Only | Reviewed | Not Tested | |
| Border | PROPERTYGET | Border | declared | 128 | Metadata Only | Reviewed | Not Tested | |
| Interior | PROPERTYGET | Interior | declared | 129 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Locked | PROPERTYGET/PROPERTYPUT | bool | declared | 269 | Metadata Only | Reviewed | Not Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | bool | declared | 558 | Metadata Only | Reviewed | Not Tested | |
| OnAction | PROPERTYGET/PROPERTYPUT | String | declared | 596 | Metadata Only | Reviewed | Not Tested | |
| Enabled | PROPERTYGET/PROPERTYPUT | bool | declared | 600 | Metadata Only | Reviewed | Not Tested | |
| Placement | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 617 | Metadata Only | Reviewed | Not Tested | |
| PrintObject | PROPERTYGET/PROPERTYPUT | bool | declared | 618 | Metadata Only | Reviewed | Not Tested | |
| RoundedCorners | PROPERTYGET/PROPERTYPUT | bool | declared | 619 | Metadata Only | Reviewed | Not Tested | |
| ShapeRange | PROPERTYGET | ShapeRange | declared | 1528 | Metadata Only | Reviewed | Not Tested | |
| ProtectChartObject | PROPERTYGET/PROPERTYPUT | bool | declared | 1529 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| _NewEnum | Unknown | 0 | declared | -4 | Implemented | Reviewed | Live Tested | |
| _Default | Object | 1 | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| Group | GroupObject | 0 | declared | 46 | Metadata Only | Reviewed | Not Tested | |
| Delete | AutomationValue | 0 | declared | 117 | Metadata Only | Reviewed | Not Tested | |
| Item | Object | 1 | declared | 170 | Implemented | Reviewed | Live Tested | |
| Add | ChartObject | 4 | declared | 181 | Implemented | Reviewed | Live Tested | |
| CopyPicture | AutomationValue | 2 | declared | 213 | Metadata Only | Reviewed | Not Tested | |
| Select | AutomationValue | 1 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| Copy | AutomationValue | 0 | declared | 551 | Metadata Only | Reviewed | Not Tested | |
| Cut | AutomationValue | 0 | declared | 565 | Metadata Only | Reviewed | Not Tested | |
| BringToFront | AutomationValue | 0 | declared | 602 | Metadata Only | Reviewed | Not Tested | |
| SendToBack | AutomationValue | 0 | declared | 605 | Metadata Only | Reviewed | Not Tested | |
| Duplicate | Object | 0 | declared | 1039 | Metadata Only | Reviewed | Not Tested | |
| _Copy | AutomationValue | 0 | declared | 2609 | Metadata Only | Reviewed | Not Tested | |
| _Dummy3 | Unknown | 0 | declared | 65539 | Metadata Only | Reviewed | Not Tested | |
| _Dummy12 | Unknown | 0 | declared | 65548 | Metadata Only | Reviewed | Not Tested | |
| _Dummy15 | Unknown | 0 | declared | 65551 | Metadata Only | Reviewed | Not Tested | |
| _Dummy22 | Unknown | 0 | declared | 65558 | Metadata Only | Reviewed | Not Tested | |
| _Dummy25 | Unknown | 0 | declared | 65561 | Metadata Only | Reviewed | Not Tested | |
| _Dummy27 | Unknown | 0 | declared | 65563 | Metadata Only | Reviewed | Not Tested | |
| _Dummy28 | Unknown | 0 | declared | 65564 | Metadata Only | Reviewed | Not Tested | |
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
