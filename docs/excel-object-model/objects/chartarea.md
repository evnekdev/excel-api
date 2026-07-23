# ChartArea

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `ChartArea` |
| GUID | `{000208cc-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::ChartArea` |
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

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Shadow | PROPERTYGET/PROPERTYPUT | bool | declared | 103 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET | String | declared | 110 | Metadata Only | Reviewed | Not Tested | |
| Format | PROPERTYGET | ChartFormat | declared | 116 | Implemented | Reviewed | Live Tested | |
| Width | PROPERTYGET/PROPERTYPUT | f64 | declared | 122 | Implemented | Reviewed | Live Tested | |
| Height | PROPERTYGET/PROPERTYPUT | f64 | declared | 123 | Implemented | Reviewed | Live Tested | |
| Top | PROPERTYGET/PROPERTYPUT | f64 | declared | 126 | Implemented | Reviewed | Live Tested | |
| Left | PROPERTYGET/PROPERTYPUT | f64 | declared | 127 | Implemented | Reviewed | Live Tested | |
| Border | PROPERTYGET | Border | declared | 128 | Metadata Only | Reviewed | Not Tested | |
| Interior | PROPERTYGET | Interior | declared | 129 | Metadata Only | Reviewed | Not Tested | |
| Font | PROPERTYGET | Font | declared | 146 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| RoundedCorners | PROPERTYGET/PROPERTYPUT | bool | declared | 619 | Metadata Only | Reviewed | Not Tested | |
| AutoScaleFont | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1525 | Metadata Only | Reviewed | Not Tested | |
| Fill | PROPERTYGET | ChartFillFormat | declared | 1663 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Clear | AutomationValue | 0 | declared | 111 | Metadata Only | Reviewed | Not Tested | |
| ClearFormats | AutomationValue | 0 | declared | 112 | Metadata Only | Reviewed | Not Tested | |
| _ClearContents | AutomationValue | 0 | declared | 113 | Metadata Only | Reviewed | Not Tested | |
| Select | AutomationValue | 0 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| Copy | AutomationValue | 0 | declared | 551 | Metadata Only | Reviewed | Not Tested | |
| ClearContents | AutomationValue | 0 | declared | 3413 | Metadata Only | Reviewed | Not Tested | |
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
