# Trendline

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Trendline` |
| GUID | `{000208be-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Trendline` |
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
| Type | PROPERTYGET/PROPERTYPUT | XlTrendlineType | declared | 108 | Implemented | Reviewed | Live Tested | |
| Name | PROPERTYGET/PROPERTYPUT | String | declared | 110 | Implemented | Reviewed | Live Tested | |
| Format | PROPERTYGET | ChartFormat | declared | 116 | Implemented | Reviewed | Live Tested | |
| Border | PROPERTYGET | Border | declared | 128 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| DataLabel | PROPERTYGET | DataLabel | declared | 158 | Implemented | Reviewed | Live Tested | |
| Period | PROPERTYGET/PROPERTYPUT | i32 | declared | 184 | Implemented | Reviewed | Live Tested | |
| Backward | PROPERTYGET/PROPERTYPUT | i32 | declared | 185 | Implemented | Reviewed | Live Tested | |
| Intercept | PROPERTYGET/PROPERTYPUT | f64 | declared | 186 | Implemented | Reviewed | Live Tested | |
| InterceptIsAuto | PROPERTYGET/PROPERTYPUT | bool | declared | 187 | Metadata Only | Reviewed | Not Tested | |
| NameIsAuto | PROPERTYGET/PROPERTYPUT | bool | declared | 188 | Metadata Only | Reviewed | Not Tested | |
| DisplayRSquared | PROPERTYGET/PROPERTYPUT | bool | declared | 189 | Implemented | Reviewed | Live Tested | |
| DisplayEquation | PROPERTYGET/PROPERTYPUT | bool | declared | 190 | Implemented | Reviewed | Live Tested | |
| Forward | PROPERTYGET/PROPERTYPUT | i32 | declared | 191 | Implemented | Reviewed | Live Tested | |
| Order | PROPERTYGET/PROPERTYPUT | i32 | declared | 192 | Implemented | Reviewed | Live Tested | |
| Index | PROPERTYGET | i32 | declared | 486 | Metadata Only | Reviewed | Not Tested | |
| Backward2 | PROPERTYGET/PROPERTYPUT | f64 | declared | 2650 | Metadata Only | Reviewed | Not Tested | |
| Forward2 | PROPERTYGET/PROPERTYPUT | f64 | declared | 2651 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| ClearFormats | AutomationValue | 0 | declared | 112 | Metadata Only | Reviewed | Not Tested | |
| Delete | AutomationValue | 0 | declared | 117 | Implemented | Reviewed | Live Tested | |
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
