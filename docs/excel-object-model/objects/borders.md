# Borders

## Summary

The enum-keyed, apartment-bound collection of Excel Border objects for a Range.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Borders` |
| GUID | `{00020855-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Borders` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Capabilities

No capability metadata is recorded for this surface.


## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `_Default` | `excel.border` | Metadata Only |
| `Item` | `excel.border` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _NewEnum | PROPERTYGET | Unknown | declared | -4 | Implemented | Reviewed | Live Tested | |
| _Default | PROPERTYGET | Border | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| Value | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 6 | Metadata Only | Reviewed | Not Tested | |
| ColorIndex | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 97 | Implemented | Reviewed | Live Tested | |
| Color | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 99 | Implemented | Reviewed | Live Tested | |
| Count | PROPERTYGET | i32 | declared | 118 | Implemented | Reviewed | Live Tested | |
| LineStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 119 | Implemented | Reviewed | Live Tested | |
| Weight | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 120 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Item | PROPERTYGET | Border | declared | 170 | Implemented | Reviewed | Live Tested | |
| ThemeColor | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2365 | Metadata Only | Reviewed | Not Tested | |
| TintAndShade | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2366 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
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
