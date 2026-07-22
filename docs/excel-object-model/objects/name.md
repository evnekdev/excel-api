# Name

## Summary

One Excel defined name, which may resolve to a Range, scalar, formula, or invalid reference. `RefersToRange` is therefore represented by a fallible `Name::range` API.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Name` |
| GUID | `{000208b9-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Name` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `RefersToRange` | `excel.range` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _Default | PROPERTYGET | String | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| Value | PROPERTYGET/PROPERTYPUT | String | declared | 6 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET/PROPERTYPUT | String | declared | 110 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Index | PROPERTYGET | i32 | declared | 486 | Metadata Only | Reviewed | Not Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | bool | declared | 558 | Implemented | Reviewed | Live Tested | |
| ShortcutKey | PROPERTYGET/PROPERTYPUT | String | declared | 597 | Metadata Only | Reviewed | Not Tested | |
| Comment | PROPERTYGET/PROPERTYPUT | String | declared | 910 | Metadata Only | Reviewed | Not Tested | |
| Category | PROPERTYGET/PROPERTYPUT | String | declared | 934 | Metadata Only | Reviewed | Not Tested | |
| CategoryLocal | PROPERTYGET/PROPERTYPUT | String | declared | 935 | Metadata Only | Reviewed | Not Tested | |
| MacroType | PROPERTYGET/PROPERTYPUT | XlXLMMacroType | declared | 936 | Metadata Only | Reviewed | Not Tested | |
| NameLocal | PROPERTYGET/PROPERTYPUT | String | declared | 937 | Metadata Only | Reviewed | Not Tested | |
| RefersTo | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 938 | Implemented | Reviewed | Live Tested | |
| RefersToLocal | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 939 | Metadata Only | Reviewed | Not Tested | |
| RefersToR1C1 | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 940 | Implemented | Reviewed | Live Tested | |
| RefersToR1C1Local | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 941 | Metadata Only | Reviewed | Not Tested | |
| RefersToRange | PROPERTYGET | Range | declared | 1160 | Implemented | Reviewed | Live Tested | |
| WorkbookParameter | PROPERTYGET/PROPERTYPUT | bool | declared | 2607 | Metadata Only | Reviewed | Not Tested | |
| ValidWorkbookParameter | PROPERTYGET | bool | declared | 2608 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Delete | Unknown | 0 | declared | 117 | Implemented | Reviewed | Live Tested | |
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
