# Workbooks

## Summary

The typed collection through which an Application exposes open workbooks. The bounded crate supports Count, Item lookup, creation, file operations, and fallible `_NewEnum` iteration without exposing a generic public collection API.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Workbooks` |
| GUID | `{000208db-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4288 |
| Crate type | `excel_com::Workbooks` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Capabilities

### Data utility

| Capability | Available |
|---|---|
| `advanced_filter` | false |
| `autofill` | false |
| `consolidate` | false |
| `data_tables` | false |
| `external_links` | false |
| `fill` | false |
| `flash_fill` | false |
| `goal_seek` | false |
| `open_text` | true |
| `scenarios` | false |
| `subtotal` | false |
| `text_export` | false |
| `text_to_columns` | false |



## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Add` | `excel.workbook` | Implemented |
| `Application` | `excel.application` | Implemented |
| `_Default` | `excel.workbook` | Metadata Only |
| `Item` | `excel.workbook` | Implemented |
| `_Open` | `excel.workbook` | Metadata Only |
| `Open` | `excel.workbook` | Implemented |
| `OpenDatabase` | `excel.workbook` | Metadata Only |
| `_OpenXML` | `excel.workbook` | Metadata Only |
| `OpenXML` | `excel.workbook` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _NewEnum | PROPERTYGET | Unknown | declared | -4 | Implemented | Reviewed | Live Tested | |
| _Default | PROPERTYGET | Workbook | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| Count | PROPERTYGET | i32 | declared | 118 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Implemented | Reviewed | Blocked | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Item | PROPERTYGET | Workbook | declared | 170 | Implemented | Reviewed | Live Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Add | Workbook | 1 | declared | 181 | Implemented | Reviewed | Live Tested | |
| Close | Unknown | 0 | declared | 277 | Metadata Only | Reviewed | Not Tested | |
| _Open | Workbook | 13 | declared | 682 | Metadata Only | Reviewed | Not Tested | |
| __OpenText | Unknown | 14 | declared | 683 | Metadata Only | Reviewed | Not Tested | |
| _OpenText | Unknown | 16 | declared | 1773 | Metadata Only | Reviewed | Not Tested | |
| Open | Workbook | 15 | declared | 1923 | Implemented | Reviewed | Live Tested | |
| OpenText | Unknown | 18 | declared | 1924 | Implemented | Reviewed | Blocked | |
| OpenDatabase | Workbook | 5 | declared | 2067 | Metadata Only | Reviewed | Not Tested | |
| CheckOut | Unknown | 1 | declared | 2069 | Metadata Only | Reviewed | Not Tested | |
| CanCheckOut | bool | 1 | declared | 2070 | Metadata Only | Reviewed | Not Tested | |
| _OpenXML | Workbook | 2 | declared | 2071 | Metadata Only | Reviewed | Not Tested | |
| OpenXML | Workbook | 3 | declared | 2280 | Metadata Only | Reviewed | Not Tested | |
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
