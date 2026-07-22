# Workbooks

## Summary

The collection through which an Application exposes open workbooks. The initial crate supports counting and creating a workbook, without a general collection abstraction.

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
| Surface class | implemented-wrapper |
| Crate type | `excel_com::Workbooks` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Add` | `excel.workbook` | Implemented |
| `Application` | `excel.application` | Metadata Only |
| `_Default` | `excel.workbook` | Metadata Only |
| `Item` | `excel.workbook` | Metadata Only |
| `_Open` | `excel.workbook` | Metadata Only |
| `Open` | `excel.workbook` | Metadata Only |
| `OpenDatabase` | `excel.workbook` | Metadata Only |
| `_OpenXML` | `excel.workbook` | Metadata Only |
| `OpenXML` | `excel.workbook` | Metadata Only |

## Properties

| Property | Access | Type | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---:|---|---|---|---|
| _NewEnum | PROPERTYGET | Unknown | -4 | Metadata Only | Reviewed | Not Tested | |
| _Default | PROPERTYGET | Workbook | 0 | Metadata Only | Reviewed | Not Tested | |
| Count | PROPERTYGET | i32 | 118 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | 150 | Metadata Only | Reviewed | Not Tested | |
| Item | PROPERTYGET | Workbook | 170 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---:|---|---|---|---|
| Add | Workbook | 1 | 181 | Implemented | Reviewed | Live Tested | |
| Close | Unknown | 0 | 277 | Metadata Only | Reviewed | Not Tested | |
| _Open | Workbook | 13 | 682 | Metadata Only | Reviewed | Not Tested | |
| __OpenText | Unknown | 14 | 683 | Metadata Only | Reviewed | Not Tested | |
| _OpenText | Unknown | 16 | 1773 | Metadata Only | Reviewed | Not Tested | |
| Open | Workbook | 15 | 1923 | Metadata Only | Reviewed | Not Tested | |
| OpenText | Unknown | 18 | 1924 | Metadata Only | Reviewed | Not Tested | |
| OpenDatabase | Workbook | 5 | 2067 | Metadata Only | Reviewed | Not Tested | |
| CheckOut | Unknown | 1 | 2069 | Metadata Only | Reviewed | Not Tested | |
| CanCheckOut | bool | 1 | 2070 | Metadata Only | Reviewed | Not Tested | |
| _OpenXML | Workbook | 2 | 2071 | Metadata Only | Reviewed | Not Tested | |
| OpenXML | Workbook | 3 | 2280 | Metadata Only | Reviewed | Not Tested | |
| QueryInterface | Unknown | 2 | 1610612736 | Metadata Only | Reviewed | Not Tested | |
| AddRef | Unknown | 0 | 1610612737 | Metadata Only | Reviewed | Not Tested | |
| Release | Unknown | 0 | 1610612738 | Metadata Only | Reviewed | Not Tested | |
| GetTypeInfoCount | Unknown | 1 | 1610678272 | Metadata Only | Reviewed | Not Tested | |
| GetTypeInfo | Unknown | 3 | 1610678273 | Metadata Only | Reviewed | Not Tested | |
| GetIDsOfNames | Unknown | 5 | 1610678274 | Metadata Only | Reviewed | Not Tested | |
| Invoke | Unknown | 8 | 1610678275 | Metadata Only | Reviewed | Not Tested | |

## Events

| Event | Arguments | DISPID | Implementation | Docs | Tests |
|---|---:|---:|---|---|---|
| -- | -- | -- | Not started | Generated | Not tested |

## Unsupported or deferred behaviour

See the global unsupported index for unimplemented object-model areas.
<!-- END GENERATED MEMBERS -->
