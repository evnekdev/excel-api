# Worksheets

## Summary

The workbook worksheet collection. It is structurally inventoried but has no production wrapper in this initial crate.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Worksheets` |
| GUID | `{000208b1-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | implemented-wrapper |
| Crate type | `excel_com::Worksheets` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Item` | `excel.worksheet` | Implemented |

## Properties

| Property | Access | Type | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---:|---|---|---|---|
| _NewEnum | PROPERTYGET | Unknown | -4 | Metadata Only | Reviewed | Not Tested | |
| _Default | PROPERTYGET | Object | 0 | Metadata Only | Reviewed | Not Tested | |
| Count | PROPERTYGET | i32 | 118 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | 150 | Metadata Only | Reviewed | Not Tested | |
| Item | PROPERTYGET | Object | 170 | Implemented | Reviewed | Live Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | AutomationValue | 558 | Metadata Only | Reviewed | Not Tested | |
| HPageBreaks | PROPERTYGET | HPageBreaks | 1418 | Metadata Only | Reviewed | Not Tested | |
| VPageBreaks | PROPERTYGET | VPageBreaks | 1419 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---:|---|---|---|---|
| Delete | Unknown | 0 | 117 | Metadata Only | Reviewed | Not Tested | |
| Add | Object | 4 | 181 | Implemented | Reviewed | Live Tested | |
| Select | Unknown | 1 | 235 | Metadata Only | Reviewed | Not Tested | |
| PrintPreview | Unknown | 1 | 281 | Metadata Only | Reviewed | Not Tested | |
| FillAcrossSheets | Unknown | 2 | 469 | Metadata Only | Reviewed | Not Tested | |
| Copy | Unknown | 2 | 551 | Metadata Only | Reviewed | Not Tested | |
| Move | Unknown | 2 | 637 | Metadata Only | Reviewed | Not Tested | |
| __PrintOut | Unknown | 7 | 905 | Metadata Only | Reviewed | Not Tested | |
| _PrintOut | Unknown | 8 | 1772 | Metadata Only | Reviewed | Not Tested | |
| PrintOut | Unknown | 9 | 2361 | Metadata Only | Reviewed | Not Tested | |
| Add2 | Object | 4 | 3054 | Metadata Only | Reviewed | Not Tested | |
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
