# ListObject

## Summary

An apartment-bound Excel table wrapper with bounded table, row, column, filter, and sort operations.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `ListObject` |
| GUID | `{00024471-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::ListObject` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Capabilities

### Structured data

| Capability | Available |
|---|---|
| `filter` | true |
| `remove_duplicates` | false |
| `sort` | true |
| `structural_editing` | true |
| `tables` | true |
| `validation` | false |



## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `_AutoFilter` | `excel.autofilter` | Metadata Only |
| `AutoFilter` | `excel.autofilter` | Implemented |
| `DataBodyRange` | `excel.range` | Implemented |
| `HeaderRowRange` | `excel.range` | Implemented |
| `InsertRowRange` | `excel.range` | Implemented |
| `ListColumns` | `excel.listcolumns` | Implemented |
| `ListRows` | `excel.listrows` | Implemented |
| `Range` | `excel.range` | Implemented |
| `_Sort` | `excel.sort` | Metadata Only |
| `Sort` | `excel.sort` | Implemented |
| `TotalsRowRange` | `excel.range` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _Default | PROPERTYGET | String | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET/PROPERTYPUT | String | declared | 110 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Range | PROPERTYGET | Range | declared | 197 | Implemented | Reviewed | Live Tested | |
| Summary | PROPERTYGET/PROPERTYPUT | String | declared | 273 | Metadata Only | Reviewed | Not Tested | |
| SourceType | PROPERTYGET | XlListObjectSourceType | declared | 685 | Metadata Only | Reviewed | Not Tested | |
| DataBodyRange | PROPERTYGET | Range | declared | 705 | Implemented | Reviewed | Live Tested | |
| _AutoFilter | PROPERTYGET | AutoFilter | declared | 793 | Metadata Only | Reviewed | Not Tested | |
| _Sort | PROPERTYGET | Sort | declared | 880 | Metadata Only | Reviewed | Not Tested | |
| Comment | PROPERTYGET/PROPERTYPUT | String | declared | 910 | Metadata Only | Reviewed | Not Tested | |
| QueryTable | PROPERTYGET | QueryTable | declared | 1386 | Metadata Only | Reviewed | Not Tested | |
| TableStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1504 | Implemented | Reviewed | Live Tested | |
| DisplayRightToLeft | PROPERTYGET | bool | declared | 1774 | Metadata Only | Reviewed | Not Tested | |
| AlternativeText | PROPERTYGET/PROPERTYPUT | String | declared | 1891 | Metadata Only | Reviewed | Not Tested | |
| XmlMap | PROPERTYGET | XmlMap | declared | 2253 | Metadata Only | Reviewed | Not Tested | |
| Active | PROPERTYGET | bool | declared | 2312 | Metadata Only | Reviewed | Not Tested | |
| HeaderRowRange | PROPERTYGET | Range | declared | 2313 | Implemented | Reviewed | Live Tested | |
| InsertRowRange | PROPERTYGET | Range | declared | 2314 | Implemented | Reviewed | Live Tested | |
| ListColumns | PROPERTYGET | ListColumns | declared | 2315 | Implemented | Reviewed | Live Tested | |
| ListRows | PROPERTYGET | ListRows | declared | 2316 | Implemented | Reviewed | Live Tested | |
| ShowAutoFilter | PROPERTYGET/PROPERTYPUT | bool | declared | 2317 | Implemented | Reviewed | Live Tested | |
| ShowTotals | PROPERTYGET/PROPERTYPUT | bool | declared | 2318 | Implemented | Reviewed | Live Tested | |
| TotalsRowRange | PROPERTYGET | Range | declared | 2319 | Implemented | Reviewed | Live Tested | |
| SharePointURL | PROPERTYGET | String | declared | 2320 | Metadata Only | Reviewed | Not Tested | |
| ShowTableStyleLastColumn | PROPERTYGET/PROPERTYPUT | bool | declared | 2563 | Metadata Only | Reviewed | Not Tested | |
| ShowTableStyleRowStripes | PROPERTYGET/PROPERTYPUT | bool | declared | 2564 | Metadata Only | Reviewed | Not Tested | |
| ShowTableStyleColumnStripes | PROPERTYGET/PROPERTYPUT | bool | declared | 2565 | Metadata Only | Reviewed | Not Tested | |
| DisplayName | PROPERTYGET/PROPERTYPUT | String | declared | 2677 | Implemented | Reviewed | Live Tested | |
| ShowHeaders | PROPERTYGET/PROPERTYPUT | bool | declared | 2678 | Implemented | Reviewed | Live Tested | |
| ShowTableStyleFirstColumn | PROPERTYGET/PROPERTYPUT | bool | declared | 2679 | Metadata Only | Reviewed | Not Tested | |
| Slicers | PROPERTYGET | Slicers | declared | 2881 | Metadata Only | Reviewed | Not Tested | |
| TableObject | PROPERTYGET | TableObject | declared | 3095 | Metadata Only | Reviewed | Not Tested | |
| ShowAutoFilterDropDown | PROPERTYGET/PROPERTYPUT | bool | declared | 3096 | Metadata Only | Reviewed | Not Tested | |
| Sort | PROPERTYGET | Sort | declared | 3288 | Implemented | Reviewed | Live Tested | |
| AutoFilter | PROPERTYGET | AutoFilter | declared | 3289 | Implemented | Reviewed | Live Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Delete | Unknown | 0 | declared | 117 | Implemented | Reviewed | Live Tested | |
| Resize | Unknown | 1 | declared | 256 | Implemented | Reviewed | Live Tested | |
| Refresh | Unknown | 0 | declared | 1417 | Metadata Only | Reviewed | Not Tested | |
| Publish | String | 2 | declared | 1895 | Metadata Only | Reviewed | Not Tested | |
| Unlink | Unknown | 0 | declared | 2308 | Metadata Only | Reviewed | Not Tested | |
| Unlist | Unknown | 0 | declared | 2309 | Implemented | Reviewed | Live Tested | |
| UpdateChanges | Unknown | 1 | declared | 2310 | Metadata Only | Reviewed | Not Tested | |
| ExportToVisio | Unknown | 0 | declared | 2680 | Metadata Only | Reviewed | Not Tested | |
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
