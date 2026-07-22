# Sort

## Summary

Excel's persistent sort configuration for a range or table.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Sort` |
| GUID | `{000244ab-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Sort` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Capabilities

### Structured data

| Capability | Available |
|---|---|
| `filter` | false |
| `remove_duplicates` | false |
| `sort` | true |
| `structural_editing` | false |
| `tables` | false |
| `validation` | false |



## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Rng` | `excel.range` | Metadata Only |
| `SortFields` | `excel.sortfields` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Orientation | PROPERTYGET/PROPERTYPUT | XlSortOrientation | declared | 134 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| MatchCase | PROPERTYGET/PROPERTYPUT | bool | declared | 426 | Implemented | Reviewed | Live Tested | |
| Header | PROPERTYGET/PROPERTYPUT | XlYesNoGuess | declared | 895 | Implemented | Reviewed | Live Tested | |
| SortMethod | PROPERTYGET/PROPERTYPUT | XlSortMethod | declared | 897 | Metadata Only | Reviewed | Not Tested | |
| Rng | PROPERTYGET | Range | declared | 2748 | Metadata Only | Reviewed | Not Tested | |
| SortFields | PROPERTYGET | SortFields | declared | 2749 | Implemented | Reviewed | Live Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Apply | Unknown | 0 | declared | 1675 | Implemented | Reviewed | Live Tested | |
| SetRange | Unknown | 1 | declared | 2750 | Implemented | Reviewed | Live Tested | |
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
