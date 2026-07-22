# ListRow

## Summary

One apartment-bound data row within an Excel table.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `ListRow` |
| GUID | `{00024475-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::ListRow` |
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
| `Range` | `excel.range` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Range | PROPERTYGET | Range | declared | 197 | Implemented | Reviewed | Live Tested | |
| Index | PROPERTYGET | i32 | declared | 486 | Implemented | Reviewed | Live Tested | |
| InvalidData | PROPERTYGET | bool | declared | 2324 | Metadata Only | Reviewed | Not Tested | |

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
