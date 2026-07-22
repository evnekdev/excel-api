# ListRows

## Summary

The typed collection of data rows belonging to one Excel table.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `ListRows` |
| GUID | `{00024474-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::ListRows` |
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
| `Add` | `excel.listrow` | Implemented |
| `_Add` | `excel.listrow` | Metadata Only |
| `Application` | `excel.application` | Metadata Only |
| `_Default` | `excel.listrow` | Metadata Only |
| `Item` | `excel.listrow` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _NewEnum | PROPERTYGET | Unknown | declared | -4 | Implemented | Reviewed | Live Tested | |
| _Default | PROPERTYGET | ListRow | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| Count | PROPERTYGET | i32 | declared | 118 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Item | PROPERTYGET | ListRow | declared | 170 | Implemented | Reviewed | Live Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Add | ListRow | 2 | declared | 181 | Implemented | Reviewed | Live Tested | |
| _Add | ListRow | 1 | declared | 2085 | Metadata Only | Reviewed | Not Tested | |
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
