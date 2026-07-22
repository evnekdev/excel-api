# Validation

## Summary

Excel data-validation state associated with a Range.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Validation` |
| GUID | `{0002442f-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Validation` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Capabilities

### Structured data

| Capability | Available |
|---|---|
| `filter` | false |
| `remove_duplicates` | false |
| `sort` | false |
| `structural_editing` | false |
| `tables` | false |
| `validation` | true |



## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Value | PROPERTYGET | bool | declared | 6 | Implemented | Reviewed | Live Tested | |
| Type | PROPERTYGET | i32 | declared | 108 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Operator | PROPERTYGET | i32 | declared | 797 | Implemented | Reviewed | Live Tested | |
| Formula1 | PROPERTYGET | String | declared | 1579 | Implemented | Reviewed | Live Tested | |
| Formula2 | PROPERTYGET | String | declared | 1580 | Implemented | Reviewed | Live Tested | |
| AlertStyle | PROPERTYGET | i32 | declared | 1605 | Implemented | Reviewed | Live Tested | |
| IgnoreBlank | PROPERTYGET/PROPERTYPUT | bool | declared | 1606 | Implemented | Reviewed | Live Tested | |
| IMEMode | PROPERTYGET/PROPERTYPUT | i32 | declared | 1607 | Metadata Only | Reviewed | Not Tested | |
| InCellDropdown | PROPERTYGET/PROPERTYPUT | bool | declared | 1608 | Implemented | Reviewed | Live Tested | |
| ErrorMessage | PROPERTYGET/PROPERTYPUT | String | declared | 1609 | Implemented | Reviewed | Live Tested | |
| ErrorTitle | PROPERTYGET/PROPERTYPUT | String | declared | 1610 | Implemented | Reviewed | Live Tested | |
| InputMessage | PROPERTYGET/PROPERTYPUT | String | declared | 1611 | Implemented | Reviewed | Live Tested | |
| InputTitle | PROPERTYGET/PROPERTYPUT | String | declared | 1612 | Implemented | Reviewed | Live Tested | |
| ShowError | PROPERTYGET/PROPERTYPUT | bool | declared | 1613 | Implemented | Reviewed | Live Tested | |
| ShowInput | PROPERTYGET/PROPERTYPUT | bool | declared | 1614 | Implemented | Reviewed | Live Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Delete | Unknown | 0 | declared | 117 | Implemented | Reviewed | Live Tested | |
| Add | Unknown | 5 | declared | 181 | Implemented | Reviewed | Live Tested | |
| Modify | Unknown | 5 | declared | 1581 | Metadata Only | Reviewed | Not Tested | |
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
