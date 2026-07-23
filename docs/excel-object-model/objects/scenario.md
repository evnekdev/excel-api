# Scenario

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Scenario` |
| GUID | `{00020897-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | secondary-public |
| Roadmap class | priority-inventory |
| Type flags | 4096 |
| Crate type | `excel_com::Scenario` |
| Implementation | Metadata Only |
| Documentation | Reviewed |
| Tests | Not Tested |

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
| `open_text` | false |
| `scenarios` | true |
| `subtotal` | false |
| `text_export` | false |
| `text_to_columns` | false |



## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `ChangingCells` | `excel.range` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Name | PROPERTYGET/PROPERTYPUT | String | declared | 110 | Implemented | Reviewed | Blocked | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Values | PROPERTYGET | AutomationValue | declared | 164 | Implemented | Reviewed | Blocked | |
| Hidden | PROPERTYGET/PROPERTYPUT | bool | declared | 268 | Metadata Only | Reviewed | Not Tested | |
| Locked | PROPERTYGET/PROPERTYPUT | bool | declared | 269 | Metadata Only | Reviewed | Not Tested | |
| Index | PROPERTYGET | i32 | declared | 486 | Metadata Only | Reviewed | Not Tested | |
| Comment | PROPERTYGET/PROPERTYPUT | String | declared | 910 | Implemented | Reviewed | Blocked | |
| ChangingCells | PROPERTYGET | Range | declared | 911 | Implemented | Reviewed | Blocked | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Delete | AutomationValue | 0 | declared | 117 | Implemented | Reviewed | Blocked | |
| Show | AutomationValue | 0 | declared | 496 | Implemented | Reviewed | Blocked | |
| ChangeScenario | AutomationValue | 2 | declared | 912 | Implemented | Reviewed | Blocked | |
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
