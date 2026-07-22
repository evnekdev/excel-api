# SortField

## Summary

One configured Excel persistent sort field.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `SortField` |
| GUID | `{000244a9-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::SortField` |
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
| `Key` | `excel.range` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Key | PROPERTYGET | Range | declared | 155 | Metadata Only | Reviewed | Not Tested | |
| Order | PROPERTYGET/PROPERTYPUT | XlSortOrder | declared | 192 | Metadata Only | Reviewed | Not Tested | |
| Priority | PROPERTYGET/PROPERTYPUT | i32 | declared | 985 | Metadata Only | Reviewed | Not Tested | |
| SortOn | PROPERTYGET/PROPERTYPUT | XlSortOn | declared | 2741 | Metadata Only | Reviewed | Not Tested | |
| SortOnValue | PROPERTYGET | Object | declared | 2742 | Metadata Only | Reviewed | Not Tested | |
| CustomOrder | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2743 | Metadata Only | Reviewed | Not Tested | |
| DataOption | PROPERTYGET/PROPERTYPUT | XlSortDataOption | declared | 2744 | Metadata Only | Reviewed | Not Tested | |
| SubField | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 3328 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Delete | Unknown | 0 | declared | 117 | Metadata Only | Reviewed | Not Tested | |
| ModifyKey | Unknown | 1 | declared | 2745 | Metadata Only | Reviewed | Not Tested | |
| SetIcon | Unknown | 1 | declared | 2746 | Metadata Only | Reviewed | Not Tested | |
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
