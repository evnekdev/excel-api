# SparklineGroup

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `SparklineGroup` |
| GUID | `{000244b7-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::SparklineGroup` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Blocked |

## Capabilities

No capability metadata is recorded for this surface.


## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Location` | `excel.range` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _NewEnum | PROPERTYGET | Unknown | declared | -4 | Metadata Only | Reviewed | Not Tested | |
| Axes | PROPERTYGET | SparkAxes | declared | 23 | Metadata Only | Reviewed | Not Tested | |
| Points | PROPERTYGET | SparkPoints | declared | 70 | Metadata Only | Reviewed | Not Tested | |
| DisplayBlanksAs | PROPERTYGET/PROPERTYPUT | XlDisplayBlanksAs | declared | 93 | Metadata Only | Reviewed | Not Tested | |
| Type | PROPERTYGET/PROPERTYPUT | XlSparkType | declared | 108 | Implemented | Reviewed | Live Tested | |
| Count | PROPERTYGET | i32 | declared | 118 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Item | PROPERTYGET | Sparkline | declared | 170 | Metadata Only | Reviewed | Not Tested | |
| PlotBy | PROPERTYGET/PROPERTYPUT | XlSparklineRowCol | declared | 202 | Metadata Only | Reviewed | Not Tested | |
| SourceData | PROPERTYGET/PROPERTYPUT | String | declared | 686 | Implemented | Reviewed | Live Tested | |
| Location | PROPERTYGET/PROPERTYPUTREF | Range | declared | 1397 | Implemented | Reviewed | Live Tested | |
| DateRange | PROPERTYGET/PROPERTYPUT | String | declared | 2948 | Metadata Only | Reviewed | Not Tested | |
| SeriesColor | PROPERTYGET | FormatColor | declared | 2952 | Metadata Only | Reviewed | Not Tested | |
| DisplayHidden | PROPERTYGET/PROPERTYPUT | bool | declared | 2953 | Metadata Only | Reviewed | Not Tested | |
| LineWeight | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2954 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Delete | Unknown | 0 | declared | 117 | Implemented | Reviewed | Live Tested | |
| Modify | Unknown | 2 | declared | 1581 | Metadata Only | Reviewed | Not Tested | |
| ModifyLocation | Unknown | 1 | declared | 2949 | Metadata Only | Reviewed | Not Tested | |
| ModifySourceData | Unknown | 1 | declared | 2950 | Metadata Only | Reviewed | Not Tested | |
| ModifyDateRange | Unknown | 1 | declared | 2951 | Metadata Only | Reviewed | Not Tested | |
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
