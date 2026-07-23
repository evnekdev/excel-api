# Series

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Series` |
| GUID | `{0002086b-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Series` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Blocked |

## Capabilities

No capability metadata is recorded for this surface.


## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Border` | `excel.border` | Metadata Only |
| `Interior` | `excel.interior` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| AxisGroup | PROPERTYGET/PROPERTYPUT | XlAxisGroup | declared | 47 | Implemented | Reviewed | Live Tested | |
| MarkerStyle | PROPERTYGET/PROPERTYPUT | XlMarkerStyle | declared | 72 | Implemented | Reviewed | Live Tested | |
| MarkerBackgroundColor | PROPERTYGET/PROPERTYPUT | i32 | declared | 73 | Implemented | Reviewed | Live Tested | |
| MarkerBackgroundColorIndex | PROPERTYGET/PROPERTYPUT | XlColorIndex | declared | 74 | Metadata Only | Reviewed | Not Tested | |
| MarkerForegroundColor | PROPERTYGET/PROPERTYPUT | i32 | declared | 75 | Implemented | Reviewed | Live Tested | |
| MarkerForegroundColorIndex | PROPERTYGET/PROPERTYPUT | XlColorIndex | declared | 76 | Metadata Only | Reviewed | Not Tested | |
| HasDataLabels | PROPERTYGET/PROPERTYPUT | bool | declared | 78 | Metadata Only | Reviewed | Not Tested | |
| Shadow | PROPERTYGET/PROPERTYPUT | bool | declared | 103 | Metadata Only | Reviewed | Not Tested | |
| Type | PROPERTYGET/PROPERTYPUT | i32 | declared | 108 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET/PROPERTYPUT | String | declared | 110 | Implemented | Reviewed | Live Tested | |
| Format | PROPERTYGET | ChartFormat | declared | 116 | Implemented | Reviewed | Live Tested | |
| Border | PROPERTYGET | Border | declared | 128 | Metadata Only | Reviewed | Not Tested | |
| Interior | PROPERTYGET | Interior | declared | 129 | Metadata Only | Reviewed | Not Tested | |
| InvertIfNegative | PROPERTYGET/PROPERTYPUT | bool | declared | 132 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| ErrorBars | PROPERTYGET | ErrorBars | declared | 159 | Implemented | Reviewed | Live Tested | |
| HasErrorBars | PROPERTYGET/PROPERTYPUT | bool | declared | 160 | Implemented | Reviewed | Live Tested | |
| PictureType | PROPERTYGET/PROPERTYPUT | XlChartPictureType | declared | 161 | Metadata Only | Reviewed | Not Tested | |
| PictureUnit | PROPERTYGET/PROPERTYPUT | i32 | declared | 162 | Metadata Only | Reviewed | Not Tested | |
| Smooth | PROPERTYGET/PROPERTYPUT | bool | declared | 163 | Implemented | Reviewed | Live Tested | |
| Values | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 164 | Implemented | Reviewed | Live Tested | |
| Explosion | PROPERTYGET/PROPERTYPUT | i32 | declared | 182 | Metadata Only | Reviewed | Not Tested | |
| PlotOrder | PROPERTYGET/PROPERTYPUT | i32 | declared | 228 | Implemented | Reviewed | Live Tested | |
| MarkerSize | PROPERTYGET/PROPERTYPUT | i32 | declared | 231 | Implemented | Reviewed | Live Tested | |
| Formula | PROPERTYGET/PROPERTYPUT | String | declared | 261 | Implemented | Reviewed | Live Tested | |
| FormulaLocal | PROPERTYGET/PROPERTYPUT | String | declared | 263 | Metadata Only | Reviewed | Not Tested | |
| FormulaR1C1 | PROPERTYGET/PROPERTYPUT | String | declared | 264 | Metadata Only | Reviewed | Not Tested | |
| FormulaR1C1Local | PROPERTYGET/PROPERTYPUT | String | declared | 265 | Metadata Only | Reviewed | Not Tested | |
| XValues | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1111 | Implemented | Reviewed | Live Tested | |
| HasLeaderLines | PROPERTYGET/PROPERTYPUT | bool | declared | 1394 | Implemented | Reviewed | Live Tested | |
| ChartType | PROPERTYGET/PROPERTYPUT | XlChartType | declared | 1400 | Implemented | Reviewed | Live Tested | |
| BarShape | PROPERTYGET/PROPERTYPUT | XlBarShape | declared | 1403 | Metadata Only | Reviewed | Not Tested | |
| ApplyPictToSides | PROPERTYGET/PROPERTYPUT | bool | declared | 1659 | Metadata Only | Reviewed | Not Tested | |
| ApplyPictToFront | PROPERTYGET/PROPERTYPUT | bool | declared | 1660 | Metadata Only | Reviewed | Not Tested | |
| ApplyPictToEnd | PROPERTYGET/PROPERTYPUT | bool | declared | 1661 | Metadata Only | Reviewed | Not Tested | |
| Fill | PROPERTYGET | ChartFillFormat | declared | 1663 | Metadata Only | Reviewed | Not Tested | |
| BubbleSizes | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1664 | Implemented | Reviewed | Live Tested | |
| Has3DEffect | PROPERTYGET/PROPERTYPUT | bool | declared | 1665 | Metadata Only | Reviewed | Not Tested | |
| LeaderLines | PROPERTYGET | LeaderLines | declared | 1666 | Metadata Only | Reviewed | Not Tested | |
| PictureUnit2 | PROPERTYGET/PROPERTYPUT | f64 | declared | 2649 | Metadata Only | Reviewed | Not Tested | |
| PlotColorIndex | PROPERTYGET | i32 | declared | 2915 | Metadata Only | Reviewed | Not Tested | |
| InvertColor | PROPERTYGET/PROPERTYPUT | i32 | declared | 2916 | Metadata Only | Reviewed | Not Tested | |
| InvertColorIndex | PROPERTYGET/PROPERTYPUT | i32 | declared | 2917 | Metadata Only | Reviewed | Not Tested | |
| IsFiltered | PROPERTYGET/PROPERTYPUT | bool | declared | 3083 | Metadata Only | Reviewed | Not Tested | |
| ParentDataLabelOption | PROPERTYGET/PROPERTYPUT | XlParentDataLabelOptions | declared | 3204 | Metadata Only | Reviewed | Not Tested | |
| QuartileCalculationInclusiveMedian | PROPERTYGET/PROPERTYPUT | bool | declared | 3205 | Metadata Only | Reviewed | Not Tested | |
| ValueSortOrder | PROPERTYGET/PROPERTYPUT | XlValueSortOrder | declared | 3229 | Metadata Only | Reviewed | Not Tested | |
| GeoProjectionType | PROPERTYGET/PROPERTYPUT | XlGeoProjectionType | declared | 3250 | Metadata Only | Reviewed | Not Tested | |
| GeoMappingLevel | PROPERTYGET/PROPERTYPUT | XlGeoMappingLevel | declared | 3251 | Metadata Only | Reviewed | Not Tested | |
| RegionLabelOption | PROPERTYGET/PROPERTYPUT | XlRegionLabelOptions | declared | 3252 | Metadata Only | Reviewed | Not Tested | |
| SeriesColorGradientStyle | PROPERTYGET/PROPERTYPUT | XlSeriesColorGradientStyle | declared | 3261 | Metadata Only | Reviewed | Not Tested | |
| SeriesColorMinGradientStop | PROPERTYGET | ChartSeriesGradientStopData | declared | 3262 | Metadata Only | Reviewed | Not Tested | |
| SeriesColorMidGradientStop | PROPERTYGET | ChartSeriesGradientStopData | declared | 3263 | Metadata Only | Reviewed | Not Tested | |
| SeriesColorMaxGradientStop | PROPERTYGET | ChartSeriesGradientStopData | declared | 3264 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Points | Object | 1 | declared | 70 | Implemented | Reviewed | Live Tested | |
| ClearFormats | AutomationValue | 0 | declared | 112 | Metadata Only | Reviewed | Not Tested | |
| Delete | AutomationValue | 0 | declared | 117 | Implemented | Reviewed | Live Tested | |
| _ApplyDataLabels | AutomationValue | 4 | declared | 151 | Implemented | Reviewed | Live Tested | |
| ErrorBar | AutomationValue | 5 | declared | 152 | Implemented | Reviewed | Live Tested | |
| Trendlines | Object | 1 | declared | 154 | Implemented | Reviewed | Live Tested | |
| DataLabels | Object | 1 | declared | 157 | Implemented | Reviewed | Live Tested | |
| Paste | AutomationValue | 0 | declared | 211 | Metadata Only | Reviewed | Not Tested | |
| Select | AutomationValue | 0 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| Copy | AutomationValue | 0 | declared | 551 | Metadata Only | Reviewed | Not Tested | |
| ApplyCustomType | Unknown | 1 | declared | 1401 | Metadata Only | Reviewed | Not Tested | |
| ApplyDataLabels | AutomationValue | 10 | declared | 1922 | Metadata Only | Reviewed | Not Tested | |
| SetProperty | Unknown | 2 | declared | 3253 | Metadata Only | Reviewed | Not Tested | |
| GetProperty | AutomationValue | 1 | declared | 3256 | Metadata Only | Reviewed | Not Tested | |
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
