# Chart

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `_Chart` |
| GUID | `{000208d6-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4288 |
| Crate type | `excel_com::Chart` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Blocked |

## Capabilities

No capability metadata is recorded for this surface.


## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _CodeName | PROPERTYGET/PROPERTYPUT | String | declared | -2147418112 | Metadata Only | Reviewed | Not Tested | |
| Area3DGroup | PROPERTYGET | ChartGroup | declared | 17 | Metadata Only | Reviewed | Not Tested | |
| Bar3DGroup | PROPERTYGET | ChartGroup | declared | 18 | Metadata Only | Reviewed | Not Tested | |
| Column3DGroup | PROPERTYGET | ChartGroup | declared | 19 | Metadata Only | Reviewed | Not Tested | |
| Line3DGroup | PROPERTYGET | ChartGroup | declared | 20 | Metadata Only | Reviewed | Not Tested | |
| Pie3DGroup | PROPERTYGET | ChartGroup | declared | 21 | Metadata Only | Reviewed | Not Tested | |
| SurfaceGroup | PROPERTYGET | ChartGroup | declared | 22 | Metadata Only | Reviewed | Not Tested | |
| DepthPercent | PROPERTYGET/PROPERTYPUT | i32 | declared | 48 | Implemented | Reviewed | Live Tested | |
| Elevation | PROPERTYGET/PROPERTYPUT | i32 | declared | 49 | Implemented | Reviewed | Live Tested | |
| GapDepth | PROPERTYGET/PROPERTYPUT | i32 | declared | 50 | Implemented | Reviewed | Live Tested | |
| HasAxis | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 52 | Metadata Only | Reviewed | Not Tested | |
| HasLegend | PROPERTYGET/PROPERTYPUT | bool | declared | 53 | Implemented | Reviewed | Live Tested | |
| HasTitle | PROPERTYGET/PROPERTYPUT | bool | declared | 54 | Implemented | Reviewed | Live Tested | |
| HeightPercent | PROPERTYGET/PROPERTYPUT | i32 | declared | 55 | Implemented | Reviewed | Live Tested | |
| Perspective | PROPERTYGET/PROPERTYPUT | i32 | declared | 57 | Implemented | Reviewed | Live Tested | |
| RightAngleAxes | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 58 | Implemented | Reviewed | Live Tested | |
| Rotation | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 59 | Implemented | Reviewed | Live Tested | |
| Corners | PROPERTYGET | Corners | declared | 79 | Metadata Only | Reviewed | Not Tested | |
| ChartArea | PROPERTYGET | ChartArea | declared | 80 | Implemented | Reviewed | Live Tested | |
| ChartTitle | PROPERTYGET | ChartTitle | declared | 81 | Implemented | Reviewed | Live Tested | |
| Floor | PROPERTYGET | Floor | declared | 83 | Implemented | Reviewed | Live Tested | |
| Legend | PROPERTYGET | Legend | declared | 84 | Implemented | Reviewed | Live Tested | |
| PlotArea | PROPERTYGET | PlotArea | declared | 85 | Implemented | Reviewed | Live Tested | |
| Walls | PROPERTYGET | Walls | declared | 86 | Implemented | Reviewed | Live Tested | |
| PlotVisibleOnly | PROPERTYGET/PROPERTYPUT | bool | declared | 92 | Metadata Only | Reviewed | Not Tested | |
| DisplayBlanksAs | PROPERTYGET/PROPERTYPUT | XlDisplayBlanksAs | declared | 93 | Metadata Only | Reviewed | Not Tested | |
| SizeWithWindow | PROPERTYGET/PROPERTYPUT | bool | declared | 94 | Metadata Only | Reviewed | Not Tested | |
| AutoScaling | PROPERTYGET/PROPERTYPUT | bool | declared | 107 | Implemented | Reviewed | Live Tested | |
| Type | PROPERTYGET/PROPERTYPUT | i32 | declared | 108 | Metadata Only | Reviewed | Not Tested | |
| SubType | PROPERTYGET/PROPERTYPUT | i32 | declared | 109 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET/PROPERTYPUT | String | declared | 110 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| PlotBy | PROPERTYGET/PROPERTYPUT | XlRowCol | declared | 202 | Implemented | Reviewed | Live Tested | |
| WallsAndGridlines2D | PROPERTYGET/PROPERTYPUT | bool | declared | 210 | Metadata Only | Reviewed | Not Tested | |
| ProtectContents | PROPERTYGET | bool | declared | 292 | Metadata Only | Reviewed | Not Tested | |
| ProtectDrawingObjects | PROPERTYGET | bool | declared | 293 | Metadata Only | Reviewed | Not Tested | |
| Index | PROPERTYGET | i32 | declared | 486 | Implemented | Reviewed | Live Tested | |
| Next | PROPERTYGET | Object | declared | 502 | Metadata Only | Reviewed | Not Tested | |
| Previous | PROPERTYGET | Object | declared | 503 | Metadata Only | Reviewed | Not Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | XlSheetVisibility | declared | 558 | Metadata Only | Reviewed | Not Tested | |
| OnDoubleClick | PROPERTYGET/PROPERTYPUT | String | declared | 628 | Metadata Only | Reviewed | Not Tested | |
| PageSetup | PROPERTYGET | PageSetup | declared | 998 | Metadata Only | Reviewed | Not Tested | |
| OnSheetActivate | PROPERTYGET/PROPERTYPUT | String | declared | 1031 | Metadata Only | Reviewed | Not Tested | |
| Tab | PROPERTYGET | Tab | declared | 1041 | Metadata Only | Reviewed | Not Tested | |
| OnSheetDeactivate | PROPERTYGET/PROPERTYPUT | String | declared | 1081 | Metadata Only | Reviewed | Not Tested | |
| ProtectionMode | PROPERTYGET | bool | declared | 1159 | Metadata Only | Reviewed | Not Tested | |
| CodeName | PROPERTYGET | String | declared | 1373 | Metadata Only | Reviewed | Not Tested | |
| Shapes | PROPERTYGET | Shapes | declared | 1377 | Implemented | Reviewed | Live Tested | |
| Hyperlinks | PROPERTYGET | Hyperlinks | declared | 1393 | Metadata Only | Reviewed | Not Tested | |
| DataTable | PROPERTYGET | DataTable | declared | 1395 | Metadata Only | Reviewed | Not Tested | |
| HasDataTable | PROPERTYGET/PROPERTYPUT | bool | declared | 1396 | Metadata Only | Reviewed | Not Tested | |
| ShowWindow | PROPERTYGET/PROPERTYPUT | bool | declared | 1399 | Metadata Only | Reviewed | Not Tested | |
| ChartType | PROPERTYGET/PROPERTYPUT | XlChartType | declared | 1400 | Implemented | Reviewed | Live Tested | |
| BarShape | PROPERTYGET/PROPERTYPUT | XlBarShape | declared | 1403 | Metadata Only | Reviewed | Not Tested | |
| ProtectFormatting | PROPERTYGET/PROPERTYPUT | bool | declared | 1405 | Metadata Only | Reviewed | Not Tested | |
| ProtectData | PROPERTYGET/PROPERTYPUT | bool | declared | 1406 | Metadata Only | Reviewed | Not Tested | |
| ProtectGoalSeek | PROPERTYGET/PROPERTYPUT | bool | declared | 1407 | Metadata Only | Reviewed | Not Tested | |
| ProtectSelection | PROPERTYGET/PROPERTYPUT | bool | declared | 1408 | Metadata Only | Reviewed | Not Tested | |
| PivotLayout | PROPERTYGET | PivotLayout | declared | 1814 | Metadata Only | Reviewed | Not Tested | |
| HasPivotFields | PROPERTYGET/PROPERTYPUT | bool | declared | 1815 | Metadata Only | Reviewed | Not Tested | |
| Scripts | PROPERTYGET | Scripts | declared | 1816 | Metadata Only | Reviewed | Not Tested | |
| MailEnvelope | PROPERTYGET | MsoEnvelope | declared | 2021 | Metadata Only | Reviewed | Not Tested | |
| ShowDataLabelsOverMaximum | PROPERTYGET/PROPERTYPUT | bool | declared | 2504 | Metadata Only | Reviewed | Not Tested | |
| SideWall | PROPERTYGET | Walls | declared | 2505 | Metadata Only | Reviewed | Not Tested | |
| BackWall | PROPERTYGET | Walls | declared | 2506 | Metadata Only | Reviewed | Not Tested | |
| ChartStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2509 | Implemented | Reviewed | Live Tested | |
| PrintedCommentPages | PROPERTYGET | i32 | declared | 2857 | Metadata Only | Reviewed | Not Tested | |
| Dummy24 | PROPERTYGET/PROPERTYPUT | bool | declared | 2858 | Metadata Only | Reviewed | Not Tested | |
| Dummy25 | PROPERTYGET/PROPERTYPUT | bool | declared | 2859 | Metadata Only | Reviewed | Not Tested | |
| ShowReportFilterFieldButtons | PROPERTYGET/PROPERTYPUT | bool | declared | 2860 | Metadata Only | Reviewed | Not Tested | |
| ShowLegendFieldButtons | PROPERTYGET/PROPERTYPUT | bool | declared | 2861 | Metadata Only | Reviewed | Not Tested | |
| ShowAxisFieldButtons | PROPERTYGET/PROPERTYPUT | bool | declared | 2862 | Metadata Only | Reviewed | Not Tested | |
| ShowValueFieldButtons | PROPERTYGET/PROPERTYPUT | bool | declared | 2863 | Metadata Only | Reviewed | Not Tested | |
| ShowAllFieldButtons | PROPERTYGET/PROPERTYPUT | bool | declared | 2864 | Metadata Only | Reviewed | Not Tested | |
| CategoryLabelLevel | PROPERTYGET/PROPERTYPUT | XlCategoryLabelLevel | declared | 3048 | Metadata Only | Reviewed | Not Tested | |
| SeriesNameLevel | PROPERTYGET/PROPERTYPUT | XlSeriesNameLevel | declared | 3049 | Metadata Only | Reviewed | Not Tested | |
| HasHiddenContent | PROPERTYGET | bool | declared | 3050 | Metadata Only | Reviewed | Not Tested | |
| ChartColor | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 3052 | Metadata Only | Reviewed | Not Tested | |
| ShowExpandCollapseEntireFieldButtons | PROPERTYGET/PROPERTYPUT | bool | declared | 3166 | Metadata Only | Reviewed | Not Tested | |
| DisplayValueNotAvailableAsBlank | PROPERTYGET/PROPERTYPUT | bool | declared | 3333 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| _Evaluate | AutomationValue | 1 | declared | -5 | Metadata Only | Reviewed | Not Tested | |
| Evaluate | AutomationValue | 1 | declared | 1 | Metadata Only | Reviewed | Not Tested | |
| ChartGroups | Object | 1 | declared | 8 | Implemented | Reviewed | Live Tested | |
| AreaGroups | Object | 1 | declared | 9 | Metadata Only | Reviewed | Not Tested | |
| BarGroups | Object | 1 | declared | 10 | Metadata Only | Reviewed | Not Tested | |
| ColumnGroups | Object | 1 | declared | 11 | Metadata Only | Reviewed | Not Tested | |
| LineGroups | Object | 1 | declared | 12 | Metadata Only | Reviewed | Not Tested | |
| PieGroups | Object | 1 | declared | 13 | Metadata Only | Reviewed | Not Tested | |
| DoughnutGroups | Object | 1 | declared | 14 | Metadata Only | Reviewed | Not Tested | |
| RadarGroups | Object | 1 | declared | 15 | Metadata Only | Reviewed | Not Tested | |
| XYGroups | Object | 1 | declared | 16 | Metadata Only | Reviewed | Not Tested | |
| Axes | Object | 2 | declared | 23 | Implemented | Reviewed | Live Tested | |
| SeriesCollection | Object | 1 | declared | 68 | Implemented | Reviewed | Live Tested | |
| DrawingObjects | Object | 1 | declared | 88 | Metadata Only | Reviewed | Not Tested | |
| AutoFormat | Unknown | 2 | declared | 114 | Metadata Only | Reviewed | Not Tested | |
| Delete | Unknown | 0 | declared | 117 | Implemented | Reviewed | Live Tested | |
| _ApplyDataLabels | Unknown | 4 | declared | 151 | Metadata Only | Reviewed | Not Tested | |
| ChartWizard | Unknown | 11 | declared | 196 | Metadata Only | Reviewed | Not Tested | |
| Paste | Unknown | 1 | declared | 211 | Metadata Only | Reviewed | Not Tested | |
| CopyPicture | Unknown | 3 | declared | 213 | Implemented | Reviewed | Live Tested | |
| SetDefaultChart | Unknown | 1 | declared | 219 | Metadata Only | Reviewed | Not Tested | |
| Select | Unknown | 1 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| PrintPreview | Unknown | 1 | declared | 281 | Metadata Only | Reviewed | Not Tested | |
| _Protect | Unknown | 5 | declared | 282 | Metadata Only | Reviewed | Not Tested | |
| __SaveAs | Unknown | 9 | declared | 284 | Metadata Only | Reviewed | Not Tested | |
| Unprotect | Unknown | 1 | declared | 285 | Metadata Only | Reviewed | Not Tested | |
| Activate | Unknown | 0 | declared | 304 | Implemented | Reviewed | Live Tested | |
| CreatePublisher | Unknown | 7 | declared | 458 | Metadata Only | Reviewed | Not Tested | |
| CheckSpelling | Unknown | 4 | declared | 505 | Metadata Only | Reviewed | Not Tested | |
| Copy | Unknown | 2 | declared | 551 | Implemented | Reviewed | Live Tested | |
| Buttons | Object | 1 | declared | 557 | Metadata Only | Reviewed | Not Tested | |
| Move | Unknown | 2 | declared | 637 | Implemented | Reviewed | Live Tested | |
| Arcs | Object | 1 | declared | 760 | Metadata Only | Reviewed | Not Tested | |
| Lines | Object | 1 | declared | 767 | Metadata Only | Reviewed | Not Tested | |
| Pictures | Object | 1 | declared | 771 | Metadata Only | Reviewed | Not Tested | |
| Drawings | Object | 1 | declared | 772 | Metadata Only | Reviewed | Not Tested | |
| Rectangles | Object | 1 | declared | 774 | Metadata Only | Reviewed | Not Tested | |
| TextBoxes | Object | 1 | declared | 777 | Metadata Only | Reviewed | Not Tested | |
| OLEObjects | Object | 1 | declared | 799 | Metadata Only | Reviewed | Not Tested | |
| Ovals | Object | 1 | declared | 801 | Metadata Only | Reviewed | Not Tested | |
| CheckBoxes | Object | 1 | declared | 824 | Metadata Only | Reviewed | Not Tested | |
| OptionButtons | Object | 1 | declared | 826 | Metadata Only | Reviewed | Not Tested | |
| ScrollBars | Object | 1 | declared | 830 | Metadata Only | Reviewed | Not Tested | |
| ListBoxes | Object | 1 | declared | 832 | Metadata Only | Reviewed | Not Tested | |
| GroupBoxes | Object | 1 | declared | 834 | Metadata Only | Reviewed | Not Tested | |
| DropDowns | Object | 1 | declared | 836 | Metadata Only | Reviewed | Not Tested | |
| Spinners | Object | 1 | declared | 838 | Metadata Only | Reviewed | Not Tested | |
| Labels | Object | 1 | declared | 841 | Metadata Only | Reviewed | Not Tested | |
| __PrintOut | Unknown | 7 | declared | 905 | Metadata Only | Reviewed | Not Tested | |
| ChartObjects | Object | 1 | declared | 1060 | Metadata Only | Reviewed | Not Tested | |
| GroupObjects | Object | 1 | declared | 1113 | Metadata Only | Reviewed | Not Tested | |
| Deselect | Unknown | 0 | declared | 1120 | Metadata Only | Reviewed | Not Tested | |
| SetBackgroundPicture | Unknown | 1 | declared | 1188 | Metadata Only | Reviewed | Not Tested | |
| Location | Chart | 2 | declared | 1397 | Metadata Only | Reviewed | Not Tested | |
| ApplyCustomType | Unknown | 2 | declared | 1401 | Metadata Only | Reviewed | Not Tested | |
| CopyChartBuild | Unknown | 0 | declared | 1404 | Metadata Only | Reviewed | Not Tested | |
| GetChartElement | Unknown | 5 | declared | 1409 | Metadata Only | Reviewed | Not Tested | |
| SetSourceData | Unknown | 2 | declared | 1413 | Implemented | Reviewed | Live Tested | |
| Export | bool | 3 | declared | 1414 | Implemented | Reviewed | Live Tested | |
| Refresh | Unknown | 0 | declared | 1417 | Implemented | Reviewed | Live Tested | |
| _PrintOut | Unknown | 8 | declared | 1772 | Metadata Only | Reviewed | Not Tested | |
| ApplyDataLabels | Unknown | 10 | declared | 1922 | Metadata Only | Reviewed | Not Tested | |
| _SaveAs | Unknown | 10 | declared | 1925 | Metadata Only | Reviewed | Not Tested | |
| Protect | Unknown | 5 | declared | 2029 | Metadata Only | Reviewed | Not Tested | |
| PrintOut | Unknown | 8 | declared | 2361 | Metadata Only | Reviewed | Not Tested | |
| _ExportAsFixedFormat | Unknown | 9 | declared | 2493 | Metadata Only | Reviewed | Not Tested | |
| ApplyLayout | Unknown | 2 | declared | 2500 | Implemented | Reviewed | Live Tested | |
| SetElement | Unknown | 1 | declared | 2502 | Metadata Only | Reviewed | Not Tested | |
| ApplyChartTemplate | Unknown | 1 | declared | 2507 | Implemented | Reviewed | Live Tested | |
| SaveChartTemplate | Unknown | 1 | declared | 2508 | Implemented | Reviewed | Live Tested | |
| ClearToMatchStyle | Unknown | 0 | declared | 2510 | Metadata Only | Reviewed | Not Tested | |
| FullSeriesCollection | Object | 1 | declared | 3047 | Metadata Only | Reviewed | Not Tested | |
| DeleteHiddenContent | Unknown | 0 | declared | 3051 | Metadata Only | Reviewed | Not Tested | |
| ClearToMatchColorStyle | Unknown | 0 | declared | 3053 | Metadata Only | Reviewed | Not Tested | |
| SaveAs | Unknown | 10 | declared | 3174 | Metadata Only | Reviewed | Not Tested | |
| ExportAsFixedFormat | Unknown | 10 | declared | 3175 | Metadata Only | Reviewed | Not Tested | |
| SetProperty | Unknown | 2 | declared | 3253 | Metadata Only | Reviewed | Not Tested | |
| GetProperty | AutomationValue | 1 | declared | 3256 | Metadata Only | Reviewed | Not Tested | |
| _Dummy23 | Unknown | 0 | declared | 65559 | Metadata Only | Reviewed | Not Tested | |
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
