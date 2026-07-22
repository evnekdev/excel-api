# Worksheet

## Summary

A worksheet object within a workbook. It is structurally inventoried while worksheet operations remain in the research tools.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `_Worksheet` |
| GUID | `{000208d8-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4288 |
| Crate type | `excel_com::Worksheet` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Cells` | `excel.range` | Metadata Only |
| `CircularReference` | `excel.range` | Metadata Only |
| `Columns` | `excel.range` | Metadata Only |
| `Range` | `excel.range` | Implemented |
| `Rows` | `excel.range` | Metadata Only |
| `UsedRange` | `excel.range` | Implemented |
| `XmlDataQuery` | `excel.range` | Metadata Only |
| `XmlMapQuery` | `excel.range` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _CodeName | PROPERTYGET/PROPERTYPUT | String | declared | -2147418112 | Metadata Only | Reviewed | Not Tested | |
| Outline | PROPERTYGET | Outline | declared | 102 | Metadata Only | Reviewed | Not Tested | |
| Type | PROPERTYGET | XlSheetType | declared | 108 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET/PROPERTYPUT | String | declared | 110 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Protection | PROPERTYGET | Protection | declared | 176 | Metadata Only | Reviewed | Not Tested | |
| Range | PROPERTYGET | Range | declared | 197 | Implemented | Reviewed | Live Tested | |
| Cells | PROPERTYGET | Range | declared | 238 | Metadata Only | Reviewed | Not Tested | |
| Columns | PROPERTYGET | Range | declared | 241 | Metadata Only | Reviewed | Not Tested | |
| Rows | PROPERTYGET | Range | declared | 258 | Metadata Only | Reviewed | Not Tested | |
| ProtectContents | PROPERTYGET | bool | declared | 292 | Metadata Only | Reviewed | Not Tested | |
| ProtectDrawingObjects | PROPERTYGET | bool | declared | 293 | Metadata Only | Reviewed | Not Tested | |
| ProtectScenarios | PROPERTYGET | bool | declared | 294 | Metadata Only | Reviewed | Not Tested | |
| TransitionExpEval | PROPERTYGET/PROPERTYPUT | bool | declared | 401 | Metadata Only | Reviewed | Not Tested | |
| TransitionFormEntry | PROPERTYGET/PROPERTYPUT | bool | declared | 402 | Metadata Only | Reviewed | Not Tested | |
| StandardHeight | PROPERTYGET | f64 | declared | 407 | Metadata Only | Reviewed | Not Tested | |
| StandardWidth | PROPERTYGET/PROPERTYPUT | f64 | declared | 408 | Metadata Only | Reviewed | Not Tested | |
| UsedRange | PROPERTYGET | Range | declared | 412 | Implemented | Reviewed | Live Tested | |
| Names | PROPERTYGET | Names | declared | 442 | Metadata Only | Reviewed | Not Tested | |
| Index | PROPERTYGET | i32 | declared | 486 | Implemented | Reviewed | Live Tested | |
| Next | PROPERTYGET | Object | declared | 502 | Metadata Only | Reviewed | Not Tested | |
| Previous | PROPERTYGET | Object | declared | 503 | Metadata Only | Reviewed | Not Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | XlSheetVisibility | declared | 558 | Implemented | Reviewed | Live Tested | |
| Comments | PROPERTYGET | Comments | declared | 575 | Metadata Only | Reviewed | Not Tested | |
| OnCalculate | PROPERTYGET/PROPERTYPUT | String | declared | 625 | Metadata Only | Reviewed | Not Tested | |
| OnEntry | PROPERTYGET/PROPERTYPUT | String | declared | 627 | Metadata Only | Reviewed | Not Tested | |
| OnDoubleClick | PROPERTYGET/PROPERTYPUT | String | declared | 628 | Metadata Only | Reviewed | Not Tested | |
| OnData | PROPERTYGET/PROPERTYPUT | String | declared | 629 | Metadata Only | Reviewed | Not Tested | |
| DisplayAutomaticPageBreaks | PROPERTYGET/PROPERTYPUT | bool | declared | 643 | Metadata Only | Reviewed | Not Tested | |
| _DisplayRightToLeft | PROPERTYGET/PROPERTYPUT | i32 | declared | 648 | Metadata Only | Reviewed | Not Tested | |
| ConsolidationFunction | PROPERTYGET | XlConsolidationFunction | declared | 789 | Metadata Only | Reviewed | Not Tested | |
| ConsolidationOptions | PROPERTYGET | AutomationValue | declared | 790 | Metadata Only | Reviewed | Not Tested | |
| ConsolidationSources | PROPERTYGET | AutomationValue | declared | 791 | Metadata Only | Reviewed | Not Tested | |
| AutoFilterMode | PROPERTYGET/PROPERTYPUT | bool | declared | 792 | Metadata Only | Reviewed | Not Tested | |
| _AutoFilter | PROPERTYGET | AutoFilter | declared | 793 | Metadata Only | Reviewed | Not Tested | |
| FilterMode | PROPERTYGET | bool | declared | 800 | Metadata Only | Reviewed | Not Tested | |
| _Sort | PROPERTYGET | Sort | declared | 880 | Metadata Only | Reviewed | Not Tested | |
| PageSetup | PROPERTYGET | PageSetup | declared | 998 | Metadata Only | Reviewed | Not Tested | |
| OnSheetActivate | PROPERTYGET/PROPERTYPUT | String | declared | 1031 | Metadata Only | Reviewed | Not Tested | |
| Tab | PROPERTYGET | Tab | declared | 1041 | Metadata Only | Reviewed | Not Tested | |
| CircularReference | PROPERTYGET | Range | declared | 1069 | Metadata Only | Reviewed | Not Tested | |
| OnSheetDeactivate | PROPERTYGET/PROPERTYPUT | String | declared | 1081 | Metadata Only | Reviewed | Not Tested | |
| EnableAutoFilter | PROPERTYGET/PROPERTYPUT | bool | declared | 1156 | Metadata Only | Reviewed | Not Tested | |
| EnableOutlining | PROPERTYGET/PROPERTYPUT | bool | declared | 1157 | Metadata Only | Reviewed | Not Tested | |
| EnablePivotTable | PROPERTYGET/PROPERTYPUT | bool | declared | 1158 | Metadata Only | Reviewed | Not Tested | |
| ProtectionMode | PROPERTYGET | bool | declared | 1159 | Metadata Only | Reviewed | Not Tested | |
| CodeName | PROPERTYGET | String | declared | 1373 | Metadata Only | Reviewed | Not Tested | |
| Shapes | PROPERTYGET | Shapes | declared | 1377 | Metadata Only | Reviewed | Not Tested | |
| Hyperlinks | PROPERTYGET | Hyperlinks | declared | 1393 | Metadata Only | Reviewed | Not Tested | |
| HPageBreaks | PROPERTYGET | HPageBreaks | declared | 1418 | Metadata Only | Reviewed | Not Tested | |
| VPageBreaks | PROPERTYGET | VPageBreaks | declared | 1419 | Metadata Only | Reviewed | Not Tested | |
| EnableCalculation | PROPERTYGET/PROPERTYPUT | bool | declared | 1424 | Metadata Only | Reviewed | Not Tested | |
| EnableSelection | PROPERTYGET/PROPERTYPUT | XlEnableSelection | declared | 1425 | Metadata Only | Reviewed | Not Tested | |
| ScrollArea | PROPERTYGET/PROPERTYPUT | String | declared | 1433 | Metadata Only | Reviewed | Not Tested | |
| QueryTables | PROPERTYGET | QueryTables | declared | 1434 | Metadata Only | Reviewed | Not Tested | |
| DisplayPageBreaks | PROPERTYGET/PROPERTYPUT | bool | declared | 1435 | Metadata Only | Reviewed | Not Tested | |
| DisplayRightToLeft | PROPERTYGET/PROPERTYPUT | bool | declared | 1774 | Metadata Only | Reviewed | Not Tested | |
| Scripts | PROPERTYGET | Scripts | declared | 1816 | Metadata Only | Reviewed | Not Tested | |
| SmartTags | PROPERTYGET | SmartTags | declared | 2016 | Metadata Only | Reviewed | Not Tested | |
| MailEnvelope | PROPERTYGET | MsoEnvelope | declared | 2021 | Metadata Only | Reviewed | Not Tested | |
| CustomProperties | PROPERTYGET | CustomProperties | declared | 2030 | Metadata Only | Reviewed | Not Tested | |
| ListObjects | PROPERTYGET | ListObjects | declared | 2259 | Metadata Only | Reviewed | Not Tested | |
| EnableFormatConditionsCalculation | PROPERTYGET/PROPERTYPUT | bool | declared | 2511 | Metadata Only | Reviewed | Not Tested | |
| PrintedCommentPages | PROPERTYGET | i32 | declared | 2857 | Metadata Only | Reviewed | Not Tested | |
| CommentsThreaded | PROPERTYGET | CommentsThreaded | declared | 3282 | Metadata Only | Reviewed | Not Tested | |
| Sort | PROPERTYGET | Sort | declared | 3288 | Metadata Only | Reviewed | Not Tested | |
| AutoFilter | PROPERTYGET | AutoFilter | declared | 3289 | Metadata Only | Reviewed | Not Tested | |
| NamedSheetViews | PROPERTYGET | NamedSheetViewCollection | declared | 3309 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| _Evaluate | AutomationValue | 1 | declared | -5 | Metadata Only | Reviewed | Not Tested | |
| Evaluate | AutomationValue | 1 | declared | 1 | Metadata Only | Reviewed | Not Tested | |
| DrawingObjects | Object | 1 | declared | 88 | Metadata Only | Reviewed | Not Tested | |
| Delete | Unknown | 0 | declared | 117 | Metadata Only | Reviewed | Not Tested | |
| Paste | Unknown | 2 | declared | 211 | Metadata Only | Reviewed | Not Tested | |
| Select | Unknown | 1 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| Calculate | Unknown | 0 | declared | 279 | Metadata Only | Reviewed | Not Tested | |
| PrintPreview | Unknown | 1 | declared | 281 | Metadata Only | Reviewed | Not Tested | |
| _Protect | Unknown | 5 | declared | 282 | Metadata Only | Reviewed | Not Tested | |
| __SaveAs | Unknown | 9 | declared | 284 | Metadata Only | Reviewed | Not Tested | |
| Unprotect | Unknown | 1 | declared | 285 | Metadata Only | Reviewed | Not Tested | |
| Activate | Unknown | 0 | declared | 304 | Metadata Only | Reviewed | Not Tested | |
| ShowDataForm | Unknown | 0 | declared | 409 | Metadata Only | Reviewed | Not Tested | |
| CheckSpelling | Unknown | 4 | declared | 505 | Metadata Only | Reviewed | Not Tested | |
| Copy | Unknown | 2 | declared | 551 | Metadata Only | Reviewed | Not Tested | |
| Buttons | Object | 1 | declared | 557 | Metadata Only | Reviewed | Not Tested | |
| Move | Unknown | 2 | declared | 637 | Metadata Only | Reviewed | Not Tested | |
| PivotTableWizard | PivotTable | 16 | declared | 684 | Metadata Only | Reviewed | Not Tested | |
| PivotTables | Object | 1 | declared | 690 | Metadata Only | Reviewed | Not Tested | |
| Arcs | Object | 1 | declared | 760 | Metadata Only | Reviewed | Not Tested | |
| Lines | Object | 1 | declared | 767 | Metadata Only | Reviewed | Not Tested | |
| Pictures | Object | 1 | declared | 771 | Metadata Only | Reviewed | Not Tested | |
| Drawings | Object | 1 | declared | 772 | Metadata Only | Reviewed | Not Tested | |
| Rectangles | Object | 1 | declared | 774 | Metadata Only | Reviewed | Not Tested | |
| TextBoxes | Object | 1 | declared | 777 | Metadata Only | Reviewed | Not Tested | |
| ShowAllData | Unknown | 0 | declared | 794 | Metadata Only | Reviewed | Not Tested | |
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
| Scenarios | Object | 1 | declared | 908 | Metadata Only | Reviewed | Not Tested | |
| ClearArrows | Unknown | 0 | declared | 970 | Metadata Only | Reviewed | Not Tested | |
| _PasteSpecial | Unknown | 6 | declared | 1027 | Metadata Only | Reviewed | Not Tested | |
| ChartObjects | Object | 1 | declared | 1060 | Metadata Only | Reviewed | Not Tested | |
| GroupObjects | Object | 1 | declared | 1113 | Metadata Only | Reviewed | Not Tested | |
| SetBackgroundPicture | Unknown | 1 | declared | 1188 | Metadata Only | Reviewed | Not Tested | |
| ResetAllPageBreaks | Unknown | 0 | declared | 1426 | Metadata Only | Reviewed | Not Tested | |
| ClearCircles | Unknown | 0 | declared | 1436 | Metadata Only | Reviewed | Not Tested | |
| CircleInvalid | Unknown | 0 | declared | 1437 | Metadata Only | Reviewed | Not Tested | |
| _PrintOut | Unknown | 8 | declared | 1772 | Metadata Only | Reviewed | Not Tested | |
| _CheckSpelling | Unknown | 6 | declared | 1817 | Metadata Only | Reviewed | Not Tested | |
| _SaveAs | Unknown | 10 | declared | 1925 | Metadata Only | Reviewed | Not Tested | |
| PasteSpecial | Unknown | 7 | declared | 1928 | Metadata Only | Reviewed | Not Tested | |
| Protect | Unknown | 16 | declared | 2029 | Metadata Only | Reviewed | Not Tested | |
| XmlDataQuery | Range | 3 | declared | 2260 | Metadata Only | Reviewed | Not Tested | |
| XmlMapQuery | Range | 3 | declared | 2263 | Metadata Only | Reviewed | Not Tested | |
| PrintOut | Unknown | 9 | declared | 2361 | Metadata Only | Reviewed | Not Tested | |
| _ExportAsFixedFormat | Unknown | 9 | declared | 2493 | Metadata Only | Reviewed | Not Tested | |
| SaveAs | Unknown | 10 | declared | 3174 | Metadata Only | Reviewed | Not Tested | |
| ExportAsFixedFormat | Unknown | 10 | declared | 3175 | Metadata Only | Reviewed | Not Tested | |
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
