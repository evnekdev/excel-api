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
| Crate type | `excel_com::Worksheet` |
| Implementation | Metadata Only |
| Documentation | Reviewed |
| Tests | Not Tested |

## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Cells` | `excel.range` | Metadata Only |
| `CircularReference` | `excel.range` | Metadata Only |
| `Columns` | `excel.range` | Metadata Only |
| `Range` | `excel.range` | Metadata Only |
| `Rows` | `excel.range` | Metadata Only |
| `UsedRange` | `excel.range` | Metadata Only |
| `XmlDataQuery` | `excel.range` | Metadata Only |
| `XmlMapQuery` | `excel.range` | Metadata Only |

## Properties

| Property | Access | Type | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---:|---|---|---|---|
| _CodeName | PROPERTYGET/PROPERTYPUT | String | -2147418112 | Metadata Only | Reviewed | Not Tested | |
| Outline | PROPERTYGET | Outline | 102 | Metadata Only | Reviewed | Not Tested | |
| Type | PROPERTYGET | XlSheetType | 108 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET/PROPERTYPUT | String | 110 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | 150 | Metadata Only | Reviewed | Not Tested | |
| Protection | PROPERTYGET | Protection | 176 | Metadata Only | Reviewed | Not Tested | |
| Range | PROPERTYGET | Range | 197 | Metadata Only | Reviewed | Not Tested | |
| Cells | PROPERTYGET | Range | 238 | Metadata Only | Reviewed | Not Tested | |
| Columns | PROPERTYGET | Range | 241 | Metadata Only | Reviewed | Not Tested | |
| Rows | PROPERTYGET | Range | 258 | Metadata Only | Reviewed | Not Tested | |
| ProtectContents | PROPERTYGET | bool | 292 | Metadata Only | Reviewed | Not Tested | |
| ProtectDrawingObjects | PROPERTYGET | bool | 293 | Metadata Only | Reviewed | Not Tested | |
| ProtectScenarios | PROPERTYGET | bool | 294 | Metadata Only | Reviewed | Not Tested | |
| TransitionExpEval | PROPERTYGET/PROPERTYPUT | bool | 401 | Metadata Only | Reviewed | Not Tested | |
| TransitionFormEntry | PROPERTYGET/PROPERTYPUT | bool | 402 | Metadata Only | Reviewed | Not Tested | |
| StandardHeight | PROPERTYGET | f64 | 407 | Metadata Only | Reviewed | Not Tested | |
| StandardWidth | PROPERTYGET/PROPERTYPUT | f64 | 408 | Metadata Only | Reviewed | Not Tested | |
| UsedRange | PROPERTYGET | Range | 412 | Metadata Only | Reviewed | Not Tested | |
| Names | PROPERTYGET | Names | 442 | Metadata Only | Reviewed | Not Tested | |
| Index | PROPERTYGET | i32 | 486 | Metadata Only | Reviewed | Not Tested | |
| Next | PROPERTYGET | Object | 502 | Metadata Only | Reviewed | Not Tested | |
| Previous | PROPERTYGET | Object | 503 | Metadata Only | Reviewed | Not Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | XlSheetVisibility | 558 | Metadata Only | Reviewed | Not Tested | |
| Comments | PROPERTYGET | Comments | 575 | Metadata Only | Reviewed | Not Tested | |
| OnCalculate | PROPERTYGET/PROPERTYPUT | String | 625 | Metadata Only | Reviewed | Not Tested | |
| OnEntry | PROPERTYGET/PROPERTYPUT | String | 627 | Metadata Only | Reviewed | Not Tested | |
| OnDoubleClick | PROPERTYGET/PROPERTYPUT | String | 628 | Metadata Only | Reviewed | Not Tested | |
| OnData | PROPERTYGET/PROPERTYPUT | String | 629 | Metadata Only | Reviewed | Not Tested | |
| DisplayAutomaticPageBreaks | PROPERTYGET/PROPERTYPUT | bool | 643 | Metadata Only | Reviewed | Not Tested | |
| _DisplayRightToLeft | PROPERTYGET/PROPERTYPUT | i32 | 648 | Metadata Only | Reviewed | Not Tested | |
| ConsolidationFunction | PROPERTYGET | XlConsolidationFunction | 789 | Metadata Only | Reviewed | Not Tested | |
| ConsolidationOptions | PROPERTYGET | AutomationValue | 790 | Metadata Only | Reviewed | Not Tested | |
| ConsolidationSources | PROPERTYGET | AutomationValue | 791 | Metadata Only | Reviewed | Not Tested | |
| AutoFilterMode | PROPERTYGET/PROPERTYPUT | bool | 792 | Metadata Only | Reviewed | Not Tested | |
| _AutoFilter | PROPERTYGET | AutoFilter | 793 | Metadata Only | Reviewed | Not Tested | |
| FilterMode | PROPERTYGET | bool | 800 | Metadata Only | Reviewed | Not Tested | |
| _Sort | PROPERTYGET | Sort | 880 | Metadata Only | Reviewed | Not Tested | |
| PageSetup | PROPERTYGET | PageSetup | 998 | Metadata Only | Reviewed | Not Tested | |
| OnSheetActivate | PROPERTYGET/PROPERTYPUT | String | 1031 | Metadata Only | Reviewed | Not Tested | |
| Tab | PROPERTYGET | Tab | 1041 | Metadata Only | Reviewed | Not Tested | |
| CircularReference | PROPERTYGET | Range | 1069 | Metadata Only | Reviewed | Not Tested | |
| OnSheetDeactivate | PROPERTYGET/PROPERTYPUT | String | 1081 | Metadata Only | Reviewed | Not Tested | |
| EnableAutoFilter | PROPERTYGET/PROPERTYPUT | bool | 1156 | Metadata Only | Reviewed | Not Tested | |
| EnableOutlining | PROPERTYGET/PROPERTYPUT | bool | 1157 | Metadata Only | Reviewed | Not Tested | |
| EnablePivotTable | PROPERTYGET/PROPERTYPUT | bool | 1158 | Metadata Only | Reviewed | Not Tested | |
| ProtectionMode | PROPERTYGET | bool | 1159 | Metadata Only | Reviewed | Not Tested | |
| CodeName | PROPERTYGET | String | 1373 | Metadata Only | Reviewed | Not Tested | |
| Shapes | PROPERTYGET | Shapes | 1377 | Metadata Only | Reviewed | Not Tested | |
| Hyperlinks | PROPERTYGET | Hyperlinks | 1393 | Metadata Only | Reviewed | Not Tested | |
| HPageBreaks | PROPERTYGET | HPageBreaks | 1418 | Metadata Only | Reviewed | Not Tested | |
| VPageBreaks | PROPERTYGET | VPageBreaks | 1419 | Metadata Only | Reviewed | Not Tested | |
| EnableCalculation | PROPERTYGET/PROPERTYPUT | bool | 1424 | Metadata Only | Reviewed | Not Tested | |
| EnableSelection | PROPERTYGET/PROPERTYPUT | XlEnableSelection | 1425 | Metadata Only | Reviewed | Not Tested | |
| ScrollArea | PROPERTYGET/PROPERTYPUT | String | 1433 | Metadata Only | Reviewed | Not Tested | |
| QueryTables | PROPERTYGET | QueryTables | 1434 | Metadata Only | Reviewed | Not Tested | |
| DisplayPageBreaks | PROPERTYGET/PROPERTYPUT | bool | 1435 | Metadata Only | Reviewed | Not Tested | |
| DisplayRightToLeft | PROPERTYGET/PROPERTYPUT | bool | 1774 | Metadata Only | Reviewed | Not Tested | |
| Scripts | PROPERTYGET | Scripts | 1816 | Metadata Only | Reviewed | Not Tested | |
| SmartTags | PROPERTYGET | SmartTags | 2016 | Metadata Only | Reviewed | Not Tested | |
| MailEnvelope | PROPERTYGET | MsoEnvelope | 2021 | Metadata Only | Reviewed | Not Tested | |
| CustomProperties | PROPERTYGET | CustomProperties | 2030 | Metadata Only | Reviewed | Not Tested | |
| ListObjects | PROPERTYGET | ListObjects | 2259 | Metadata Only | Reviewed | Not Tested | |
| EnableFormatConditionsCalculation | PROPERTYGET/PROPERTYPUT | bool | 2511 | Metadata Only | Reviewed | Not Tested | |
| PrintedCommentPages | PROPERTYGET | i32 | 2857 | Metadata Only | Reviewed | Not Tested | |
| CommentsThreaded | PROPERTYGET | CommentsThreaded | 3282 | Metadata Only | Reviewed | Not Tested | |
| Sort | PROPERTYGET | Sort | 3288 | Metadata Only | Reviewed | Not Tested | |
| AutoFilter | PROPERTYGET | AutoFilter | 3289 | Metadata Only | Reviewed | Not Tested | |
| NamedSheetViews | PROPERTYGET | NamedSheetViewCollection | 3309 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---:|---|---|---|---|
| _Evaluate | AutomationValue | 1 | -5 | Metadata Only | Reviewed | Not Tested | |
| Evaluate | AutomationValue | 1 | 1 | Metadata Only | Reviewed | Not Tested | |
| DrawingObjects | Object | 1 | 88 | Metadata Only | Reviewed | Not Tested | |
| Delete | Unknown | 0 | 117 | Metadata Only | Reviewed | Not Tested | |
| Paste | Unknown | 2 | 211 | Metadata Only | Reviewed | Not Tested | |
| Select | Unknown | 1 | 235 | Metadata Only | Reviewed | Not Tested | |
| Calculate | Unknown | 0 | 279 | Metadata Only | Reviewed | Not Tested | |
| PrintPreview | Unknown | 1 | 281 | Metadata Only | Reviewed | Not Tested | |
| _Protect | Unknown | 5 | 282 | Metadata Only | Reviewed | Not Tested | |
| __SaveAs | Unknown | 9 | 284 | Metadata Only | Reviewed | Not Tested | |
| Unprotect | Unknown | 1 | 285 | Metadata Only | Reviewed | Not Tested | |
| Activate | Unknown | 0 | 304 | Metadata Only | Reviewed | Not Tested | |
| ShowDataForm | Unknown | 0 | 409 | Metadata Only | Reviewed | Not Tested | |
| CheckSpelling | Unknown | 4 | 505 | Metadata Only | Reviewed | Not Tested | |
| Copy | Unknown | 2 | 551 | Metadata Only | Reviewed | Not Tested | |
| Buttons | Object | 1 | 557 | Metadata Only | Reviewed | Not Tested | |
| Move | Unknown | 2 | 637 | Metadata Only | Reviewed | Not Tested | |
| PivotTableWizard | PivotTable | 16 | 684 | Metadata Only | Reviewed | Not Tested | |
| PivotTables | Object | 1 | 690 | Metadata Only | Reviewed | Not Tested | |
| Arcs | Object | 1 | 760 | Metadata Only | Reviewed | Not Tested | |
| Lines | Object | 1 | 767 | Metadata Only | Reviewed | Not Tested | |
| Pictures | Object | 1 | 771 | Metadata Only | Reviewed | Not Tested | |
| Drawings | Object | 1 | 772 | Metadata Only | Reviewed | Not Tested | |
| Rectangles | Object | 1 | 774 | Metadata Only | Reviewed | Not Tested | |
| TextBoxes | Object | 1 | 777 | Metadata Only | Reviewed | Not Tested | |
| ShowAllData | Unknown | 0 | 794 | Metadata Only | Reviewed | Not Tested | |
| OLEObjects | Object | 1 | 799 | Metadata Only | Reviewed | Not Tested | |
| Ovals | Object | 1 | 801 | Metadata Only | Reviewed | Not Tested | |
| CheckBoxes | Object | 1 | 824 | Metadata Only | Reviewed | Not Tested | |
| OptionButtons | Object | 1 | 826 | Metadata Only | Reviewed | Not Tested | |
| ScrollBars | Object | 1 | 830 | Metadata Only | Reviewed | Not Tested | |
| ListBoxes | Object | 1 | 832 | Metadata Only | Reviewed | Not Tested | |
| GroupBoxes | Object | 1 | 834 | Metadata Only | Reviewed | Not Tested | |
| DropDowns | Object | 1 | 836 | Metadata Only | Reviewed | Not Tested | |
| Spinners | Object | 1 | 838 | Metadata Only | Reviewed | Not Tested | |
| Labels | Object | 1 | 841 | Metadata Only | Reviewed | Not Tested | |
| __PrintOut | Unknown | 7 | 905 | Metadata Only | Reviewed | Not Tested | |
| Scenarios | Object | 1 | 908 | Metadata Only | Reviewed | Not Tested | |
| ClearArrows | Unknown | 0 | 970 | Metadata Only | Reviewed | Not Tested | |
| _PasteSpecial | Unknown | 6 | 1027 | Metadata Only | Reviewed | Not Tested | |
| ChartObjects | Object | 1 | 1060 | Metadata Only | Reviewed | Not Tested | |
| GroupObjects | Object | 1 | 1113 | Metadata Only | Reviewed | Not Tested | |
| SetBackgroundPicture | Unknown | 1 | 1188 | Metadata Only | Reviewed | Not Tested | |
| ResetAllPageBreaks | Unknown | 0 | 1426 | Metadata Only | Reviewed | Not Tested | |
| ClearCircles | Unknown | 0 | 1436 | Metadata Only | Reviewed | Not Tested | |
| CircleInvalid | Unknown | 0 | 1437 | Metadata Only | Reviewed | Not Tested | |
| _PrintOut | Unknown | 8 | 1772 | Metadata Only | Reviewed | Not Tested | |
| _CheckSpelling | Unknown | 6 | 1817 | Metadata Only | Reviewed | Not Tested | |
| _SaveAs | Unknown | 10 | 1925 | Metadata Only | Reviewed | Not Tested | |
| PasteSpecial | Unknown | 7 | 1928 | Metadata Only | Reviewed | Not Tested | |
| Protect | Unknown | 16 | 2029 | Metadata Only | Reviewed | Not Tested | |
| XmlDataQuery | Range | 3 | 2260 | Metadata Only | Reviewed | Not Tested | |
| XmlMapQuery | Range | 3 | 2263 | Metadata Only | Reviewed | Not Tested | |
| PrintOut | Unknown | 9 | 2361 | Metadata Only | Reviewed | Not Tested | |
| _ExportAsFixedFormat | Unknown | 9 | 2493 | Metadata Only | Reviewed | Not Tested | |
| SaveAs | Unknown | 10 | 3174 | Metadata Only | Reviewed | Not Tested | |
| ExportAsFixedFormat | Unknown | 10 | 3175 | Metadata Only | Reviewed | Not Tested | |
| QueryInterface | Unknown | 2 | 1610612736 | Metadata Only | Reviewed | Not Tested | |
| AddRef | Unknown | 0 | 1610612737 | Metadata Only | Reviewed | Not Tested | |
| Release | Unknown | 0 | 1610612738 | Metadata Only | Reviewed | Not Tested | |
| GetTypeInfoCount | Unknown | 1 | 1610678272 | Metadata Only | Reviewed | Not Tested | |
| GetTypeInfo | Unknown | 3 | 1610678273 | Metadata Only | Reviewed | Not Tested | |
| GetIDsOfNames | Unknown | 5 | 1610678274 | Metadata Only | Reviewed | Not Tested | |
| Invoke | Unknown | 8 | 1610678275 | Metadata Only | Reviewed | Not Tested | |

## Events

| Event | Arguments | DISPID | Implementation | Docs | Tests |
|---|---:|---:|---|---|---|
| -- | -- | -- | Not started | Generated | Not tested |

## Unsupported or deferred behaviour

See the global unsupported index for unimplemented object-model areas.
<!-- END GENERATED MEMBERS -->
