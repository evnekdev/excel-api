# Range

## Summary

The cell and rectangular-value object. It is structurally inventoried; production Range support remains deferred pending the wrapper architecture review.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Range` |
| GUID | `{00020846-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Crate type | `excel_com::Range` |
| Implementation | Metadata Only |
| Documentation | Reviewed |
| Tests | Not Tested |

## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Cells` | `excel.range` | Metadata Only |
| `ColumnDifferences` | `excel.range` | Metadata Only |
| `Columns` | `excel.range` | Metadata Only |
| `CurrentArray` | `excel.range` | Metadata Only |
| `CurrentRegion` | `excel.range` | Metadata Only |
| `Dependents` | `excel.range` | Metadata Only |
| `DirectDependents` | `excel.range` | Metadata Only |
| `DirectPrecedents` | `excel.range` | Metadata Only |
| `End` | `excel.range` | Metadata Only |
| `EntireColumn` | `excel.range` | Metadata Only |
| `EntireRow` | `excel.range` | Metadata Only |
| `Find` | `excel.range` | Metadata Only |
| `FindNext` | `excel.range` | Metadata Only |
| `FindPrevious` | `excel.range` | Metadata Only |
| `MergeArea` | `excel.range` | Metadata Only |
| `Next` | `excel.range` | Metadata Only |
| `Offset` | `excel.range` | Metadata Only |
| `Precedents` | `excel.range` | Metadata Only |
| `Previous` | `excel.range` | Metadata Only |
| `Range` | `excel.range` | Metadata Only |
| `Resize` | `excel.range` | Metadata Only |
| `RowDifferences` | `excel.range` | Metadata Only |
| `Rows` | `excel.range` | Metadata Only |
| `SpecialCells` | `excel.range` | Metadata Only |
| `SpillingToRange` | `excel.range` | Metadata Only |
| `SpillParent` | `excel.range` | Metadata Only |
| `Worksheet` | `excel.worksheet` | Metadata Only |

## Properties

| Property | Access | Type | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---:|---|---|---|---|
| _NewEnum | PROPERTYGET | Unknown | -4 | Metadata Only | Reviewed | Not Tested | |
| _Default | PROPERTYGET/PROPERTYPUT | AutomationValue | 0 | Metadata Only | Reviewed | Not Tested | |
| Value | PROPERTYGET/PROPERTYPUT | AutomationValue | 6 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET/PROPERTYPUT | AutomationValue | 110 | Metadata Only | Reviewed | Not Tested | |
| Count | PROPERTYGET | i32 | 118 | Metadata Only | Reviewed | Not Tested | |
| Width | PROPERTYGET | AutomationValue | 122 | Metadata Only | Reviewed | Not Tested | |
| Height | PROPERTYGET | AutomationValue | 123 | Metadata Only | Reviewed | Not Tested | |
| Top | PROPERTYGET | AutomationValue | 126 | Metadata Only | Reviewed | Not Tested | |
| Left | PROPERTYGET | AutomationValue | 127 | Metadata Only | Reviewed | Not Tested | |
| Interior | PROPERTYGET | Interior | 129 | Metadata Only | Reviewed | Not Tested | |
| Orientation | PROPERTYGET/PROPERTYPUT | AutomationValue | 134 | Metadata Only | Reviewed | Not Tested | |
| HorizontalAlignment | PROPERTYGET/PROPERTYPUT | AutomationValue | 136 | Metadata Only | Reviewed | Not Tested | |
| VerticalAlignment | PROPERTYGET/PROPERTYPUT | AutomationValue | 137 | Metadata Only | Reviewed | Not Tested | |
| Text | PROPERTYGET | AutomationValue | 138 | Metadata Only | Reviewed | Not Tested | |
| Font | PROPERTYGET | Font | 146 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | 150 | Metadata Only | Reviewed | Not Tested | |
| Item | PROPERTYGET/PROPERTYPUT | AutomationValue | 170 | Metadata Only | Reviewed | Not Tested | |
| NumberFormat | PROPERTYGET/PROPERTYPUT | AutomationValue | 193 | Metadata Only | Reviewed | Not Tested | |
| Range | PROPERTYGET | Range | 197 | Metadata Only | Reviewed | Not Tested | |
| IndentLevel | PROPERTYGET/PROPERTYPUT | AutomationValue | 201 | Metadata Only | Reviewed | Not Tested | |
| MergeCells | PROPERTYGET/PROPERTYPUT | AutomationValue | 208 | Metadata Only | Reviewed | Not Tested | |
| ShrinkToFit | PROPERTYGET/PROPERTYPUT | AutomationValue | 209 | Metadata Only | Reviewed | Not Tested | |
| Address | PROPERTYGET | String | 236 | Metadata Only | Reviewed | Not Tested | |
| Cells | PROPERTYGET | Range | 238 | Metadata Only | Reviewed | Not Tested | |
| Column | PROPERTYGET | i32 | 240 | Metadata Only | Reviewed | Not Tested | |
| Columns | PROPERTYGET | Range | 241 | Metadata Only | Reviewed | Not Tested | |
| ColumnWidth | PROPERTYGET/PROPERTYPUT | AutomationValue | 242 | Metadata Only | Reviewed | Not Tested | |
| CurrentRegion | PROPERTYGET | Range | 243 | Metadata Only | Reviewed | Not Tested | |
| EntireColumn | PROPERTYGET | Range | 246 | Metadata Only | Reviewed | Not Tested | |
| EntireRow | PROPERTYGET | Range | 247 | Metadata Only | Reviewed | Not Tested | |
| Offset | PROPERTYGET | Range | 254 | Metadata Only | Reviewed | Not Tested | |
| PageBreak | PROPERTYGET/PROPERTYPUT | i32 | 255 | Metadata Only | Reviewed | Not Tested | |
| Resize | PROPERTYGET | Range | 256 | Metadata Only | Reviewed | Not Tested | |
| Row | PROPERTYGET | i32 | 257 | Metadata Only | Reviewed | Not Tested | |
| Rows | PROPERTYGET | Range | 258 | Metadata Only | Reviewed | Not Tested | |
| Style | PROPERTYGET/PROPERTYPUT | AutomationValue | 260 | Metadata Only | Reviewed | Not Tested | |
| Formula | PROPERTYGET/PROPERTYPUT | AutomationValue | 261 | Metadata Only | Reviewed | Not Tested | |
| FormulaHidden | PROPERTYGET/PROPERTYPUT | AutomationValue | 262 | Metadata Only | Reviewed | Not Tested | |
| FormulaLocal | PROPERTYGET/PROPERTYPUT | AutomationValue | 263 | Metadata Only | Reviewed | Not Tested | |
| FormulaR1C1 | PROPERTYGET/PROPERTYPUT | AutomationValue | 264 | Metadata Only | Reviewed | Not Tested | |
| FormulaR1C1Local | PROPERTYGET/PROPERTYPUT | AutomationValue | 265 | Metadata Only | Reviewed | Not Tested | |
| HasArray | PROPERTYGET | AutomationValue | 266 | Metadata Only | Reviewed | Not Tested | |
| HasFormula | PROPERTYGET | AutomationValue | 267 | Metadata Only | Reviewed | Not Tested | |
| Hidden | PROPERTYGET/PROPERTYPUT | AutomationValue | 268 | Metadata Only | Reviewed | Not Tested | |
| Locked | PROPERTYGET/PROPERTYPUT | AutomationValue | 269 | Metadata Only | Reviewed | Not Tested | |
| OutlineLevel | PROPERTYGET/PROPERTYPUT | AutomationValue | 271 | Metadata Only | Reviewed | Not Tested | |
| RowHeight | PROPERTYGET/PROPERTYPUT | AutomationValue | 272 | Metadata Only | Reviewed | Not Tested | |
| Summary | PROPERTYGET | AutomationValue | 273 | Metadata Only | Reviewed | Not Tested | |
| UseStandardHeight | PROPERTYGET/PROPERTYPUT | AutomationValue | 274 | Metadata Only | Reviewed | Not Tested | |
| UseStandardWidth | PROPERTYGET/PROPERTYPUT | AutomationValue | 275 | Metadata Only | Reviewed | Not Tested | |
| WrapText | PROPERTYGET/PROPERTYPUT | AutomationValue | 276 | Metadata Only | Reviewed | Not Tested | |
| Worksheet | PROPERTYGET | Worksheet | 348 | Metadata Only | Reviewed | Not Tested | |
| Borders | PROPERTYGET | Borders | 435 | Metadata Only | Reviewed | Not Tested | |
| AddressLocal | PROPERTYGET | String | 437 | Metadata Only | Reviewed | Not Tested | |
| End | PROPERTYGET | Range | 500 | Metadata Only | Reviewed | Not Tested | |
| CurrentArray | PROPERTYGET | Range | 501 | Metadata Only | Reviewed | Not Tested | |
| Next | PROPERTYGET | Range | 502 | Metadata Only | Reviewed | Not Tested | |
| Previous | PROPERTYGET | Range | 503 | Metadata Only | Reviewed | Not Tested | |
| PrefixCharacter | PROPERTYGET | AutomationValue | 504 | Metadata Only | Reviewed | Not Tested | |
| Dependents | PROPERTYGET | Range | 543 | Metadata Only | Reviewed | Not Tested | |
| Precedents | PROPERTYGET | Range | 544 | Metadata Only | Reviewed | Not Tested | |
| DirectDependents | PROPERTYGET | Range | 545 | Metadata Only | Reviewed | Not Tested | |
| DirectPrecedents | PROPERTYGET | Range | 546 | Metadata Only | Reviewed | Not Tested | |
| Areas | PROPERTYGET | Areas | 568 | Metadata Only | Reviewed | Not Tested | |
| ShowDetail | PROPERTYGET/PROPERTYPUT | AutomationValue | 585 | Metadata Only | Reviewed | Not Tested | |
| FormulaArray | PROPERTYGET/PROPERTYPUT | AutomationValue | 586 | Metadata Only | Reviewed | Not Tested | |
| Characters | PROPERTYGET | Characters | 603 | Metadata Only | Reviewed | Not Tested | |
| DisplayFormat | PROPERTYGET | DisplayFormat | 666 | Metadata Only | Reviewed | Not Tested | |
| LocationInTable | PROPERTYGET | XlLocationInTable | 691 | Metadata Only | Reviewed | Not Tested | |
| PivotTable | PROPERTYGET | PivotTable | 716 | Metadata Only | Reviewed | Not Tested | |
| PivotField | PROPERTYGET | PivotField | 731 | Metadata Only | Reviewed | Not Tested | |
| PivotItem | PROPERTYGET | PivotItem | 740 | Metadata Only | Reviewed | Not Tested | |
| Comment | PROPERTYGET | Comment | 910 | Metadata Only | Reviewed | Not Tested | |
| SoundNote | PROPERTYGET | SoundNote | 916 | Metadata Only | Reviewed | Not Tested | |
| ReadingOrder | PROPERTYGET/PROPERTYPUT | i32 | 975 | Metadata Only | Reviewed | Not Tested | |
| AddIndent | PROPERTYGET/PROPERTYPUT | AutomationValue | 1063 | Metadata Only | Reviewed | Not Tested | |
| NumberFormatLocal | PROPERTYGET/PROPERTYPUT | AutomationValue | 1097 | Metadata Only | Reviewed | Not Tested | |
| ListHeaderRows | PROPERTYGET | i32 | 1187 | Metadata Only | Reviewed | Not Tested | |
| FormulaLabel | PROPERTYGET/PROPERTYPUT | XlFormulaLabel | 1380 | Metadata Only | Reviewed | Not Tested | |
| MergeArea | PROPERTYGET | Range | 1385 | Metadata Only | Reviewed | Not Tested | |
| QueryTable | PROPERTYGET | QueryTable | 1386 | Metadata Only | Reviewed | Not Tested | |
| Validation | PROPERTYGET | Validation | 1387 | Metadata Only | Reviewed | Not Tested | |
| Value2 | PROPERTYGET/PROPERTYPUT | AutomationValue | 1388 | Metadata Only | Reviewed | Not Tested | |
| Phonetic | PROPERTYGET | Phonetic | 1391 | Metadata Only | Reviewed | Not Tested | |
| FormatConditions | PROPERTYGET | FormatConditions | 1392 | Metadata Only | Reviewed | Not Tested | |
| Hyperlinks | PROPERTYGET | Hyperlinks | 1393 | Metadata Only | Reviewed | Not Tested | |
| Formula2 | PROPERTYGET/PROPERTYPUT | AutomationValue | 1580 | Metadata Only | Reviewed | Not Tested | |
| Phonetics | PROPERTYGET | Phonetics | 1811 | Metadata Only | Reviewed | Not Tested | |
| ID | PROPERTYGET/PROPERTYPUT | String | 1813 | Metadata Only | Reviewed | Not Tested | |
| PivotCell | PROPERTYGET | PivotCell | 2013 | Metadata Only | Reviewed | Not Tested | |
| Errors | PROPERTYGET | Errors | 2015 | Metadata Only | Reviewed | Not Tested | |
| SmartTags | PROPERTYGET | SmartTags | 2016 | Metadata Only | Reviewed | Not Tested | |
| AllowEdit | PROPERTYGET | bool | 2020 | Metadata Only | Reviewed | Not Tested | |
| MDX | PROPERTYGET | String | 2123 | Metadata Only | Reviewed | Not Tested | |
| ListObject | PROPERTYGET | ListObject | 2257 | Metadata Only | Reviewed | Not Tested | |
| XPath | PROPERTYGET | XPath | 2258 | Metadata Only | Reviewed | Not Tested | |
| ServerActions | PROPERTYGET | Actions | 2491 | Metadata Only | Reviewed | Not Tested | |
| CountLarge | PROPERTYGET | AutomationValue | 2499 | Metadata Only | Reviewed | Not Tested | |
| SparklineGroups | PROPERTYGET | SparklineGroups | 2853 | Metadata Only | Reviewed | Not Tested | |
| CommentThreaded | PROPERTYGET | CommentThreaded | 3281 | Metadata Only | Reviewed | Not Tested | |
| LinkedDataTypeState | PROPERTYGET | AutomationValue | 3291 | Metadata Only | Reviewed | Not Tested | |
| HasSpill | PROPERTYGET | AutomationValue | 3295 | Metadata Only | Reviewed | Not Tested | |
| SpillingToRange | PROPERTYGET | Range | 3296 | Metadata Only | Reviewed | Not Tested | |
| SpillParent | PROPERTYGET | Range | 3297 | Metadata Only | Reviewed | Not Tested | |
| Formula2Local | PROPERTYGET/PROPERTYPUT | AutomationValue | 3300 | Metadata Only | Reviewed | Not Tested | |
| Formula2R1C1 | PROPERTYGET/PROPERTYPUT | AutomationValue | 3301 | Metadata Only | Reviewed | Not Tested | |
| Formula2R1C1Local | PROPERTYGET/PROPERTYPUT | AutomationValue | 3302 | Metadata Only | Reviewed | Not Tested | |
| SavedAsArray | PROPERTYGET | AutomationValue | 3303 | Metadata Only | Reviewed | Not Tested | |
| HasRichDataType | PROPERTYGET | AutomationValue | 3326 | Metadata Only | Reviewed | Not Tested | |
| CellControl | PROPERTYGET | CellControl | 3411 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---:|---|---|---|---|
| Group | AutomationValue | 4 | 46 | Metadata Only | Reviewed | Not Tested | |
| Clear | AutomationValue | 0 | 111 | Metadata Only | Reviewed | Not Tested | |
| ClearFormats | AutomationValue | 0 | 112 | Metadata Only | Reviewed | Not Tested | |
| _ClearContents | AutomationValue | 0 | 113 | Metadata Only | Reviewed | Not Tested | |
| AutoFormat | AutomationValue | 7 | 114 | Metadata Only | Reviewed | Not Tested | |
| Delete | AutomationValue | 1 | 117 | Metadata Only | Reviewed | Not Tested | |
| CopyPicture | AutomationValue | 2 | 213 | Metadata Only | Reviewed | Not Tested | |
| _Replace | bool | 8 | 226 | Metadata Only | Reviewed | Not Tested | |
| Select | AutomationValue | 0 | 235 | Metadata Only | Reviewed | Not Tested | |
| AutoFit | AutomationValue | 0 | 237 | Metadata Only | Reviewed | Not Tested | |
| ClearNotes | AutomationValue | 0 | 239 | Metadata Only | Reviewed | Not Tested | |
| Ungroup | AutomationValue | 0 | 244 | Metadata Only | Reviewed | Not Tested | |
| DialogBox | AutomationValue | 0 | 245 | Metadata Only | Reviewed | Not Tested | |
| FillDown | AutomationValue | 0 | 248 | Metadata Only | Reviewed | Not Tested | |
| FillLeft | AutomationValue | 0 | 249 | Metadata Only | Reviewed | Not Tested | |
| FillRight | AutomationValue | 0 | 250 | Metadata Only | Reviewed | Not Tested | |
| FillUp | AutomationValue | 0 | 251 | Metadata Only | Reviewed | Not Tested | |
| Insert | AutomationValue | 2 | 252 | Metadata Only | Reviewed | Not Tested | |
| ListNames | AutomationValue | 0 | 253 | Metadata Only | Reviewed | Not Tested | |
| Run | AutomationValue | 30 | 259 | Metadata Only | Reviewed | Not Tested | |
| Calculate | AutomationValue | 0 | 279 | Metadata Only | Reviewed | Not Tested | |
| PrintPreview | AutomationValue | 1 | 281 | Metadata Only | Reviewed | Not Tested | |
| Activate | AutomationValue | 0 | 304 | Metadata Only | Reviewed | Not Tested | |
| Find | Range | 9 | 398 | Metadata Only | Reviewed | Not Tested | |
| FindNext | Range | 1 | 399 | Metadata Only | Reviewed | Not Tested | |
| FindPrevious | Range | 1 | 400 | Metadata Only | Reviewed | Not Tested | |
| SpecialCells | Range | 2 | 410 | Metadata Only | Reviewed | Not Tested | |
| ApplyNames | AutomationValue | 7 | 441 | Metadata Only | Reviewed | Not Tested | |
| ApplyOutlineStyles | AutomationValue | 0 | 448 | Metadata Only | Reviewed | Not Tested | |
| AutoFill | AutomationValue | 2 | 449 | Metadata Only | Reviewed | Not Tested | |
| CreateNames | AutomationValue | 4 | 457 | Metadata Only | Reviewed | Not Tested | |
| CreatePublisher | AutomationValue | 6 | 458 | Metadata Only | Reviewed | Not Tested | |
| DataSeries | AutomationValue | 6 | 464 | Metadata Only | Reviewed | Not Tested | |
| GoalSeek | bool | 2 | 472 | Metadata Only | Reviewed | Not Tested | |
| Parse | AutomationValue | 2 | 477 | Metadata Only | Reviewed | Not Tested | |
| SubscribeTo | AutomationValue | 2 | 481 | Metadata Only | Reviewed | Not Tested | |
| Consolidate | AutomationValue | 5 | 482 | Metadata Only | Reviewed | Not Tested | |
| Justify | AutomationValue | 0 | 495 | Metadata Only | Reviewed | Not Tested | |
| Show | AutomationValue | 0 | 496 | Metadata Only | Reviewed | Not Tested | |
| Table | AutomationValue | 2 | 497 | Metadata Only | Reviewed | Not Tested | |
| CheckSpelling | AutomationValue | 4 | 505 | Metadata Only | Reviewed | Not Tested | |
| ColumnDifferences | Range | 1 | 510 | Metadata Only | Reviewed | Not Tested | |
| RowDifferences | Range | 1 | 511 | Metadata Only | Reviewed | Not Tested | |
| Copy | AutomationValue | 1 | 551 | Metadata Only | Reviewed | Not Tested | |
| Merge | Unknown | 1 | 564 | Metadata Only | Reviewed | Not Tested | |
| Cut | AutomationValue | 1 | 565 | Metadata Only | Reviewed | Not Tested | |
| FunctionWizard | AutomationValue | 0 | 571 | Metadata Only | Reviewed | Not Tested | |
| _AutoFilter | AutomationValue | 5 | 793 | Metadata Only | Reviewed | Not Tested | |
| AdvancedFilter | AutomationValue | 4 | 876 | Metadata Only | Reviewed | Not Tested | |
| ShowDependents | AutomationValue | 1 | 877 | Metadata Only | Reviewed | Not Tested | |
| ShowErrors | AutomationValue | 0 | 878 | Metadata Only | Reviewed | Not Tested | |
| ShowPrecedents | AutomationValue | 1 | 879 | Metadata Only | Reviewed | Not Tested | |
| _Sort | AutomationValue | 15 | 880 | Metadata Only | Reviewed | Not Tested | |
| SortSpecial | AutomationValue | 15 | 881 | Metadata Only | Reviewed | Not Tested | |
| Subtotal | AutomationValue | 6 | 882 | Metadata Only | Reviewed | Not Tested | |
| RemoveSubtotal | AutomationValue | 0 | 883 | Metadata Only | Reviewed | Not Tested | |
| __PrintOut | AutomationValue | 7 | 905 | Metadata Only | Reviewed | Not Tested | |
| _PasteSpecial | AutomationValue | 4 | 1027 | Metadata Only | Reviewed | Not Tested | |
| NavigateArrow | AutomationValue | 3 | 1032 | Metadata Only | Reviewed | Not Tested | |
| AutoOutline | AutomationValue | 0 | 1036 | Metadata Only | Reviewed | Not Tested | |
| ClearOutline | AutomationValue | 0 | 1037 | Metadata Only | Reviewed | Not Tested | |
| TextToColumns | AutomationValue | 14 | 1040 | Metadata Only | Reviewed | Not Tested | |
| _BorderAround | AutomationValue | 4 | 1067 | Metadata Only | Reviewed | Not Tested | |
| NoteText | String | 3 | 1127 | Metadata Only | Reviewed | Not Tested | |
| EditionOptions | AutomationValue | 7 | 1131 | Metadata Only | Reviewed | Not Tested | |
| CopyFromRecordset | i32 | 3 | 1152 | Metadata Only | Reviewed | Not Tested | |
| AutoComplete | String | 1 | 1185 | Metadata Only | Reviewed | Not Tested | |
| InsertIndent | Unknown | 1 | 1381 | Metadata Only | Reviewed | Not Tested | |
| UnMerge | Unknown | 0 | 1384 | Metadata Only | Reviewed | Not Tested | |
| AddComment | Comment | 1 | 1389 | Metadata Only | Reviewed | Not Tested | |
| ClearComments | Unknown | 0 | 1390 | Metadata Only | Reviewed | Not Tested | |
| _PrintOut | AutomationValue | 8 | 1772 | Metadata Only | Reviewed | Not Tested | |
| SetPhonetic | Unknown | 0 | 1812 | Metadata Only | Reviewed | Not Tested | |
| PasteSpecial | AutomationValue | 4 | 1928 | Metadata Only | Reviewed | Not Tested | |
| Dirty | Unknown | 0 | 2014 | Metadata Only | Reviewed | Not Tested | |
| Speak | Unknown | 2 | 2017 | Metadata Only | Reviewed | Not Tested | |
| PrintOut | AutomationValue | 8 | 2361 | Metadata Only | Reviewed | Not Tested | |
| CalculateRowMajorOrder | AutomationValue | 0 | 2364 | Metadata Only | Reviewed | Not Tested | |
| RemoveDuplicates | Unknown | 2 | 2492 | Metadata Only | Reviewed | Not Tested | |
| _ExportAsFixedFormat | Unknown | 9 | 2493 | Metadata Only | Reviewed | Not Tested | |
| BorderAround | AutomationValue | 5 | 2771 | Metadata Only | Reviewed | Not Tested | |
| ClearHyperlinks | Unknown | 0 | 2854 | Metadata Only | Reviewed | Not Tested | |
| AllocateChanges | Unknown | 0 | 2855 | Metadata Only | Reviewed | Not Tested | |
| DiscardChanges | Unknown | 0 | 2856 | Metadata Only | Reviewed | Not Tested | |
| FlashFill | Unknown | 0 | 2996 | Metadata Only | Reviewed | Not Tested | |
| ExportAsFixedFormat | Unknown | 10 | 3175 | Metadata Only | Reviewed | Not Tested | |
| ShowCard | Unknown | 0 | 3274 | Metadata Only | Reviewed | Not Tested | |
| AddCommentThreaded | CommentThreaded | 1 | 3280 | Metadata Only | Reviewed | Not Tested | |
| Sort | AutomationValue | 16 | 3288 | Metadata Only | Reviewed | Not Tested | |
| AutoFilter | AutomationValue | 6 | 3289 | Metadata Only | Reviewed | Not Tested | |
| ConvertToLinkedDataType | Unknown | 2 | 3290 | Metadata Only | Reviewed | Not Tested | |
| SetCellDataTypeFromCell | Unknown | 1 | 3293 | Metadata Only | Reviewed | Not Tested | |
| DataTypeToText | Unknown | 0 | 3294 | Metadata Only | Reviewed | Not Tested | |
| RefreshLinkedDataType | Unknown | 1 | 3299 | Metadata Only | Reviewed | Not Tested | |
| Replace | bool | 9 | 3305 | Metadata Only | Reviewed | Not Tested | |
| InsertPictureInCell | Unknown | 1 | 3402 | Metadata Only | Reviewed | Not Tested | |
| PastePictureInCell | Unknown | 0 | 3405 | Metadata Only | Reviewed | Not Tested | |
| PlacePictureOverCells | Unknown | 1 | 3407 | Metadata Only | Reviewed | Not Tested | |
| UpdatePictureInCellAlternativeText | Unknown | 1 | 3410 | Metadata Only | Reviewed | Not Tested | |
| ClearContents | AutomationValue | 1 | 3413 | Metadata Only | Reviewed | Not Tested | |
| RemoveControls | Unknown | 0 | 3414 | Metadata Only | Reviewed | Not Tested | |
| ResetContents | Unknown | 0 | 3415 | Metadata Only | Reviewed | Not Tested | |
| TogglePythonMarshalMode | Unknown | 1 | 3419 | Metadata Only | Reviewed | Not Tested | |
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
