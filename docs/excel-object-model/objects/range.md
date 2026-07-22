# Range

## Summary

The cell and rectangular-value object. The bounded crate slice supports values, explicit A1/R1C1/external address output, and Cells, Item, Offset, Resize, Rows, Columns, Areas, EntireRow, and EntireColumn navigation.

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
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Range` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Capabilities

### Formula

| Capability | Available |
|---|---|
| `a1` | true |
| `dynamic_array` | true |
| `formula2` | true |
| `formula2_r1c1` | true |
| `legacy_array` | true |
| `locale_formula` | true |
| `mixed_values` | true |
| `r1c1` | true |

### Calculation

| Capability | Available |
|---|---|
| `calculate` | true |
| `mark_dirty` | true |

### Auditing and search

| Capability | Available |
|---|---|
| `dependents` | true |
| `find` | true |
| `precedents` | true |
| `replace` | true |
| `special_cells` | true |
| `wrap_safe_iterator` | true |

### Formatting

| Capability | Available |
|---|---|
| `alignment` | true |
| `autofit` | true |
| `borders` | true |
| `dimensions` | true |
| `fill` | true |
| `font` | true |
| `number_format` | true |



## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Areas` | `excel.areas` | Implemented |
| `Borders` | `excel.borders` | Implemented |
| `Cells` | `excel.range` | Implemented |
| `ColumnDifferences` | `excel.range` | Metadata Only |
| `Columns` | `excel.range` | Implemented |
| `CurrentArray` | `excel.range` | Implemented |
| `CurrentRegion` | `excel.range` | Metadata Only |
| `Dependents` | `excel.range` | Implemented |
| `DirectDependents` | `excel.range` | Implemented |
| `DirectPrecedents` | `excel.range` | Implemented |
| `End` | `excel.range` | Metadata Only |
| `EntireColumn` | `excel.range` | Implemented |
| `EntireRow` | `excel.range` | Implemented |
| `Find` | `excel.range` | Implemented |
| `FindNext` | `excel.range` | Implemented |
| `FindPrevious` | `excel.range` | Implemented |
| `Font` | `excel.font` | Implemented |
| `Interior` | `excel.interior` | Implemented |
| `MergeArea` | `excel.range` | Metadata Only |
| `Next` | `excel.range` | Metadata Only |
| `Offset` | `excel.range` | Implemented |
| `Precedents` | `excel.range` | Implemented |
| `Previous` | `excel.range` | Metadata Only |
| `Range` | `excel.range` | Metadata Only |
| `Resize` | `excel.range` | Implemented |
| `RowDifferences` | `excel.range` | Metadata Only |
| `Rows` | `excel.range` | Implemented |
| `SpecialCells` | `excel.range` | Implemented |
| `SpillingToRange` | `excel.range` | Implemented |
| `SpillParent` | `excel.range` | Implemented |
| `Worksheet` | `excel.worksheet` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _NewEnum | PROPERTYGET | Unknown | declared | -4 | Implemented | Reviewed | Live Tested | |
| _Default | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| Value | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 6 | Implemented | Reviewed | Live Tested | |
| Name | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 110 | Metadata Only | Reviewed | Not Tested | |
| Count | PROPERTYGET | i32 | declared | 118 | Implemented | Reviewed | Live Tested | |
| Width | PROPERTYGET | AutomationValue | declared | 122 | Metadata Only | Reviewed | Not Tested | |
| Height | PROPERTYGET | AutomationValue | declared | 123 | Metadata Only | Reviewed | Not Tested | |
| Top | PROPERTYGET | AutomationValue | declared | 126 | Metadata Only | Reviewed | Not Tested | |
| Left | PROPERTYGET | AutomationValue | declared | 127 | Metadata Only | Reviewed | Not Tested | |
| Interior | PROPERTYGET | Interior | declared | 129 | Implemented | Reviewed | Live Tested | |
| Orientation | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 134 | Metadata Only | Reviewed | Not Tested | |
| HorizontalAlignment | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 136 | Implemented | Reviewed | Live Tested | |
| VerticalAlignment | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 137 | Implemented | Reviewed | Live Tested | |
| Text | PROPERTYGET | AutomationValue | declared | 138 | Metadata Only | Reviewed | Not Tested | |
| Font | PROPERTYGET | Font | declared | 146 | Implemented | Reviewed | Live Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Item | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 170 | Implemented | Reviewed | Live Tested | |
| NumberFormat | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 193 | Implemented | Reviewed | Live Tested | |
| Range | PROPERTYGET | Range | declared | 197 | Metadata Only | Reviewed | Not Tested | |
| IndentLevel | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 201 | Metadata Only | Reviewed | Not Tested | |
| MergeCells | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 208 | Metadata Only | Reviewed | Not Tested | |
| ShrinkToFit | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 209 | Metadata Only | Reviewed | Not Tested | |
| Address | PROPERTYGET | String | declared | 236 | Implemented | Reviewed | Live Tested | |
| Cells | PROPERTYGET | Range | declared | 238 | Implemented | Reviewed | Live Tested | |
| Column | PROPERTYGET | i32 | declared | 240 | Implemented | Reviewed | Live Tested | |
| Columns | PROPERTYGET | Range | declared | 241 | Implemented | Reviewed | Live Tested | |
| ColumnWidth | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 242 | Implemented | Reviewed | Live Tested | |
| CurrentRegion | PROPERTYGET | Range | declared | 243 | Metadata Only | Reviewed | Not Tested | |
| EntireColumn | PROPERTYGET | Range | declared | 246 | Implemented | Reviewed | Live Tested | |
| EntireRow | PROPERTYGET | Range | declared | 247 | Implemented | Reviewed | Live Tested | |
| Offset | PROPERTYGET | Range | declared | 254 | Implemented | Reviewed | Live Tested | |
| PageBreak | PROPERTYGET/PROPERTYPUT | i32 | declared | 255 | Metadata Only | Reviewed | Not Tested | |
| Resize | PROPERTYGET | Range | declared | 256 | Implemented | Reviewed | Live Tested | |
| Row | PROPERTYGET | i32 | declared | 257 | Implemented | Reviewed | Live Tested | |
| Rows | PROPERTYGET | Range | declared | 258 | Implemented | Reviewed | Live Tested | |
| Style | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 260 | Metadata Only | Reviewed | Not Tested | |
| Formula | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 261 | Implemented | Reviewed | Live Tested | |
| FormulaHidden | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 262 | Metadata Only | Reviewed | Not Tested | |
| FormulaLocal | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 263 | Implemented | Reviewed | Live Tested | |
| FormulaR1C1 | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 264 | Implemented | Reviewed | Live Tested | |
| FormulaR1C1Local | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 265 | Implemented | Reviewed | Live Tested | |
| HasArray | PROPERTYGET | AutomationValue | declared | 266 | Implemented | Reviewed | Live Tested | |
| HasFormula | PROPERTYGET | AutomationValue | declared | 267 | Implemented | Reviewed | Live Tested | |
| Hidden | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 268 | Metadata Only | Reviewed | Not Tested | |
| Locked | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 269 | Metadata Only | Reviewed | Not Tested | |
| OutlineLevel | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 271 | Metadata Only | Reviewed | Not Tested | |
| RowHeight | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 272 | Implemented | Reviewed | Live Tested | |
| Summary | PROPERTYGET | AutomationValue | declared | 273 | Metadata Only | Reviewed | Not Tested | |
| UseStandardHeight | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 274 | Metadata Only | Reviewed | Not Tested | |
| UseStandardWidth | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 275 | Metadata Only | Reviewed | Not Tested | |
| WrapText | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 276 | Implemented | Reviewed | Live Tested | |
| Worksheet | PROPERTYGET | Worksheet | declared | 348 | Metadata Only | Reviewed | Not Tested | |
| Borders | PROPERTYGET | Borders | declared | 435 | Implemented | Reviewed | Live Tested | |
| AddressLocal | PROPERTYGET | String | declared | 437 | Metadata Only | Reviewed | Not Tested | |
| End | PROPERTYGET | Range | declared | 500 | Metadata Only | Reviewed | Not Tested | |
| CurrentArray | PROPERTYGET | Range | declared | 501 | Implemented | Reviewed | Live Tested | |
| Next | PROPERTYGET | Range | declared | 502 | Metadata Only | Reviewed | Not Tested | |
| Previous | PROPERTYGET | Range | declared | 503 | Metadata Only | Reviewed | Not Tested | |
| PrefixCharacter | PROPERTYGET | AutomationValue | declared | 504 | Metadata Only | Reviewed | Not Tested | |
| Dependents | PROPERTYGET | Range | declared | 543 | Implemented | Reviewed | Live Tested | |
| Precedents | PROPERTYGET | Range | declared | 544 | Implemented | Reviewed | Live Tested | |
| DirectDependents | PROPERTYGET | Range | declared | 545 | Implemented | Reviewed | Live Tested | |
| DirectPrecedents | PROPERTYGET | Range | declared | 546 | Implemented | Reviewed | Live Tested | |
| Areas | PROPERTYGET | Areas | declared | 568 | Implemented | Reviewed | Live Tested | |
| ShowDetail | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 585 | Metadata Only | Reviewed | Not Tested | |
| FormulaArray | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 586 | Implemented | Reviewed | Live Tested | |
| Characters | PROPERTYGET | Characters | declared | 603 | Metadata Only | Reviewed | Not Tested | |
| DisplayFormat | PROPERTYGET | DisplayFormat | declared | 666 | Metadata Only | Reviewed | Not Tested | |
| LocationInTable | PROPERTYGET | XlLocationInTable | declared | 691 | Metadata Only | Reviewed | Not Tested | |
| PivotTable | PROPERTYGET | PivotTable | declared | 716 | Metadata Only | Reviewed | Not Tested | |
| PivotField | PROPERTYGET | PivotField | declared | 731 | Metadata Only | Reviewed | Not Tested | |
| PivotItem | PROPERTYGET | PivotItem | declared | 740 | Metadata Only | Reviewed | Not Tested | |
| Comment | PROPERTYGET | Comment | declared | 910 | Metadata Only | Reviewed | Not Tested | |
| SoundNote | PROPERTYGET | SoundNote | declared | 916 | Metadata Only | Reviewed | Not Tested | |
| ReadingOrder | PROPERTYGET/PROPERTYPUT | i32 | declared | 975 | Metadata Only | Reviewed | Not Tested | |
| AddIndent | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1063 | Metadata Only | Reviewed | Not Tested | |
| NumberFormatLocal | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1097 | Metadata Only | Reviewed | Not Tested | |
| ListHeaderRows | PROPERTYGET | i32 | declared | 1187 | Metadata Only | Reviewed | Not Tested | |
| FormulaLabel | PROPERTYGET/PROPERTYPUT | XlFormulaLabel | declared | 1380 | Metadata Only | Reviewed | Not Tested | |
| MergeArea | PROPERTYGET | Range | declared | 1385 | Metadata Only | Reviewed | Not Tested | |
| QueryTable | PROPERTYGET | QueryTable | declared | 1386 | Metadata Only | Reviewed | Not Tested | |
| Validation | PROPERTYGET | Validation | declared | 1387 | Metadata Only | Reviewed | Not Tested | |
| Value2 | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1388 | Implemented | Reviewed | Live Tested | |
| Phonetic | PROPERTYGET | Phonetic | declared | 1391 | Metadata Only | Reviewed | Not Tested | |
| FormatConditions | PROPERTYGET | FormatConditions | declared | 1392 | Metadata Only | Reviewed | Not Tested | |
| Hyperlinks | PROPERTYGET | Hyperlinks | declared | 1393 | Metadata Only | Reviewed | Not Tested | |
| Formula2 | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1580 | Implemented | Reviewed | Live Tested | |
| Phonetics | PROPERTYGET | Phonetics | declared | 1811 | Metadata Only | Reviewed | Not Tested | |
| ID | PROPERTYGET/PROPERTYPUT | String | declared | 1813 | Metadata Only | Reviewed | Not Tested | |
| PivotCell | PROPERTYGET | PivotCell | declared | 2013 | Metadata Only | Reviewed | Not Tested | |
| Errors | PROPERTYGET | Errors | declared | 2015 | Metadata Only | Reviewed | Not Tested | |
| SmartTags | PROPERTYGET | SmartTags | declared | 2016 | Metadata Only | Reviewed | Not Tested | |
| AllowEdit | PROPERTYGET | bool | declared | 2020 | Metadata Only | Reviewed | Not Tested | |
| MDX | PROPERTYGET | String | declared | 2123 | Metadata Only | Reviewed | Not Tested | |
| ListObject | PROPERTYGET | ListObject | declared | 2257 | Metadata Only | Reviewed | Not Tested | |
| XPath | PROPERTYGET | XPath | declared | 2258 | Metadata Only | Reviewed | Not Tested | |
| ServerActions | PROPERTYGET | Actions | declared | 2491 | Metadata Only | Reviewed | Not Tested | |
| CountLarge | PROPERTYGET | AutomationValue | declared | 2499 | Metadata Only | Reviewed | Not Tested | |
| SparklineGroups | PROPERTYGET | SparklineGroups | declared | 2853 | Metadata Only | Reviewed | Not Tested | |
| CommentThreaded | PROPERTYGET | CommentThreaded | declared | 3281 | Metadata Only | Reviewed | Not Tested | |
| LinkedDataTypeState | PROPERTYGET | AutomationValue | declared | 3291 | Metadata Only | Reviewed | Not Tested | |
| HasSpill | PROPERTYGET | AutomationValue | declared | 3295 | Implemented | Reviewed | Live Tested | |
| SpillingToRange | PROPERTYGET | Range | declared | 3296 | Implemented | Reviewed | Live Tested | |
| SpillParent | PROPERTYGET | Range | declared | 3297 | Implemented | Reviewed | Live Tested | |
| Formula2Local | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 3300 | Metadata Only | Reviewed | Not Tested | |
| Formula2R1C1 | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 3301 | Implemented | Reviewed | Live Tested | |
| Formula2R1C1Local | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 3302 | Metadata Only | Reviewed | Not Tested | |
| SavedAsArray | PROPERTYGET | AutomationValue | declared | 3303 | Metadata Only | Reviewed | Not Tested | |
| HasRichDataType | PROPERTYGET | AutomationValue | declared | 3326 | Metadata Only | Reviewed | Not Tested | |
| CellControl | PROPERTYGET | CellControl | declared | 3411 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Group | AutomationValue | 4 | declared | 46 | Metadata Only | Reviewed | Not Tested | |
| Clear | AutomationValue | 0 | declared | 111 | Metadata Only | Reviewed | Not Tested | |
| ClearFormats | AutomationValue | 0 | declared | 112 | Metadata Only | Reviewed | Not Tested | |
| _ClearContents | AutomationValue | 0 | declared | 113 | Implemented | Reviewed | Live Tested | |
| AutoFormat | AutomationValue | 7 | declared | 114 | Metadata Only | Reviewed | Not Tested | |
| Delete | AutomationValue | 1 | declared | 117 | Metadata Only | Reviewed | Not Tested | |
| CopyPicture | AutomationValue | 2 | declared | 213 | Metadata Only | Reviewed | Not Tested | |
| _Replace | bool | 8 | declared | 226 | Metadata Only | Reviewed | Not Tested | |
| Select | AutomationValue | 0 | declared | 235 | Metadata Only | Reviewed | Not Tested | |
| AutoFit | AutomationValue | 0 | declared | 237 | Implemented | Reviewed | Live Tested | |
| ClearNotes | AutomationValue | 0 | declared | 239 | Metadata Only | Reviewed | Not Tested | |
| Ungroup | AutomationValue | 0 | declared | 244 | Metadata Only | Reviewed | Not Tested | |
| DialogBox | AutomationValue | 0 | declared | 245 | Metadata Only | Reviewed | Not Tested | |
| FillDown | AutomationValue | 0 | declared | 248 | Metadata Only | Reviewed | Not Tested | |
| FillLeft | AutomationValue | 0 | declared | 249 | Metadata Only | Reviewed | Not Tested | |
| FillRight | AutomationValue | 0 | declared | 250 | Metadata Only | Reviewed | Not Tested | |
| FillUp | AutomationValue | 0 | declared | 251 | Metadata Only | Reviewed | Not Tested | |
| Insert | AutomationValue | 2 | declared | 252 | Metadata Only | Reviewed | Not Tested | |
| ListNames | AutomationValue | 0 | declared | 253 | Metadata Only | Reviewed | Not Tested | |
| Run | AutomationValue | 30 | declared | 259 | Metadata Only | Reviewed | Not Tested | |
| Calculate | AutomationValue | 0 | declared | 279 | Implemented | Reviewed | Live Tested | |
| PrintPreview | AutomationValue | 1 | declared | 281 | Metadata Only | Reviewed | Not Tested | |
| Activate | AutomationValue | 0 | declared | 304 | Metadata Only | Reviewed | Not Tested | |
| Find | Range | 9 | declared | 398 | Implemented | Reviewed | Live Tested | |
| FindNext | Range | 1 | declared | 399 | Implemented | Reviewed | Live Tested | |
| FindPrevious | Range | 1 | declared | 400 | Implemented | Reviewed | Live Tested | |
| SpecialCells | Range | 2 | declared | 410 | Implemented | Reviewed | Live Tested | |
| ApplyNames | AutomationValue | 7 | declared | 441 | Metadata Only | Reviewed | Not Tested | |
| ApplyOutlineStyles | AutomationValue | 0 | declared | 448 | Metadata Only | Reviewed | Not Tested | |
| AutoFill | AutomationValue | 2 | declared | 449 | Metadata Only | Reviewed | Not Tested | |
| CreateNames | AutomationValue | 4 | declared | 457 | Metadata Only | Reviewed | Not Tested | |
| CreatePublisher | AutomationValue | 6 | declared | 458 | Metadata Only | Reviewed | Not Tested | |
| DataSeries | AutomationValue | 6 | declared | 464 | Metadata Only | Reviewed | Not Tested | |
| GoalSeek | bool | 2 | declared | 472 | Metadata Only | Reviewed | Not Tested | |
| Parse | AutomationValue | 2 | declared | 477 | Metadata Only | Reviewed | Not Tested | |
| SubscribeTo | AutomationValue | 2 | declared | 481 | Metadata Only | Reviewed | Not Tested | |
| Consolidate | AutomationValue | 5 | declared | 482 | Metadata Only | Reviewed | Not Tested | |
| Justify | AutomationValue | 0 | declared | 495 | Metadata Only | Reviewed | Not Tested | |
| Show | AutomationValue | 0 | declared | 496 | Metadata Only | Reviewed | Not Tested | |
| Table | AutomationValue | 2 | declared | 497 | Metadata Only | Reviewed | Not Tested | |
| CheckSpelling | AutomationValue | 4 | declared | 505 | Metadata Only | Reviewed | Not Tested | |
| ColumnDifferences | Range | 1 | declared | 510 | Metadata Only | Reviewed | Not Tested | |
| RowDifferences | Range | 1 | declared | 511 | Metadata Only | Reviewed | Not Tested | |
| Copy | AutomationValue | 1 | declared | 551 | Metadata Only | Reviewed | Not Tested | |
| Merge | Unknown | 1 | declared | 564 | Metadata Only | Reviewed | Not Tested | |
| Cut | AutomationValue | 1 | declared | 565 | Metadata Only | Reviewed | Not Tested | |
| FunctionWizard | AutomationValue | 0 | declared | 571 | Metadata Only | Reviewed | Not Tested | |
| _AutoFilter | AutomationValue | 5 | declared | 793 | Metadata Only | Reviewed | Not Tested | |
| AdvancedFilter | AutomationValue | 4 | declared | 876 | Metadata Only | Reviewed | Not Tested | |
| ShowDependents | AutomationValue | 1 | declared | 877 | Metadata Only | Reviewed | Not Tested | |
| ShowErrors | AutomationValue | 0 | declared | 878 | Metadata Only | Reviewed | Not Tested | |
| ShowPrecedents | AutomationValue | 1 | declared | 879 | Metadata Only | Reviewed | Not Tested | |
| _Sort | AutomationValue | 15 | declared | 880 | Metadata Only | Reviewed | Not Tested | |
| SortSpecial | AutomationValue | 15 | declared | 881 | Metadata Only | Reviewed | Not Tested | |
| Subtotal | AutomationValue | 6 | declared | 882 | Metadata Only | Reviewed | Not Tested | |
| RemoveSubtotal | AutomationValue | 0 | declared | 883 | Metadata Only | Reviewed | Not Tested | |
| __PrintOut | AutomationValue | 7 | declared | 905 | Metadata Only | Reviewed | Not Tested | |
| _PasteSpecial | AutomationValue | 4 | declared | 1027 | Metadata Only | Reviewed | Not Tested | |
| NavigateArrow | AutomationValue | 3 | declared | 1032 | Metadata Only | Reviewed | Not Tested | |
| AutoOutline | AutomationValue | 0 | declared | 1036 | Metadata Only | Reviewed | Not Tested | |
| ClearOutline | AutomationValue | 0 | declared | 1037 | Metadata Only | Reviewed | Not Tested | |
| TextToColumns | AutomationValue | 14 | declared | 1040 | Metadata Only | Reviewed | Not Tested | |
| _BorderAround | AutomationValue | 4 | declared | 1067 | Metadata Only | Reviewed | Not Tested | |
| NoteText | String | 3 | declared | 1127 | Metadata Only | Reviewed | Not Tested | |
| EditionOptions | AutomationValue | 7 | declared | 1131 | Metadata Only | Reviewed | Not Tested | |
| CopyFromRecordset | i32 | 3 | declared | 1152 | Metadata Only | Reviewed | Not Tested | |
| AutoComplete | String | 1 | declared | 1185 | Metadata Only | Reviewed | Not Tested | |
| InsertIndent | Unknown | 1 | declared | 1381 | Metadata Only | Reviewed | Not Tested | |
| UnMerge | Unknown | 0 | declared | 1384 | Metadata Only | Reviewed | Not Tested | |
| AddComment | Comment | 1 | declared | 1389 | Metadata Only | Reviewed | Not Tested | |
| ClearComments | Unknown | 0 | declared | 1390 | Metadata Only | Reviewed | Not Tested | |
| _PrintOut | AutomationValue | 8 | declared | 1772 | Metadata Only | Reviewed | Not Tested | |
| SetPhonetic | Unknown | 0 | declared | 1812 | Metadata Only | Reviewed | Not Tested | |
| PasteSpecial | AutomationValue | 4 | declared | 1928 | Metadata Only | Reviewed | Not Tested | |
| Dirty | Unknown | 0 | declared | 2014 | Implemented | Reviewed | Live Tested | |
| Speak | Unknown | 2 | declared | 2017 | Metadata Only | Reviewed | Not Tested | |
| PrintOut | AutomationValue | 8 | declared | 2361 | Metadata Only | Reviewed | Not Tested | |
| CalculateRowMajorOrder | AutomationValue | 0 | declared | 2364 | Metadata Only | Reviewed | Not Tested | |
| RemoveDuplicates | Unknown | 2 | declared | 2492 | Metadata Only | Reviewed | Not Tested | |
| _ExportAsFixedFormat | Unknown | 9 | declared | 2493 | Metadata Only | Reviewed | Not Tested | |
| BorderAround | AutomationValue | 5 | declared | 2771 | Metadata Only | Reviewed | Not Tested | |
| ClearHyperlinks | Unknown | 0 | declared | 2854 | Metadata Only | Reviewed | Not Tested | |
| AllocateChanges | Unknown | 0 | declared | 2855 | Metadata Only | Reviewed | Not Tested | |
| DiscardChanges | Unknown | 0 | declared | 2856 | Metadata Only | Reviewed | Not Tested | |
| FlashFill | Unknown | 0 | declared | 2996 | Metadata Only | Reviewed | Not Tested | |
| ExportAsFixedFormat | Unknown | 10 | declared | 3175 | Metadata Only | Reviewed | Not Tested | |
| ShowCard | Unknown | 0 | declared | 3274 | Metadata Only | Reviewed | Not Tested | |
| AddCommentThreaded | CommentThreaded | 1 | declared | 3280 | Metadata Only | Reviewed | Not Tested | |
| Sort | AutomationValue | 16 | declared | 3288 | Metadata Only | Reviewed | Not Tested | |
| AutoFilter | AutomationValue | 6 | declared | 3289 | Metadata Only | Reviewed | Not Tested | |
| ConvertToLinkedDataType | Unknown | 2 | declared | 3290 | Metadata Only | Reviewed | Not Tested | |
| SetCellDataTypeFromCell | Unknown | 1 | declared | 3293 | Metadata Only | Reviewed | Not Tested | |
| DataTypeToText | Unknown | 0 | declared | 3294 | Metadata Only | Reviewed | Not Tested | |
| RefreshLinkedDataType | Unknown | 1 | declared | 3299 | Metadata Only | Reviewed | Not Tested | |
| Replace | bool | 9 | declared | 3305 | Implemented | Reviewed | Live Tested | |
| InsertPictureInCell | Unknown | 1 | declared | 3402 | Metadata Only | Reviewed | Not Tested | |
| PastePictureInCell | Unknown | 0 | declared | 3405 | Metadata Only | Reviewed | Not Tested | |
| PlacePictureOverCells | Unknown | 1 | declared | 3407 | Metadata Only | Reviewed | Not Tested | |
| UpdatePictureInCellAlternativeText | Unknown | 1 | declared | 3410 | Metadata Only | Reviewed | Not Tested | |
| ClearContents | AutomationValue | 1 | declared | 3413 | Metadata Only | Reviewed | Not Tested | |
| RemoveControls | Unknown | 0 | declared | 3414 | Metadata Only | Reviewed | Not Tested | |
| ResetContents | Unknown | 0 | declared | 3415 | Metadata Only | Reviewed | Not Tested | |
| TogglePythonMarshalMode | Unknown | 1 | declared | 3419 | Metadata Only | Reviewed | Not Tested | |
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
