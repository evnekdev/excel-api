# Application

## Summary

The root Automation object for a locally activated Excel instance. The bounded crate slice includes lifecycle, workbook navigation, reference-style guarding, Excel-backed formula conversion, and separately typed scalar/Range evaluation.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `_Application` |
| GUID | `{000208d5-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4160 |
| Crate type | `excel_com::Application` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Capabilities

### Calculation

| Capability | Available |
|---|---|
| `before_save` | true |
| `calculate` | true |
| `full` | true |
| `full_rebuild` | true |
| `mode` | true |
| `state` | true |
| `version` | true |



## Relationships

| Relationship | Target | Status |
|---|---|---|
| `ActiveCell` | `excel.range` | Implemented |
| `ActiveWorkbook` | `excel.workbook` | Implemented |
| `Application` | `excel.application` | Metadata Only |
| `Cells` | `excel.range` | Metadata Only |
| `Columns` | `excel.range` | Metadata Only |
| `Intersect` | `excel.range` | Metadata Only |
| `NextLetter` | `excel.workbook` | Metadata Only |
| `Parent` | `excel.application` | Metadata Only |
| `Range` | `excel.range` | Metadata Only |
| `Rows` | `excel.range` | Metadata Only |
| `ThisCell` | `excel.range` | Metadata Only |
| `ThisWorkbook` | `excel.workbook` | Metadata Only |
| `Union` | `excel.range` | Implemented |
| `Workbooks` | `excel.workbooks` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _Default | PROPERTYGET | String | declared | 0 | Metadata Only | Reviewed | Not Tested | |
| UILanguage | PROPERTYGET/PROPERTYPUT | i32 | declared | 2 | Metadata Only | Reviewed | Not Tested | |
| Value | PROPERTYGET | String | declared | 6 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET | String | declared | 110 | Metadata Only | Reviewed | Not Tested | |
| Charts | PROPERTYGET | Sheets | declared | 121 | Metadata Only | Reviewed | Not Tested | |
| Width | PROPERTYGET/PROPERTYPUT | f64 | declared | 122 | Metadata Only | Reviewed | Not Tested | |
| Height | PROPERTYGET/PROPERTYPUT | f64 | declared | 123 | Metadata Only | Reviewed | Not Tested | |
| Top | PROPERTYGET/PROPERTYPUT | f64 | declared | 126 | Metadata Only | Reviewed | Not Tested | |
| Left | PROPERTYGET/PROPERTYPUT | f64 | declared | 127 | Metadata Only | Reviewed | Not Tested | |
| Caption | PROPERTYGET/PROPERTYPUT | String | declared | 139 | Metadata Only | Reviewed | Not Tested | |
| Selection | PROPERTYGET | Object | declared | 147 | Implemented | Reviewed | Blocked | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Application | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| ActiveChart | PROPERTYGET | Chart | declared | 183 | Metadata Only | Reviewed | Not Tested | |
| Range | PROPERTYGET | Range | declared | 197 | Metadata Only | Reviewed | Not Tested | |
| DefaultSheetDirection | PROPERTYGET/PROPERTYPUT | i32 | declared | 229 | Metadata Only | Reviewed | Not Tested | |
| CursorMovement | PROPERTYGET/PROPERTYPUT | i32 | declared | 232 | Metadata Only | Reviewed | Not Tested | |
| ControlCharacters | PROPERTYGET/PROPERTYPUT | bool | declared | 233 | Metadata Only | Reviewed | Not Tested | |
| Cells | PROPERTYGET | Range | declared | 238 | Metadata Only | Reviewed | Not Tested | |
| Columns | PROPERTYGET | Range | declared | 241 | Metadata Only | Reviewed | Not Tested | |
| Rows | PROPERTYGET | Range | declared | 258 | Metadata Only | Reviewed | Not Tested | |
| Path | PROPERTYGET | String | declared | 291 | Metadata Only | Reviewed | Not Tested | |
| ActiveCell | PROPERTYGET | Range | declared | 305 | Implemented | Reviewed | Blocked | |
| ActivePrinter | PROPERTYGET/PROPERTYPUT | String | declared | 306 | Metadata Only | Reviewed | Not Tested | |
| ActiveSheet | PROPERTYGET | Object | declared | 307 | Implemented | Reviewed | Blocked | |
| ActiveWorkbook | PROPERTYGET | Workbook | declared | 308 | Implemented | Reviewed | Blocked | |
| TransitionMenuKey | PROPERTYGET/PROPERTYPUT | String | declared | 310 | Metadata Only | Reviewed | Not Tested | |
| TransitionMenuKeyAction | PROPERTYGET/PROPERTYPUT | i32 | declared | 311 | Metadata Only | Reviewed | Not Tested | |
| TransitionNavigKeys | PROPERTYGET/PROPERTYPUT | bool | declared | 312 | Metadata Only | Reviewed | Not Tested | |
| AltStartupPath | PROPERTYGET/PROPERTYPUT | String | declared | 313 | Metadata Only | Reviewed | Not Tested | |
| Build | PROPERTYGET | i32 | declared | 314 | Metadata Only | Reviewed | Not Tested | |
| CalculateBeforeSave | PROPERTYGET/PROPERTYPUT | bool | declared | 315 | Implemented | Reviewed | Live Tested | |
| Calculation | PROPERTYGET/PROPERTYPUT | XlCalculation | declared | 316 | Implemented | Reviewed | Live Tested | |
| Caller | PROPERTYGET | AutomationValue | declared | 317 | Metadata Only | Reviewed | Not Tested | |
| CanPlaySounds | PROPERTYGET | bool | declared | 318 | Metadata Only | Reviewed | Not Tested | |
| CanRecordSounds | PROPERTYGET | bool | declared | 319 | Metadata Only | Reviewed | Not Tested | |
| CellDragAndDrop | PROPERTYGET/PROPERTYPUT | bool | declared | 320 | Metadata Only | Reviewed | Not Tested | |
| ClipboardFormats | PROPERTYGET | AutomationValue | declared | 321 | Metadata Only | Reviewed | Not Tested | |
| DisplayClipboardWindow | PROPERTYGET/PROPERTYPUT | bool | declared | 322 | Metadata Only | Reviewed | Not Tested | |
| CommandUnderlines | PROPERTYGET/PROPERTYPUT | XlCommandUnderlines | declared | 323 | Metadata Only | Reviewed | Not Tested | |
| ConstrainNumeric | PROPERTYGET/PROPERTYPUT | bool | declared | 324 | Metadata Only | Reviewed | Not Tested | |
| CutCopyMode | PROPERTYGET/PROPERTYPUT | XlCutCopyMode | declared | 330 | Implemented | Reviewed | Live Tested | |
| DataEntryMode | PROPERTYGET/PROPERTYPUT | i32 | declared | 331 | Metadata Only | Reviewed | Not Tested | |
| DDEAppReturnCode | PROPERTYGET | i32 | declared | 332 | Metadata Only | Reviewed | Not Tested | |
| DisplayAlerts | PROPERTYGET/PROPERTYPUT | bool | declared | 343 | Implemented | Reviewed | Live Tested | |
| DisplayFormulaBar | PROPERTYGET/PROPERTYPUT | bool | declared | 344 | Metadata Only | Reviewed | Not Tested | |
| DisplayNoteIndicator | PROPERTYGET/PROPERTYPUT | bool | declared | 345 | Metadata Only | Reviewed | Not Tested | |
| DisplayScrollBars | PROPERTYGET/PROPERTYPUT | bool | declared | 346 | Metadata Only | Reviewed | Not Tested | |
| DisplayStatusBar | PROPERTYGET/PROPERTYPUT | bool | declared | 347 | Metadata Only | Reviewed | Not Tested | |
| FixedDecimal | PROPERTYGET/PROPERTYPUT | bool | declared | 351 | Metadata Only | Reviewed | Not Tested | |
| FixedDecimalPlaces | PROPERTYGET/PROPERTYPUT | i32 | declared | 352 | Metadata Only | Reviewed | Not Tested | |
| IgnoreRemoteRequests | PROPERTYGET/PROPERTYPUT | bool | declared | 356 | Metadata Only | Reviewed | Not Tested | |
| Interactive | PROPERTYGET/PROPERTYPUT | bool | declared | 361 | Implemented | Reviewed | Live Tested | |
| International | PROPERTYGET | AutomationValue | declared | 362 | Metadata Only | Reviewed | Not Tested | |
| Iteration | PROPERTYGET/PROPERTYPUT | bool | declared | 363 | Metadata Only | Reviewed | Not Tested | |
| LargeButtons | PROPERTYGET/PROPERTYPUT | bool | declared | 364 | Metadata Only | Reviewed | Not Tested | |
| ColorButtons | PROPERTYGET/PROPERTYPUT | bool | declared | 365 | Metadata Only | Reviewed | Not Tested | |
| LibraryPath | PROPERTYGET | String | declared | 366 | Metadata Only | Reviewed | Not Tested | |
| MathCoprocessorAvailable | PROPERTYGET | bool | declared | 367 | Metadata Only | Reviewed | Not Tested | |
| MaxChange | PROPERTYGET/PROPERTYPUT | f64 | declared | 368 | Metadata Only | Reviewed | Not Tested | |
| MaxIterations | PROPERTYGET/PROPERTYPUT | i32 | declared | 369 | Metadata Only | Reviewed | Not Tested | |
| MemoryFree | PROPERTYGET | i32 | declared | 370 | Metadata Only | Reviewed | Not Tested | |
| MemoryTotal | PROPERTYGET | i32 | declared | 371 | Metadata Only | Reviewed | Not Tested | |
| MemoryUsed | PROPERTYGET | i32 | declared | 372 | Metadata Only | Reviewed | Not Tested | |
| MouseAvailable | PROPERTYGET | bool | declared | 373 | Metadata Only | Reviewed | Not Tested | |
| MoveAfterReturn | PROPERTYGET/PROPERTYPUT | bool | declared | 374 | Metadata Only | Reviewed | Not Tested | |
| OperatingSystem | PROPERTYGET | String | declared | 375 | Metadata Only | Reviewed | Not Tested | |
| OrganizationName | PROPERTYGET | String | declared | 376 | Metadata Only | Reviewed | Not Tested | |
| PathSeparator | PROPERTYGET | String | declared | 377 | Metadata Only | Reviewed | Not Tested | |
| PreviousSelections | PROPERTYGET | AutomationValue | declared | 378 | Metadata Only | Reviewed | Not Tested | |
| RecordRelative | PROPERTYGET | bool | declared | 379 | Metadata Only | Reviewed | Not Tested | |
| ReferenceStyle | PROPERTYGET/PROPERTYPUT | XlReferenceStyle | declared | 380 | Implemented | Reviewed | Live Tested | |
| TemplatesPath | PROPERTYGET | String | declared | 381 | Metadata Only | Reviewed | Not Tested | |
| ScreenUpdating | PROPERTYGET/PROPERTYPUT | bool | declared | 382 | Metadata Only | Reviewed | Not Tested | |
| StartupPath | PROPERTYGET | String | declared | 385 | Metadata Only | Reviewed | Not Tested | |
| StatusBar | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 386 | Metadata Only | Reviewed | Not Tested | |
| ShowToolTips | PROPERTYGET/PROPERTYPUT | bool | declared | 387 | Metadata Only | Reviewed | Not Tested | |
| NetworkTemplatesPath | PROPERTYGET | String | declared | 388 | Metadata Only | Reviewed | Not Tested | |
| UsableHeight | PROPERTYGET | f64 | declared | 389 | Metadata Only | Reviewed | Not Tested | |
| UsableWidth | PROPERTYGET | f64 | declared | 390 | Metadata Only | Reviewed | Not Tested | |
| UserName | PROPERTYGET/PROPERTYPUT | String | declared | 391 | Metadata Only | Reviewed | Not Tested | |
| Version | PROPERTYGET | String | declared | 392 | Implemented | Reviewed | Live Tested | |
| WindowsForPens | PROPERTYGET | bool | declared | 395 | Metadata Only | Reviewed | Not Tested | |
| WindowState | PROPERTYGET/PROPERTYPUT | XlWindowState | declared | 396 | Metadata Only | Reviewed | Not Tested | |
| Windows | PROPERTYGET | Windows | declared | 430 | Implemented | Reviewed | Blocked | |
| Names | PROPERTYGET | Names | declared | 442 | Metadata Only | Reviewed | Not Tested | |
| Sheets | PROPERTYGET | Sheets | declared | 485 | Implemented | Reviewed | Blocked | |
| Worksheets | PROPERTYGET | Sheets | declared | 494 | Implemented | Reviewed | Blocked | |
| AddIns | PROPERTYGET | AddIns | declared | 549 | Metadata Only | Reviewed | Not Tested | |
| Toolbars | PROPERTYGET | Toolbars | declared | 552 | Metadata Only | Reviewed | Not Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | bool | declared | 558 | Implemented | Reviewed | Live Tested | |
| Workbooks | PROPERTYGET | Workbooks | declared | 572 | Implemented | Reviewed | Live Tested | |
| Excel4MacroSheets | PROPERTYGET | Sheets | declared | 579 | Metadata Only | Reviewed | Not Tested | |
| Excel4IntlMacroSheets | PROPERTYGET | Sheets | declared | 581 | Metadata Only | Reviewed | Not Tested | |
| Modules | PROPERTYGET | Modules | declared | 582 | Metadata Only | Reviewed | Not Tested | |
| MenuBars | PROPERTYGET | MenuBars | declared | 589 | Metadata Only | Reviewed | Not Tested | |
| OnWindow | PROPERTYGET/PROPERTYPUT | String | declared | 623 | Metadata Only | Reviewed | Not Tested | |
| OnCalculate | PROPERTYGET/PROPERTYPUT | String | declared | 625 | Metadata Only | Reviewed | Not Tested | |
| OnEntry | PROPERTYGET/PROPERTYPUT | String | declared | 627 | Metadata Only | Reviewed | Not Tested | |
| OnDoubleClick | PROPERTYGET/PROPERTYPUT | String | declared | 628 | Metadata Only | Reviewed | Not Tested | |
| OnData | PROPERTYGET/PROPERTYPUT | String | declared | 629 | Metadata Only | Reviewed | Not Tested | |
| ActiveMenuBar | PROPERTYGET | MenuBar | declared | 758 | Metadata Only | Reviewed | Not Tested | |
| ActiveWindow | PROPERTYGET | Window | declared | 759 | Implemented | Reviewed | Blocked | |
| Dialogs | PROPERTYGET | Dialogs | declared | 761 | Metadata Only | Reviewed | Not Tested | |
| DialogSheets | PROPERTYGET | Sheets | declared | 764 | Metadata Only | Reviewed | Not Tested | |
| RegisteredFunctions | PROPERTYGET | AutomationValue | declared | 775 | Metadata Only | Reviewed | Not Tested | |
| ShortcutMenus | PROPERTYGET | Menu | declared | 776 | Metadata Only | Reviewed | Not Tested | |
| ThisWorkbook | PROPERTYGET | Workbook | declared | 778 | Metadata Only | Reviewed | Not Tested | |
| CustomListCount | PROPERTYGET | i32 | declared | 787 | Metadata Only | Reviewed | Not Tested | |
| ActiveDialog | PROPERTYGET | DialogSheet | declared | 815 | Metadata Only | Reviewed | Not Tested | |
| StandardFont | PROPERTYGET/PROPERTYPUT | String | declared | 924 | Metadata Only | Reviewed | Not Tested | |
| StandardFontSize | PROPERTYGET/PROPERTYPUT | f64 | declared | 925 | Metadata Only | Reviewed | Not Tested | |
| DisplayRecentFiles | PROPERTYGET/PROPERTYPUT | bool | declared | 926 | Metadata Only | Reviewed | Not Tested | |
| DisplayExcel4Menus | PROPERTYGET/PROPERTYPUT | bool | declared | 927 | Metadata Only | Reviewed | Not Tested | |
| EditDirectlyInCell | PROPERTYGET/PROPERTYPUT | bool | declared | 929 | Metadata Only | Reviewed | Not Tested | |
| AlertBeforeOverwriting | PROPERTYGET/PROPERTYPUT | bool | declared | 930 | Metadata Only | Reviewed | Not Tested | |
| FileConverters | PROPERTYGET | AutomationValue | declared | 931 | Metadata Only | Reviewed | Not Tested | |
| MailSession | PROPERTYGET | AutomationValue | declared | 942 | Metadata Only | Reviewed | Not Tested | |
| MailSystem | PROPERTYGET | XlMailSystem | declared | 971 | Metadata Only | Reviewed | Not Tested | |
| CopyObjectsWithCells | PROPERTYGET/PROPERTYPUT | bool | declared | 991 | Metadata Only | Reviewed | Not Tested | |
| AskToUpdateLinks | PROPERTYGET/PROPERTYPUT | bool | declared | 992 | Implemented | Reviewed | Blocked | |
| SheetsInNewWorkbook | PROPERTYGET/PROPERTYPUT | i32 | declared | 993 | Metadata Only | Reviewed | Not Tested | |
| OnSheetActivate | PROPERTYGET/PROPERTYPUT | String | declared | 1031 | Metadata Only | Reviewed | Not Tested | |
| DefaultFilePath | PROPERTYGET/PROPERTYPUT | String | declared | 1038 | Metadata Only | Reviewed | Not Tested | |
| DisplayFullScreen | PROPERTYGET/PROPERTYPUT | bool | declared | 1061 | Metadata Only | Reviewed | Not Tested | |
| PromptForSummaryInfo | PROPERTYGET/PROPERTYPUT | bool | declared | 1062 | Metadata Only | Reviewed | Not Tested | |
| EnableTipWizard | PROPERTYGET/PROPERTYPUT | bool | declared | 1064 | Metadata Only | Reviewed | Not Tested | |
| OnSheetDeactivate | PROPERTYGET/PROPERTYPUT | String | declared | 1081 | Metadata Only | Reviewed | Not Tested | |
| EnableCancelKey | PROPERTYGET/PROPERTYPUT | XlEnableCancelKey | declared | 1096 | Metadata Only | Reviewed | Not Tested | |
| MoveAfterReturnDirection | PROPERTYGET/PROPERTYPUT | XlDirection | declared | 1144 | Metadata Only | Reviewed | Not Tested | |
| AutoCorrect | PROPERTYGET | AutoCorrect | declared | 1145 | Metadata Only | Reviewed | Not Tested | |
| Cursor | PROPERTYGET/PROPERTYPUT | XlMousePointer | declared | 1161 | Metadata Only | Reviewed | Not Tested | |
| EnableAutoComplete | PROPERTYGET/PROPERTYPUT | bool | declared | 1179 | Metadata Only | Reviewed | Not Tested | |
| EnableAnimations | PROPERTYGET/PROPERTYPUT | bool | declared | 1180 | Metadata Only | Reviewed | Not Tested | |
| DisplayCommentIndicator | PROPERTYGET/PROPERTYPUT | XlCommentDisplayMode | declared | 1196 | Metadata Only | Reviewed | Not Tested | |
| EnableSound | PROPERTYGET/PROPERTYPUT | bool | declared | 1197 | Metadata Only | Reviewed | Not Tested | |
| FileSearch | PROPERTYGET | FileSearch | declared | 1200 | Metadata Only | Reviewed | Not Tested | |
| FileFind | PROPERTYGET | IFind | declared | 1201 | Metadata Only | Reviewed | Not Tested | |
| RecentFiles | PROPERTYGET | RecentFiles | declared | 1202 | Metadata Only | Reviewed | Not Tested | |
| ODBCErrors | PROPERTYGET | ODBCErrors | declared | 1203 | Metadata Only | Reviewed | Not Tested | |
| ODBCTimeout | PROPERTYGET/PROPERTYPUT | i32 | declared | 1204 | Metadata Only | Reviewed | Not Tested | |
| PivotTableSelection | PROPERTYGET/PROPERTYPUT | bool | declared | 1205 | Metadata Only | Reviewed | Not Tested | |
| RollZoom | PROPERTYGET/PROPERTYPUT | bool | declared | 1206 | Metadata Only | Reviewed | Not Tested | |
| ShowChartTipNames | PROPERTYGET/PROPERTYPUT | bool | declared | 1207 | Metadata Only | Reviewed | Not Tested | |
| ShowChartTipValues | PROPERTYGET/PROPERTYPUT | bool | declared | 1208 | Metadata Only | Reviewed | Not Tested | |
| DefaultSaveFormat | PROPERTYGET/PROPERTYPUT | XlFileFormat | declared | 1209 | Metadata Only | Reviewed | Not Tested | |
| UserControl | PROPERTYGET/PROPERTYPUT | bool | declared | 1210 | Metadata Only | Reviewed | Not Tested | |
| VBE | PROPERTYGET | VBE | declared | 1211 | Metadata Only | Reviewed | Not Tested | |
| EnableEvents | PROPERTYGET/PROPERTYPUT | bool | declared | 1212 | Metadata Only | Reviewed | Not Tested | |
| DisplayInfoWindow | PROPERTYGET/PROPERTYPUT | bool | declared | 1213 | Metadata Only | Reviewed | Not Tested | |
| Assistant | PROPERTYGET | Assistant | declared | 1438 | Metadata Only | Reviewed | Not Tested | |
| CommandBars | PROPERTYGET | CommandBars | declared | 1439 | Metadata Only | Reviewed | Not Tested | |
| WorksheetFunction | PROPERTYGET | WorksheetFunction | declared | 1440 | Metadata Only | Reviewed | Not Tested | |
| NewWorkbook | PROPERTYGET | NewFile | declared | 1565 | Metadata Only | Reviewed | Not Tested | |
| ExtendList | PROPERTYGET/PROPERTYPUT | bool | declared | 1793 | Metadata Only | Reviewed | Not Tested | |
| OLEDBErrors | PROPERTYGET | OLEDBErrors | declared | 1794 | Metadata Only | Reviewed | Not Tested | |
| COMAddIns | PROPERTYGET | COMAddIns | declared | 1796 | Metadata Only | Reviewed | Not Tested | |
| DefaultWebOptions | PROPERTYGET | DefaultWebOptions | declared | 1797 | Metadata Only | Reviewed | Not Tested | |
| ProductCode | PROPERTYGET | String | declared | 1798 | Metadata Only | Reviewed | Not Tested | |
| UserLibraryPath | PROPERTYGET | String | declared | 1799 | Metadata Only | Reviewed | Not Tested | |
| AutoPercentEntry | PROPERTYGET/PROPERTYPUT | bool | declared | 1800 | Metadata Only | Reviewed | Not Tested | |
| LanguageSettings | PROPERTYGET | LanguageSettings | declared | 1801 | Metadata Only | Reviewed | Not Tested | |
| Dummy101 | PROPERTYGET | Object | declared | 1802 | Metadata Only | Reviewed | Not Tested | |
| AnswerWizard | PROPERTYGET | AnswerWizard | declared | 1804 | Metadata Only | Reviewed | Not Tested | |
| CalculationVersion | PROPERTYGET | i32 | declared | 1806 | Implemented | Reviewed | Live Tested | |
| ShowWindowsInTaskbar | PROPERTYGET/PROPERTYPUT | bool | declared | 1807 | Metadata Only | Reviewed | Not Tested | |
| FeatureInstall | PROPERTYGET/PROPERTYPUT | MsoFeatureInstall | declared | 1808 | Metadata Only | Reviewed | Not Tested | |
| DecimalSeparator | PROPERTYGET/PROPERTYPUT | String | declared | 1809 | Metadata Only | Reviewed | Not Tested | |
| ThousandsSeparator | PROPERTYGET/PROPERTYPUT | String | declared | 1810 | Metadata Only | Reviewed | Not Tested | |
| Ready | PROPERTYGET | bool | declared | 1932 | Implemented | Reviewed | Live Tested | |
| FindFormat | PROPERTYGET/PROPERTYPUTREF | CellFormat | declared | 1934 | Metadata Only | Reviewed | Not Tested | |
| ReplaceFormat | PROPERTYGET/PROPERTYPUTREF | CellFormat | declared | 1935 | Metadata Only | Reviewed | Not Tested | |
| UsedObjects | PROPERTYGET | UsedObjects | declared | 1936 | Metadata Only | Reviewed | Not Tested | |
| CalculationState | PROPERTYGET | XlCalculationState | declared | 1937 | Implemented | Reviewed | Live Tested | |
| CalculationInterruptKey | PROPERTYGET/PROPERTYPUT | XlCalculationInterruptKey | declared | 1938 | Metadata Only | Reviewed | Not Tested | |
| Watches | PROPERTYGET | Watches | declared | 1939 | Metadata Only | Reviewed | Not Tested | |
| DisplayFunctionToolTips | PROPERTYGET/PROPERTYPUT | bool | declared | 1940 | Metadata Only | Reviewed | Not Tested | |
| AutomationSecurity | PROPERTYGET/PROPERTYPUT | MsoAutomationSecurity | declared | 1941 | Implemented | Reviewed | Blocked | |
| FileDialog | PROPERTYGET | FileDialog | declared | 1942 | Metadata Only | Reviewed | Not Tested | |
| DisplayPasteOptions | PROPERTYGET/PROPERTYPUT | bool | declared | 1946 | Metadata Only | Reviewed | Not Tested | |
| DisplayInsertOptions | PROPERTYGET/PROPERTYPUT | bool | declared | 1947 | Metadata Only | Reviewed | Not Tested | |
| GenerateGetPivotData | PROPERTYGET/PROPERTYPUT | bool | declared | 1948 | Metadata Only | Reviewed | Not Tested | |
| AutoRecover | PROPERTYGET | AutoRecover | declared | 1949 | Metadata Only | Reviewed | Not Tested | |
| Hwnd | PROPERTYGET | i32 | declared | 1950 | Implemented | Reviewed | Live Tested | |
| Hinstance | PROPERTYGET | i32 | declared | 1951 | Metadata Only | Reviewed | Not Tested | |
| ErrorCheckingOptions | PROPERTYGET | ErrorCheckingOptions | declared | 1954 | Metadata Only | Reviewed | Not Tested | |
| AutoFormatAsYouTypeReplaceHyperlinks | PROPERTYGET/PROPERTYPUT | bool | declared | 1955 | Metadata Only | Reviewed | Not Tested | |
| SmartTagRecognizers | PROPERTYGET | SmartTagRecognizers | declared | 1956 | Metadata Only | Reviewed | Not Tested | |
| SpellingOptions | PROPERTYGET | SpellingOptions | declared | 1957 | Metadata Only | Reviewed | Not Tested | |
| Speech | PROPERTYGET | Speech | declared | 1958 | Metadata Only | Reviewed | Not Tested | |
| MapPaperSize | PROPERTYGET/PROPERTYPUT | bool | declared | 1959 | Metadata Only | Reviewed | Not Tested | |
| ShowStartupDialog | PROPERTYGET/PROPERTYPUT | bool | declared | 1960 | Metadata Only | Reviewed | Not Tested | |
| UseSystemSeparators | PROPERTYGET/PROPERTYPUT | bool | declared | 1961 | Metadata Only | Reviewed | Not Tested | |
| ThisCell | PROPERTYGET | Range | declared | 1962 | Metadata Only | Reviewed | Not Tested | |
| RTD | PROPERTYGET | RTD | declared | 1963 | Metadata Only | Reviewed | Not Tested | |
| DisplayDocumentActionTaskPane | PROPERTYGET/PROPERTYPUT | bool | declared | 2251 | Metadata Only | Reviewed | Not Tested | |
| ArbitraryXMLSupportAvailable | PROPERTYGET | bool | declared | 2254 | Metadata Only | Reviewed | Not Tested | |
| MeasurementUnit | PROPERTYGET/PROPERTYPUT | i32 | declared | 2375 | Metadata Only | Reviewed | Not Tested | |
| ShowSelectionFloaties | PROPERTYGET/PROPERTYPUT | bool | declared | 2376 | Metadata Only | Reviewed | Not Tested | |
| ShowMenuFloaties | PROPERTYGET/PROPERTYPUT | bool | declared | 2377 | Metadata Only | Reviewed | Not Tested | |
| ShowDevTools | PROPERTYGET/PROPERTYPUT | bool | declared | 2378 | Metadata Only | Reviewed | Not Tested | |
| EnableLivePreview | PROPERTYGET/PROPERTYPUT | bool | declared | 2379 | Metadata Only | Reviewed | Not Tested | |
| DisplayDocumentInformationPanel | PROPERTYGET/PROPERTYPUT | bool | declared | 2380 | Metadata Only | Reviewed | Not Tested | |
| AlwaysUseClearType | PROPERTYGET/PROPERTYPUT | bool | declared | 2381 | Metadata Only | Reviewed | Not Tested | |
| WarnOnFunctionNameConflict | PROPERTYGET/PROPERTYPUT | bool | declared | 2382 | Metadata Only | Reviewed | Not Tested | |
| FormulaBarHeight | PROPERTYGET/PROPERTYPUT | i32 | declared | 2383 | Metadata Only | Reviewed | Not Tested | |
| DisplayFormulaAutoComplete | PROPERTYGET/PROPERTYPUT | bool | declared | 2384 | Metadata Only | Reviewed | Not Tested | |
| GenerateTableRefs | PROPERTYGET/PROPERTYPUT | XlGenerateTableRefs | declared | 2385 | Metadata Only | Reviewed | Not Tested | |
| Assistance | PROPERTYGET | IAssistance | declared | 2386 | Metadata Only | Reviewed | Not Tested | |
| EnableLargeOperationAlert | PROPERTYGET/PROPERTYPUT | bool | declared | 2388 | Metadata Only | Reviewed | Not Tested | |
| LargeOperationCellThousandCount | PROPERTYGET/PROPERTYPUT | i32 | declared | 2389 | Metadata Only | Reviewed | Not Tested | |
| DeferAsyncQueries | PROPERTYGET/PROPERTYPUT | bool | declared | 2390 | Metadata Only | Reviewed | Not Tested | |
| MultiThreadedCalculation | PROPERTYGET | MultiThreadedCalculation | declared | 2391 | Metadata Only | Reviewed | Not Tested | |
| ActiveEncryptionSession | PROPERTYGET | i32 | declared | 2394 | Metadata Only | Reviewed | Not Tested | |
| HighQualityModeForGraphics | PROPERTYGET/PROPERTYPUT | bool | declared | 2395 | Metadata Only | Reviewed | Not Tested | |
| FileExportConverters | PROPERTYGET | FileExportConverters | declared | 2768 | Metadata Only | Reviewed | Not Tested | |
| SmartArtLayouts | PROPERTYGET | SmartArtLayouts | declared | 2772 | Metadata Only | Reviewed | Not Tested | |
| SmartArtQuickStyles | PROPERTYGET | SmartArtQuickStyles | declared | 2773 | Metadata Only | Reviewed | Not Tested | |
| SmartArtColors | PROPERTYGET | SmartArtColors | declared | 2774 | Metadata Only | Reviewed | Not Tested | |
| AddIns2 | PROPERTYGET | AddIns2 | declared | 2775 | Metadata Only | Reviewed | Not Tested | |
| PrintCommunication | PROPERTYGET/PROPERTYPUT | bool | declared | 2776 | Metadata Only | Reviewed | Not Tested | |
| UseClusterConnector | PROPERTYGET/PROPERTYPUT | bool | declared | 2778 | Metadata Only | Reviewed | Not Tested | |
| ClusterConnector | PROPERTYGET/PROPERTYPUT | String | declared | 2779 | Metadata Only | Reviewed | Not Tested | |
| Quitting | PROPERTYGET | bool | declared | 2780 | Metadata Only | Reviewed | Not Tested | |
| Dummy22 | PROPERTYGET/PROPERTYPUT | bool | declared | 2781 | Metadata Only | Reviewed | Not Tested | |
| Dummy23 | PROPERTYGET/PROPERTYPUT | bool | declared | 2782 | Metadata Only | Reviewed | Not Tested | |
| ProtectedViewWindows | PROPERTYGET | ProtectedViewWindows | declared | 2783 | Metadata Only | Reviewed | Not Tested | |
| ActiveProtectedViewWindow | PROPERTYGET | ProtectedViewWindow | declared | 2784 | Metadata Only | Reviewed | Not Tested | |
| IsSandboxed | PROPERTYGET | bool | declared | 2785 | Metadata Only | Reviewed | Not Tested | |
| SaveISO8601Dates | PROPERTYGET/PROPERTYPUT | bool | declared | 2786 | Metadata Only | Reviewed | Not Tested | |
| HinstancePtr | PROPERTYGET | AutomationValue | declared | 2787 | Metadata Only | Reviewed | Not Tested | |
| FileValidation | PROPERTYGET/PROPERTYPUT | MsoFileValidationMode | declared | 2788 | Metadata Only | Reviewed | Not Tested | |
| FileValidationPivot | PROPERTYGET/PROPERTYPUT | XlFileValidationPivotMode | declared | 2789 | Metadata Only | Reviewed | Not Tested | |
| ShowQuickAnalysis | PROPERTYGET/PROPERTYPUT | bool | declared | 2994 | Metadata Only | Reviewed | Not Tested | |
| QuickAnalysis | PROPERTYGET | QuickAnalysis | declared | 2995 | Metadata Only | Reviewed | Not Tested | |
| FlashFill | PROPERTYGET/PROPERTYPUT | bool | declared | 2996 | Metadata Only | Reviewed | Not Tested | |
| EnableMacroAnimations | PROPERTYGET/PROPERTYPUT | bool | declared | 2997 | Metadata Only | Reviewed | Not Tested | |
| ChartDataPointTrack | PROPERTYGET/PROPERTYPUT | bool | declared | 2998 | Metadata Only | Reviewed | Not Tested | |
| FlashFillMode | PROPERTYGET/PROPERTYPUT | bool | declared | 2999 | Metadata Only | Reviewed | Not Tested | |
| MergeInstances | PROPERTYGET/PROPERTYPUT | bool | declared | 3000 | Metadata Only | Reviewed | Not Tested | |
| EnableCheckFileExtensions | PROPERTYGET/PROPERTYPUT | bool | declared | 3158 | Metadata Only | Reviewed | Not Tested | |
| DefaultPivotTableLayoutOptions | PROPERTYGET | DefaultPivotTableLayoutOptions | declared | 3271 | Metadata Only | Reviewed | Not Tested | |
| ShowConvertToDataType | PROPERTYGET/PROPERTYPUT | bool | declared | 3311 | Metadata Only | Reviewed | Not Tested | |
| TruncateLeadingZeros | PROPERTYGET/PROPERTYPUT | bool | declared | 3312 | Metadata Only | Reviewed | Not Tested | |
| TruncateLargeNumbers | PROPERTYGET/PROPERTYPUT | bool | declared | 3313 | Metadata Only | Reviewed | Not Tested | |
| ConvertNumbersWithECharacter | PROPERTYGET/PROPERTYPUT | bool | declared | 3314 | Metadata Only | Reviewed | Not Tested | |
| CSVDisplayNumberConversionWarning | PROPERTYGET/PROPERTYPUT | bool | declared | 3315 | Metadata Only | Reviewed | Not Tested | |
| CSVKeepColumnAsTextIfMultipleEntriesAreText | PROPERTYGET/PROPERTYPUT | bool | declared | 3316 | Metadata Only | Reviewed | Not Tested | |
| DataPrivacyOptions | PROPERTYGET | DataPrivacyOptions | declared | 3317 | Metadata Only | Reviewed | Not Tested | |
| SensitivityLabelPolicy | PROPERTYGET | SensitivityLabelPolicy | declared | 3365 | Metadata Only | Reviewed | Not Tested | |
| FormatStaleValues | PROPERTYGET/PROPERTYPUT | bool | declared | 3401 | Metadata Only | Reviewed | Not Tested | |
| MaxSupportedCompatibilityVersion | PROPERTYGET | i32 | declared | 3417 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| _Evaluate | AutomationValue | 1 | declared | -5 | Metadata Only | Reviewed | Not Tested | |
| Evaluate | AutomationValue | 1 | declared | 1 | Implemented | Reviewed | Live Tested | |
| RegisterXLL | bool | 1 | declared | 30 | Metadata Only | Reviewed | Not Tested | |
| _WSFunction | AutomationValue | 30 | declared | 169 | Metadata Only | Reviewed | Not Tested | |
| SaveWorkspace | Unknown | 1 | declared | 212 | Metadata Only | Reviewed | Not Tested | |
| AddChartAutoFormat | Unknown | 3 | declared | 216 | Metadata Only | Reviewed | Not Tested | |
| DeleteChartAutoFormat | Unknown | 1 | declared | 217 | Metadata Only | Reviewed | Not Tested | |
| SetDefaultChart | Unknown | 2 | declared | 219 | Metadata Only | Reviewed | Not Tested | |
| Run | AutomationValue | 31 | declared | 259 | Implemented | Reviewed | Blocked | |
| Calculate | Unknown | 0 | declared | 279 | Implemented | Reviewed | Live Tested | |
| Save | Unknown | 1 | declared | 283 | Metadata Only | Reviewed | Not Tested | |
| Repeat | Unknown | 0 | declared | 301 | Metadata Only | Reviewed | Not Tested | |
| Quit | Unknown | 0 | declared | 302 | Implemented | Reviewed | Live Tested | |
| Undo | Unknown | 0 | declared | 303 | Metadata Only | Reviewed | Not Tested | |
| ConvertFormula | AutomationValue | 5 | declared | 325 | Implemented | Reviewed | Live Tested | |
| DDEExecute | Unknown | 2 | declared | 333 | Metadata Only | Reviewed | Not Tested | |
| DDEInitiate | i32 | 2 | declared | 334 | Metadata Only | Reviewed | Not Tested | |
| DDEPoke | Unknown | 3 | declared | 335 | Metadata Only | Reviewed | Not Tested | |
| DDERequest | AutomationValue | 2 | declared | 336 | Metadata Only | Reviewed | Not Tested | |
| DDETerminate | Unknown | 1 | declared | 337 | Metadata Only | Reviewed | Not Tested | |
| DoubleClick | Unknown | 0 | declared | 349 | Metadata Only | Reviewed | Not Tested | |
| ExecuteExcel4Macro | AutomationValue | 1 | declared | 350 | Metadata Only | Reviewed | Not Tested | |
| Help | Unknown | 2 | declared | 354 | Metadata Only | Reviewed | Not Tested | |
| InputBox | AutomationValue | 8 | declared | 357 | Metadata Only | Reviewed | Not Tested | |
| SendKeys | Unknown | 2 | declared | 383 | Metadata Only | Reviewed | Not Tested | |
| _Wait | Unknown | 1 | declared | 393 | Metadata Only | Reviewed | Not Tested | |
| Goto | Unknown | 2 | declared | 475 | Implemented | Reviewed | Blocked | |
| CheckSpelling | bool | 3 | declared | 505 | Metadata Only | Reviewed | Not Tested | |
| OnTime | Unknown | 4 | declared | 624 | Metadata Only | Reviewed | Not Tested | |
| OnKey | Unknown | 2 | declared | 626 | Metadata Only | Reviewed | Not Tested | |
| Intersect | Range | 30 | declared | 766 | Metadata Only | Reviewed | Not Tested | |
| OnRepeat | Unknown | 2 | declared | 769 | Metadata Only | Reviewed | Not Tested | |
| OnUndo | Unknown | 2 | declared | 770 | Metadata Only | Reviewed | Not Tested | |
| RecordMacro | Unknown | 2 | declared | 773 | Metadata Only | Reviewed | Not Tested | |
| Union | Range | 30 | declared | 779 | Implemented | Reviewed | Live Tested | |
| AddCustomList | Unknown | 2 | declared | 780 | Metadata Only | Reviewed | Not Tested | |
| DeleteCustomList | Unknown | 1 | declared | 783 | Metadata Only | Reviewed | Not Tested | |
| GetCustomListNum | i32 | 1 | declared | 785 | Metadata Only | Reviewed | Not Tested | |
| GetCustomListContents | AutomationValue | 1 | declared | 786 | Metadata Only | Reviewed | Not Tested | |
| Volatile | Unknown | 1 | declared | 788 | Metadata Only | Reviewed | Not Tested | |
| _Run2 | AutomationValue | 31 | declared | 806 | Metadata Only | Reviewed | Not Tested | |
| ResetTipWizard | Unknown | 0 | declared | 928 | Metadata Only | Reviewed | Not Tested | |
| MailLogon | Unknown | 3 | declared | 943 | Metadata Only | Reviewed | Not Tested | |
| MailLogoff | Unknown | 0 | declared | 945 | Metadata Only | Reviewed | Not Tested | |
| NextLetter | Workbook | 0 | declared | 972 | Metadata Only | Reviewed | Not Tested | |
| _FindFile | Unknown | 0 | declared | 1068 | Metadata Only | Reviewed | Not Tested | |
| GetOpenFilename | AutomationValue | 5 | declared | 1075 | Metadata Only | Reviewed | Not Tested | |
| GetSaveAsFilename | AutomationValue | 5 | declared | 1076 | Metadata Only | Reviewed | Not Tested | |
| CentimetersToPoints | f64 | 1 | declared | 1086 | Metadata Only | Reviewed | Not Tested | |
| InchesToPoints | f64 | 1 | declared | 1087 | Metadata Only | Reviewed | Not Tested | |
| ActivateMicrosoftApp | Unknown | 1 | declared | 1095 | Metadata Only | Reviewed | Not Tested | |
| _MacroOptions | Unknown | 10 | declared | 1135 | Metadata Only | Reviewed | Not Tested | |
| Wait | bool | 1 | declared | 1770 | Metadata Only | Reviewed | Not Tested | |
| FindFile | bool | 0 | declared | 1771 | Metadata Only | Reviewed | Not Tested | |
| Dummy1 | AutomationValue | 4 | declared | 1782 | Metadata Only | Reviewed | Not Tested | |
| Dummy2 | AutomationValue | 8 | declared | 1783 | Metadata Only | Reviewed | Not Tested | |
| Dummy3 | AutomationValue | 0 | declared | 1784 | Metadata Only | Reviewed | Not Tested | |
| Dummy4 | AutomationValue | 15 | declared | 1785 | Metadata Only | Reviewed | Not Tested | |
| Dummy5 | AutomationValue | 13 | declared | 1786 | Metadata Only | Reviewed | Not Tested | |
| Dummy6 | AutomationValue | 0 | declared | 1787 | Metadata Only | Reviewed | Not Tested | |
| Dummy7 | AutomationValue | 0 | declared | 1788 | Metadata Only | Reviewed | Not Tested | |
| Dummy8 | AutomationValue | 1 | declared | 1789 | Metadata Only | Reviewed | Not Tested | |
| Dummy9 | AutomationValue | 0 | declared | 1790 | Metadata Only | Reviewed | Not Tested | |
| Dummy10 | bool | 1 | declared | 1791 | Metadata Only | Reviewed | Not Tested | |
| Dummy11 | Unknown | 0 | declared | 1792 | Metadata Only | Reviewed | Not Tested | |
| GetPhonetic | String | 1 | declared | 1795 | Metadata Only | Reviewed | Not Tested | |
| Dummy12 | Unknown | 2 | declared | 1803 | Metadata Only | Reviewed | Not Tested | |
| CalculateFull | Unknown | 0 | declared | 1805 | Implemented | Reviewed | Live Tested | |
| Dummy13 | AutomationValue | 30 | declared | 1933 | Metadata Only | Reviewed | Not Tested | |
| Dummy14 | Unknown | 0 | declared | 1944 | Metadata Only | Reviewed | Not Tested | |
| CalculateFullRebuild | Unknown | 0 | declared | 1945 | Implemented | Reviewed | Live Tested | |
| CheckAbort | Unknown | 1 | declared | 1952 | Metadata Only | Reviewed | Not Tested | |
| DisplayXMLSourcePane | Unknown | 1 | declared | 2252 | Metadata Only | Reviewed | Not Tested | |
| Support | AutomationValue | 3 | declared | 2255 | Metadata Only | Reviewed | Not Tested | |
| Dummy20 | AutomationValue | 1 | declared | 2373 | Metadata Only | Reviewed | Not Tested | |
| CalculateUntilAsyncQueriesDone | Unknown | 0 | declared | 2387 | Implemented | Reviewed | Live Tested | |
| SharePointVersion | i32 | 1 | declared | 2392 | Metadata Only | Reviewed | Not Tested | |
| MacroOptions | Unknown | 11 | declared | 2770 | Metadata Only | Reviewed | Not Tested | |
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
