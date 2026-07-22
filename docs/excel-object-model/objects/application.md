# Application

## Summary

The root Automation object for a locally activated Excel instance. The initial crate exposes only a deliberately small lifecycle and workbook-navigation slice.

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
| Surface class | implemented-wrapper |
| Crate type | `excel_com::Application` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Relationships

| Relationship | Target | Status |
|---|---|---|
| `ActiveCell` | `excel.range` | Metadata Only |
| `ActiveWorkbook` | `excel.workbook` | Metadata Only |
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
| `Union` | `excel.range` | Metadata Only |
| `Workbooks` | `excel.workbooks` | Implemented |

## Properties

| Property | Access | Type | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---:|---|---|---|---|
| _Default | PROPERTYGET | String | 0 | Metadata Only | Reviewed | Not Tested | |
| UILanguage | PROPERTYGET/PROPERTYPUT | i32 | 2 | Metadata Only | Reviewed | Not Tested | |
| Value | PROPERTYGET | String | 6 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET | String | 110 | Metadata Only | Reviewed | Not Tested | |
| Charts | PROPERTYGET | Sheets | 121 | Metadata Only | Reviewed | Not Tested | |
| Width | PROPERTYGET/PROPERTYPUT | f64 | 122 | Metadata Only | Reviewed | Not Tested | |
| Height | PROPERTYGET/PROPERTYPUT | f64 | 123 | Metadata Only | Reviewed | Not Tested | |
| Top | PROPERTYGET/PROPERTYPUT | f64 | 126 | Metadata Only | Reviewed | Not Tested | |
| Left | PROPERTYGET/PROPERTYPUT | f64 | 127 | Metadata Only | Reviewed | Not Tested | |
| Caption | PROPERTYGET/PROPERTYPUT | String | 139 | Metadata Only | Reviewed | Not Tested | |
| Selection | PROPERTYGET | Object | 147 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Application | 150 | Metadata Only | Reviewed | Not Tested | |
| ActiveChart | PROPERTYGET | Chart | 183 | Metadata Only | Reviewed | Not Tested | |
| Range | PROPERTYGET | Range | 197 | Metadata Only | Reviewed | Not Tested | |
| DefaultSheetDirection | PROPERTYGET/PROPERTYPUT | i32 | 229 | Metadata Only | Reviewed | Not Tested | |
| CursorMovement | PROPERTYGET/PROPERTYPUT | i32 | 232 | Metadata Only | Reviewed | Not Tested | |
| ControlCharacters | PROPERTYGET/PROPERTYPUT | bool | 233 | Metadata Only | Reviewed | Not Tested | |
| Cells | PROPERTYGET | Range | 238 | Metadata Only | Reviewed | Not Tested | |
| Columns | PROPERTYGET | Range | 241 | Metadata Only | Reviewed | Not Tested | |
| Rows | PROPERTYGET | Range | 258 | Metadata Only | Reviewed | Not Tested | |
| Path | PROPERTYGET | String | 291 | Metadata Only | Reviewed | Not Tested | |
| ActiveCell | PROPERTYGET | Range | 305 | Metadata Only | Reviewed | Not Tested | |
| ActivePrinter | PROPERTYGET/PROPERTYPUT | String | 306 | Metadata Only | Reviewed | Not Tested | |
| ActiveSheet | PROPERTYGET | Object | 307 | Metadata Only | Reviewed | Not Tested | |
| ActiveWorkbook | PROPERTYGET | Workbook | 308 | Metadata Only | Reviewed | Not Tested | |
| TransitionMenuKey | PROPERTYGET/PROPERTYPUT | String | 310 | Metadata Only | Reviewed | Not Tested | |
| TransitionMenuKeyAction | PROPERTYGET/PROPERTYPUT | i32 | 311 | Metadata Only | Reviewed | Not Tested | |
| TransitionNavigKeys | PROPERTYGET/PROPERTYPUT | bool | 312 | Metadata Only | Reviewed | Not Tested | |
| AltStartupPath | PROPERTYGET/PROPERTYPUT | String | 313 | Metadata Only | Reviewed | Not Tested | |
| Build | PROPERTYGET | i32 | 314 | Metadata Only | Reviewed | Not Tested | |
| CalculateBeforeSave | PROPERTYGET/PROPERTYPUT | bool | 315 | Metadata Only | Reviewed | Not Tested | |
| Calculation | PROPERTYGET/PROPERTYPUT | XlCalculation | 316 | Metadata Only | Reviewed | Not Tested | |
| Caller | PROPERTYGET | AutomationValue | 317 | Metadata Only | Reviewed | Not Tested | |
| CanPlaySounds | PROPERTYGET | bool | 318 | Metadata Only | Reviewed | Not Tested | |
| CanRecordSounds | PROPERTYGET | bool | 319 | Metadata Only | Reviewed | Not Tested | |
| CellDragAndDrop | PROPERTYGET/PROPERTYPUT | bool | 320 | Metadata Only | Reviewed | Not Tested | |
| ClipboardFormats | PROPERTYGET | AutomationValue | 321 | Metadata Only | Reviewed | Not Tested | |
| DisplayClipboardWindow | PROPERTYGET/PROPERTYPUT | bool | 322 | Metadata Only | Reviewed | Not Tested | |
| CommandUnderlines | PROPERTYGET/PROPERTYPUT | XlCommandUnderlines | 323 | Metadata Only | Reviewed | Not Tested | |
| ConstrainNumeric | PROPERTYGET/PROPERTYPUT | bool | 324 | Metadata Only | Reviewed | Not Tested | |
| CutCopyMode | PROPERTYGET/PROPERTYPUT | XlCutCopyMode | 330 | Metadata Only | Reviewed | Not Tested | |
| DataEntryMode | PROPERTYGET/PROPERTYPUT | i32 | 331 | Metadata Only | Reviewed | Not Tested | |
| DDEAppReturnCode | PROPERTYGET | i32 | 332 | Metadata Only | Reviewed | Not Tested | |
| DisplayAlerts | PROPERTYGET/PROPERTYPUT | bool | 343 | Metadata Only | Reviewed | Not Tested | |
| DisplayFormulaBar | PROPERTYGET/PROPERTYPUT | bool | 344 | Metadata Only | Reviewed | Not Tested | |
| DisplayNoteIndicator | PROPERTYGET/PROPERTYPUT | bool | 345 | Metadata Only | Reviewed | Not Tested | |
| DisplayScrollBars | PROPERTYGET/PROPERTYPUT | bool | 346 | Metadata Only | Reviewed | Not Tested | |
| DisplayStatusBar | PROPERTYGET/PROPERTYPUT | bool | 347 | Metadata Only | Reviewed | Not Tested | |
| FixedDecimal | PROPERTYGET/PROPERTYPUT | bool | 351 | Metadata Only | Reviewed | Not Tested | |
| FixedDecimalPlaces | PROPERTYGET/PROPERTYPUT | i32 | 352 | Metadata Only | Reviewed | Not Tested | |
| IgnoreRemoteRequests | PROPERTYGET/PROPERTYPUT | bool | 356 | Metadata Only | Reviewed | Not Tested | |
| Interactive | PROPERTYGET/PROPERTYPUT | bool | 361 | Metadata Only | Reviewed | Not Tested | |
| International | PROPERTYGET | AutomationValue | 362 | Metadata Only | Reviewed | Not Tested | |
| Iteration | PROPERTYGET/PROPERTYPUT | bool | 363 | Metadata Only | Reviewed | Not Tested | |
| LargeButtons | PROPERTYGET/PROPERTYPUT | bool | 364 | Metadata Only | Reviewed | Not Tested | |
| ColorButtons | PROPERTYGET/PROPERTYPUT | bool | 365 | Metadata Only | Reviewed | Not Tested | |
| LibraryPath | PROPERTYGET | String | 366 | Metadata Only | Reviewed | Not Tested | |
| MathCoprocessorAvailable | PROPERTYGET | bool | 367 | Metadata Only | Reviewed | Not Tested | |
| MaxChange | PROPERTYGET/PROPERTYPUT | f64 | 368 | Metadata Only | Reviewed | Not Tested | |
| MaxIterations | PROPERTYGET/PROPERTYPUT | i32 | 369 | Metadata Only | Reviewed | Not Tested | |
| MemoryFree | PROPERTYGET | i32 | 370 | Metadata Only | Reviewed | Not Tested | |
| MemoryTotal | PROPERTYGET | i32 | 371 | Metadata Only | Reviewed | Not Tested | |
| MemoryUsed | PROPERTYGET | i32 | 372 | Metadata Only | Reviewed | Not Tested | |
| MouseAvailable | PROPERTYGET | bool | 373 | Metadata Only | Reviewed | Not Tested | |
| MoveAfterReturn | PROPERTYGET/PROPERTYPUT | bool | 374 | Metadata Only | Reviewed | Not Tested | |
| OperatingSystem | PROPERTYGET | String | 375 | Metadata Only | Reviewed | Not Tested | |
| OrganizationName | PROPERTYGET | String | 376 | Metadata Only | Reviewed | Not Tested | |
| PathSeparator | PROPERTYGET | String | 377 | Metadata Only | Reviewed | Not Tested | |
| PreviousSelections | PROPERTYGET | AutomationValue | 378 | Metadata Only | Reviewed | Not Tested | |
| RecordRelative | PROPERTYGET | bool | 379 | Metadata Only | Reviewed | Not Tested | |
| ReferenceStyle | PROPERTYGET/PROPERTYPUT | XlReferenceStyle | 380 | Metadata Only | Reviewed | Not Tested | |
| TemplatesPath | PROPERTYGET | String | 381 | Metadata Only | Reviewed | Not Tested | |
| ScreenUpdating | PROPERTYGET/PROPERTYPUT | bool | 382 | Metadata Only | Reviewed | Not Tested | |
| StartupPath | PROPERTYGET | String | 385 | Metadata Only | Reviewed | Not Tested | |
| StatusBar | PROPERTYGET/PROPERTYPUT | AutomationValue | 386 | Metadata Only | Reviewed | Not Tested | |
| ShowToolTips | PROPERTYGET/PROPERTYPUT | bool | 387 | Metadata Only | Reviewed | Not Tested | |
| NetworkTemplatesPath | PROPERTYGET | String | 388 | Metadata Only | Reviewed | Not Tested | |
| UsableHeight | PROPERTYGET | f64 | 389 | Metadata Only | Reviewed | Not Tested | |
| UsableWidth | PROPERTYGET | f64 | 390 | Metadata Only | Reviewed | Not Tested | |
| UserName | PROPERTYGET/PROPERTYPUT | String | 391 | Metadata Only | Reviewed | Not Tested | |
| Version | PROPERTYGET | String | 392 | Implemented | Reviewed | Live Tested | |
| WindowsForPens | PROPERTYGET | bool | 395 | Metadata Only | Reviewed | Not Tested | |
| WindowState | PROPERTYGET/PROPERTYPUT | XlWindowState | 396 | Metadata Only | Reviewed | Not Tested | |
| Windows | PROPERTYGET | Windows | 430 | Metadata Only | Reviewed | Not Tested | |
| Names | PROPERTYGET | Names | 442 | Metadata Only | Reviewed | Not Tested | |
| Sheets | PROPERTYGET | Sheets | 485 | Metadata Only | Reviewed | Not Tested | |
| Worksheets | PROPERTYGET | Sheets | 494 | Metadata Only | Reviewed | Not Tested | |
| AddIns | PROPERTYGET | AddIns | 549 | Metadata Only | Reviewed | Not Tested | |
| Toolbars | PROPERTYGET | Toolbars | 552 | Metadata Only | Reviewed | Not Tested | |
| Visible | PROPERTYGET/PROPERTYPUT | bool | 558 | Implemented | Reviewed | Live Tested | |
| Workbooks | PROPERTYGET | Workbooks | 572 | Implemented | Reviewed | Live Tested | |
| Excel4MacroSheets | PROPERTYGET | Sheets | 579 | Metadata Only | Reviewed | Not Tested | |
| Excel4IntlMacroSheets | PROPERTYGET | Sheets | 581 | Metadata Only | Reviewed | Not Tested | |
| Modules | PROPERTYGET | Modules | 582 | Metadata Only | Reviewed | Not Tested | |
| MenuBars | PROPERTYGET | MenuBars | 589 | Metadata Only | Reviewed | Not Tested | |
| OnWindow | PROPERTYGET/PROPERTYPUT | String | 623 | Metadata Only | Reviewed | Not Tested | |
| OnCalculate | PROPERTYGET/PROPERTYPUT | String | 625 | Metadata Only | Reviewed | Not Tested | |
| OnEntry | PROPERTYGET/PROPERTYPUT | String | 627 | Metadata Only | Reviewed | Not Tested | |
| OnDoubleClick | PROPERTYGET/PROPERTYPUT | String | 628 | Metadata Only | Reviewed | Not Tested | |
| OnData | PROPERTYGET/PROPERTYPUT | String | 629 | Metadata Only | Reviewed | Not Tested | |
| ActiveMenuBar | PROPERTYGET | MenuBar | 758 | Metadata Only | Reviewed | Not Tested | |
| ActiveWindow | PROPERTYGET | Window | 759 | Metadata Only | Reviewed | Not Tested | |
| Dialogs | PROPERTYGET | Dialogs | 761 | Metadata Only | Reviewed | Not Tested | |
| DialogSheets | PROPERTYGET | Sheets | 764 | Metadata Only | Reviewed | Not Tested | |
| RegisteredFunctions | PROPERTYGET | AutomationValue | 775 | Metadata Only | Reviewed | Not Tested | |
| ShortcutMenus | PROPERTYGET | Menu | 776 | Metadata Only | Reviewed | Not Tested | |
| ThisWorkbook | PROPERTYGET | Workbook | 778 | Metadata Only | Reviewed | Not Tested | |
| CustomListCount | PROPERTYGET | i32 | 787 | Metadata Only | Reviewed | Not Tested | |
| ActiveDialog | PROPERTYGET | DialogSheet | 815 | Metadata Only | Reviewed | Not Tested | |
| StandardFont | PROPERTYGET/PROPERTYPUT | String | 924 | Metadata Only | Reviewed | Not Tested | |
| StandardFontSize | PROPERTYGET/PROPERTYPUT | f64 | 925 | Metadata Only | Reviewed | Not Tested | |
| DisplayRecentFiles | PROPERTYGET/PROPERTYPUT | bool | 926 | Metadata Only | Reviewed | Not Tested | |
| DisplayExcel4Menus | PROPERTYGET/PROPERTYPUT | bool | 927 | Metadata Only | Reviewed | Not Tested | |
| EditDirectlyInCell | PROPERTYGET/PROPERTYPUT | bool | 929 | Metadata Only | Reviewed | Not Tested | |
| AlertBeforeOverwriting | PROPERTYGET/PROPERTYPUT | bool | 930 | Metadata Only | Reviewed | Not Tested | |
| FileConverters | PROPERTYGET | AutomationValue | 931 | Metadata Only | Reviewed | Not Tested | |
| MailSession | PROPERTYGET | AutomationValue | 942 | Metadata Only | Reviewed | Not Tested | |
| MailSystem | PROPERTYGET | XlMailSystem | 971 | Metadata Only | Reviewed | Not Tested | |
| CopyObjectsWithCells | PROPERTYGET/PROPERTYPUT | bool | 991 | Metadata Only | Reviewed | Not Tested | |
| AskToUpdateLinks | PROPERTYGET/PROPERTYPUT | bool | 992 | Metadata Only | Reviewed | Not Tested | |
| SheetsInNewWorkbook | PROPERTYGET/PROPERTYPUT | i32 | 993 | Metadata Only | Reviewed | Not Tested | |
| OnSheetActivate | PROPERTYGET/PROPERTYPUT | String | 1031 | Metadata Only | Reviewed | Not Tested | |
| DefaultFilePath | PROPERTYGET/PROPERTYPUT | String | 1038 | Metadata Only | Reviewed | Not Tested | |
| DisplayFullScreen | PROPERTYGET/PROPERTYPUT | bool | 1061 | Metadata Only | Reviewed | Not Tested | |
| PromptForSummaryInfo | PROPERTYGET/PROPERTYPUT | bool | 1062 | Metadata Only | Reviewed | Not Tested | |
| EnableTipWizard | PROPERTYGET/PROPERTYPUT | bool | 1064 | Metadata Only | Reviewed | Not Tested | |
| OnSheetDeactivate | PROPERTYGET/PROPERTYPUT | String | 1081 | Metadata Only | Reviewed | Not Tested | |
| EnableCancelKey | PROPERTYGET/PROPERTYPUT | XlEnableCancelKey | 1096 | Metadata Only | Reviewed | Not Tested | |
| MoveAfterReturnDirection | PROPERTYGET/PROPERTYPUT | XlDirection | 1144 | Metadata Only | Reviewed | Not Tested | |
| AutoCorrect | PROPERTYGET | AutoCorrect | 1145 | Metadata Only | Reviewed | Not Tested | |
| Cursor | PROPERTYGET/PROPERTYPUT | XlMousePointer | 1161 | Metadata Only | Reviewed | Not Tested | |
| EnableAutoComplete | PROPERTYGET/PROPERTYPUT | bool | 1179 | Metadata Only | Reviewed | Not Tested | |
| EnableAnimations | PROPERTYGET/PROPERTYPUT | bool | 1180 | Metadata Only | Reviewed | Not Tested | |
| DisplayCommentIndicator | PROPERTYGET/PROPERTYPUT | XlCommentDisplayMode | 1196 | Metadata Only | Reviewed | Not Tested | |
| EnableSound | PROPERTYGET/PROPERTYPUT | bool | 1197 | Metadata Only | Reviewed | Not Tested | |
| FileSearch | PROPERTYGET | FileSearch | 1200 | Metadata Only | Reviewed | Not Tested | |
| FileFind | PROPERTYGET | IFind | 1201 | Metadata Only | Reviewed | Not Tested | |
| RecentFiles | PROPERTYGET | RecentFiles | 1202 | Metadata Only | Reviewed | Not Tested | |
| ODBCErrors | PROPERTYGET | ODBCErrors | 1203 | Metadata Only | Reviewed | Not Tested | |
| ODBCTimeout | PROPERTYGET/PROPERTYPUT | i32 | 1204 | Metadata Only | Reviewed | Not Tested | |
| PivotTableSelection | PROPERTYGET/PROPERTYPUT | bool | 1205 | Metadata Only | Reviewed | Not Tested | |
| RollZoom | PROPERTYGET/PROPERTYPUT | bool | 1206 | Metadata Only | Reviewed | Not Tested | |
| ShowChartTipNames | PROPERTYGET/PROPERTYPUT | bool | 1207 | Metadata Only | Reviewed | Not Tested | |
| ShowChartTipValues | PROPERTYGET/PROPERTYPUT | bool | 1208 | Metadata Only | Reviewed | Not Tested | |
| DefaultSaveFormat | PROPERTYGET/PROPERTYPUT | XlFileFormat | 1209 | Metadata Only | Reviewed | Not Tested | |
| UserControl | PROPERTYGET/PROPERTYPUT | bool | 1210 | Metadata Only | Reviewed | Not Tested | |
| VBE | PROPERTYGET | VBE | 1211 | Metadata Only | Reviewed | Not Tested | |
| EnableEvents | PROPERTYGET/PROPERTYPUT | bool | 1212 | Metadata Only | Reviewed | Not Tested | |
| DisplayInfoWindow | PROPERTYGET/PROPERTYPUT | bool | 1213 | Metadata Only | Reviewed | Not Tested | |
| Assistant | PROPERTYGET | Assistant | 1438 | Metadata Only | Reviewed | Not Tested | |
| CommandBars | PROPERTYGET | CommandBars | 1439 | Metadata Only | Reviewed | Not Tested | |
| WorksheetFunction | PROPERTYGET | WorksheetFunction | 1440 | Metadata Only | Reviewed | Not Tested | |
| NewWorkbook | PROPERTYGET | NewFile | 1565 | Metadata Only | Reviewed | Not Tested | |
| ExtendList | PROPERTYGET/PROPERTYPUT | bool | 1793 | Metadata Only | Reviewed | Not Tested | |
| OLEDBErrors | PROPERTYGET | OLEDBErrors | 1794 | Metadata Only | Reviewed | Not Tested | |
| COMAddIns | PROPERTYGET | COMAddIns | 1796 | Metadata Only | Reviewed | Not Tested | |
| DefaultWebOptions | PROPERTYGET | DefaultWebOptions | 1797 | Metadata Only | Reviewed | Not Tested | |
| ProductCode | PROPERTYGET | String | 1798 | Metadata Only | Reviewed | Not Tested | |
| UserLibraryPath | PROPERTYGET | String | 1799 | Metadata Only | Reviewed | Not Tested | |
| AutoPercentEntry | PROPERTYGET/PROPERTYPUT | bool | 1800 | Metadata Only | Reviewed | Not Tested | |
| LanguageSettings | PROPERTYGET | LanguageSettings | 1801 | Metadata Only | Reviewed | Not Tested | |
| Dummy101 | PROPERTYGET | Object | 1802 | Metadata Only | Reviewed | Not Tested | |
| AnswerWizard | PROPERTYGET | AnswerWizard | 1804 | Metadata Only | Reviewed | Not Tested | |
| CalculationVersion | PROPERTYGET | i32 | 1806 | Metadata Only | Reviewed | Not Tested | |
| ShowWindowsInTaskbar | PROPERTYGET/PROPERTYPUT | bool | 1807 | Metadata Only | Reviewed | Not Tested | |
| FeatureInstall | PROPERTYGET/PROPERTYPUT | MsoFeatureInstall | 1808 | Metadata Only | Reviewed | Not Tested | |
| DecimalSeparator | PROPERTYGET/PROPERTYPUT | String | 1809 | Metadata Only | Reviewed | Not Tested | |
| ThousandsSeparator | PROPERTYGET/PROPERTYPUT | String | 1810 | Metadata Only | Reviewed | Not Tested | |
| Ready | PROPERTYGET | bool | 1932 | Metadata Only | Reviewed | Not Tested | |
| FindFormat | PROPERTYGET/PROPERTYPUTREF | CellFormat | 1934 | Metadata Only | Reviewed | Not Tested | |
| ReplaceFormat | PROPERTYGET/PROPERTYPUTREF | CellFormat | 1935 | Metadata Only | Reviewed | Not Tested | |
| UsedObjects | PROPERTYGET | UsedObjects | 1936 | Metadata Only | Reviewed | Not Tested | |
| CalculationState | PROPERTYGET | XlCalculationState | 1937 | Metadata Only | Reviewed | Not Tested | |
| CalculationInterruptKey | PROPERTYGET/PROPERTYPUT | XlCalculationInterruptKey | 1938 | Metadata Only | Reviewed | Not Tested | |
| Watches | PROPERTYGET | Watches | 1939 | Metadata Only | Reviewed | Not Tested | |
| DisplayFunctionToolTips | PROPERTYGET/PROPERTYPUT | bool | 1940 | Metadata Only | Reviewed | Not Tested | |
| AutomationSecurity | PROPERTYGET/PROPERTYPUT | MsoAutomationSecurity | 1941 | Metadata Only | Reviewed | Not Tested | |
| FileDialog | PROPERTYGET | FileDialog | 1942 | Metadata Only | Reviewed | Not Tested | |
| DisplayPasteOptions | PROPERTYGET/PROPERTYPUT | bool | 1946 | Metadata Only | Reviewed | Not Tested | |
| DisplayInsertOptions | PROPERTYGET/PROPERTYPUT | bool | 1947 | Metadata Only | Reviewed | Not Tested | |
| GenerateGetPivotData | PROPERTYGET/PROPERTYPUT | bool | 1948 | Metadata Only | Reviewed | Not Tested | |
| AutoRecover | PROPERTYGET | AutoRecover | 1949 | Metadata Only | Reviewed | Not Tested | |
| Hwnd | PROPERTYGET | i32 | 1950 | Metadata Only | Reviewed | Not Tested | |
| Hinstance | PROPERTYGET | i32 | 1951 | Metadata Only | Reviewed | Not Tested | |
| ErrorCheckingOptions | PROPERTYGET | ErrorCheckingOptions | 1954 | Metadata Only | Reviewed | Not Tested | |
| AutoFormatAsYouTypeReplaceHyperlinks | PROPERTYGET/PROPERTYPUT | bool | 1955 | Metadata Only | Reviewed | Not Tested | |
| SmartTagRecognizers | PROPERTYGET | SmartTagRecognizers | 1956 | Metadata Only | Reviewed | Not Tested | |
| SpellingOptions | PROPERTYGET | SpellingOptions | 1957 | Metadata Only | Reviewed | Not Tested | |
| Speech | PROPERTYGET | Speech | 1958 | Metadata Only | Reviewed | Not Tested | |
| MapPaperSize | PROPERTYGET/PROPERTYPUT | bool | 1959 | Metadata Only | Reviewed | Not Tested | |
| ShowStartupDialog | PROPERTYGET/PROPERTYPUT | bool | 1960 | Metadata Only | Reviewed | Not Tested | |
| UseSystemSeparators | PROPERTYGET/PROPERTYPUT | bool | 1961 | Metadata Only | Reviewed | Not Tested | |
| ThisCell | PROPERTYGET | Range | 1962 | Metadata Only | Reviewed | Not Tested | |
| RTD | PROPERTYGET | RTD | 1963 | Metadata Only | Reviewed | Not Tested | |
| DisplayDocumentActionTaskPane | PROPERTYGET/PROPERTYPUT | bool | 2251 | Metadata Only | Reviewed | Not Tested | |
| ArbitraryXMLSupportAvailable | PROPERTYGET | bool | 2254 | Metadata Only | Reviewed | Not Tested | |
| MeasurementUnit | PROPERTYGET/PROPERTYPUT | i32 | 2375 | Metadata Only | Reviewed | Not Tested | |
| ShowSelectionFloaties | PROPERTYGET/PROPERTYPUT | bool | 2376 | Metadata Only | Reviewed | Not Tested | |
| ShowMenuFloaties | PROPERTYGET/PROPERTYPUT | bool | 2377 | Metadata Only | Reviewed | Not Tested | |
| ShowDevTools | PROPERTYGET/PROPERTYPUT | bool | 2378 | Metadata Only | Reviewed | Not Tested | |
| EnableLivePreview | PROPERTYGET/PROPERTYPUT | bool | 2379 | Metadata Only | Reviewed | Not Tested | |
| DisplayDocumentInformationPanel | PROPERTYGET/PROPERTYPUT | bool | 2380 | Metadata Only | Reviewed | Not Tested | |
| AlwaysUseClearType | PROPERTYGET/PROPERTYPUT | bool | 2381 | Metadata Only | Reviewed | Not Tested | |
| WarnOnFunctionNameConflict | PROPERTYGET/PROPERTYPUT | bool | 2382 | Metadata Only | Reviewed | Not Tested | |
| FormulaBarHeight | PROPERTYGET/PROPERTYPUT | i32 | 2383 | Metadata Only | Reviewed | Not Tested | |
| DisplayFormulaAutoComplete | PROPERTYGET/PROPERTYPUT | bool | 2384 | Metadata Only | Reviewed | Not Tested | |
| GenerateTableRefs | PROPERTYGET/PROPERTYPUT | XlGenerateTableRefs | 2385 | Metadata Only | Reviewed | Not Tested | |
| Assistance | PROPERTYGET | IAssistance | 2386 | Metadata Only | Reviewed | Not Tested | |
| EnableLargeOperationAlert | PROPERTYGET/PROPERTYPUT | bool | 2388 | Metadata Only | Reviewed | Not Tested | |
| LargeOperationCellThousandCount | PROPERTYGET/PROPERTYPUT | i32 | 2389 | Metadata Only | Reviewed | Not Tested | |
| DeferAsyncQueries | PROPERTYGET/PROPERTYPUT | bool | 2390 | Metadata Only | Reviewed | Not Tested | |
| MultiThreadedCalculation | PROPERTYGET | MultiThreadedCalculation | 2391 | Metadata Only | Reviewed | Not Tested | |
| ActiveEncryptionSession | PROPERTYGET | i32 | 2394 | Metadata Only | Reviewed | Not Tested | |
| HighQualityModeForGraphics | PROPERTYGET/PROPERTYPUT | bool | 2395 | Metadata Only | Reviewed | Not Tested | |
| FileExportConverters | PROPERTYGET | FileExportConverters | 2768 | Metadata Only | Reviewed | Not Tested | |
| SmartArtLayouts | PROPERTYGET | SmartArtLayouts | 2772 | Metadata Only | Reviewed | Not Tested | |
| SmartArtQuickStyles | PROPERTYGET | SmartArtQuickStyles | 2773 | Metadata Only | Reviewed | Not Tested | |
| SmartArtColors | PROPERTYGET | SmartArtColors | 2774 | Metadata Only | Reviewed | Not Tested | |
| AddIns2 | PROPERTYGET | AddIns2 | 2775 | Metadata Only | Reviewed | Not Tested | |
| PrintCommunication | PROPERTYGET/PROPERTYPUT | bool | 2776 | Metadata Only | Reviewed | Not Tested | |
| UseClusterConnector | PROPERTYGET/PROPERTYPUT | bool | 2778 | Metadata Only | Reviewed | Not Tested | |
| ClusterConnector | PROPERTYGET/PROPERTYPUT | String | 2779 | Metadata Only | Reviewed | Not Tested | |
| Quitting | PROPERTYGET | bool | 2780 | Metadata Only | Reviewed | Not Tested | |
| Dummy22 | PROPERTYGET/PROPERTYPUT | bool | 2781 | Metadata Only | Reviewed | Not Tested | |
| Dummy23 | PROPERTYGET/PROPERTYPUT | bool | 2782 | Metadata Only | Reviewed | Not Tested | |
| ProtectedViewWindows | PROPERTYGET | ProtectedViewWindows | 2783 | Metadata Only | Reviewed | Not Tested | |
| ActiveProtectedViewWindow | PROPERTYGET | ProtectedViewWindow | 2784 | Metadata Only | Reviewed | Not Tested | |
| IsSandboxed | PROPERTYGET | bool | 2785 | Metadata Only | Reviewed | Not Tested | |
| SaveISO8601Dates | PROPERTYGET/PROPERTYPUT | bool | 2786 | Metadata Only | Reviewed | Not Tested | |
| HinstancePtr | PROPERTYGET | AutomationValue | 2787 | Metadata Only | Reviewed | Not Tested | |
| FileValidation | PROPERTYGET/PROPERTYPUT | MsoFileValidationMode | 2788 | Metadata Only | Reviewed | Not Tested | |
| FileValidationPivot | PROPERTYGET/PROPERTYPUT | XlFileValidationPivotMode | 2789 | Metadata Only | Reviewed | Not Tested | |
| ShowQuickAnalysis | PROPERTYGET/PROPERTYPUT | bool | 2994 | Metadata Only | Reviewed | Not Tested | |
| QuickAnalysis | PROPERTYGET | QuickAnalysis | 2995 | Metadata Only | Reviewed | Not Tested | |
| FlashFill | PROPERTYGET/PROPERTYPUT | bool | 2996 | Metadata Only | Reviewed | Not Tested | |
| EnableMacroAnimations | PROPERTYGET/PROPERTYPUT | bool | 2997 | Metadata Only | Reviewed | Not Tested | |
| ChartDataPointTrack | PROPERTYGET/PROPERTYPUT | bool | 2998 | Metadata Only | Reviewed | Not Tested | |
| FlashFillMode | PROPERTYGET/PROPERTYPUT | bool | 2999 | Metadata Only | Reviewed | Not Tested | |
| MergeInstances | PROPERTYGET/PROPERTYPUT | bool | 3000 | Metadata Only | Reviewed | Not Tested | |
| EnableCheckFileExtensions | PROPERTYGET/PROPERTYPUT | bool | 3158 | Metadata Only | Reviewed | Not Tested | |
| DefaultPivotTableLayoutOptions | PROPERTYGET | DefaultPivotTableLayoutOptions | 3271 | Metadata Only | Reviewed | Not Tested | |
| ShowConvertToDataType | PROPERTYGET/PROPERTYPUT | bool | 3311 | Metadata Only | Reviewed | Not Tested | |
| TruncateLeadingZeros | PROPERTYGET/PROPERTYPUT | bool | 3312 | Metadata Only | Reviewed | Not Tested | |
| TruncateLargeNumbers | PROPERTYGET/PROPERTYPUT | bool | 3313 | Metadata Only | Reviewed | Not Tested | |
| ConvertNumbersWithECharacter | PROPERTYGET/PROPERTYPUT | bool | 3314 | Metadata Only | Reviewed | Not Tested | |
| CSVDisplayNumberConversionWarning | PROPERTYGET/PROPERTYPUT | bool | 3315 | Metadata Only | Reviewed | Not Tested | |
| CSVKeepColumnAsTextIfMultipleEntriesAreText | PROPERTYGET/PROPERTYPUT | bool | 3316 | Metadata Only | Reviewed | Not Tested | |
| DataPrivacyOptions | PROPERTYGET | DataPrivacyOptions | 3317 | Metadata Only | Reviewed | Not Tested | |
| SensitivityLabelPolicy | PROPERTYGET | SensitivityLabelPolicy | 3365 | Metadata Only | Reviewed | Not Tested | |
| FormatStaleValues | PROPERTYGET/PROPERTYPUT | bool | 3401 | Metadata Only | Reviewed | Not Tested | |
| MaxSupportedCompatibilityVersion | PROPERTYGET | i32 | 3417 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---:|---|---|---|---|
| _Evaluate | AutomationValue | 1 | -5 | Metadata Only | Reviewed | Not Tested | |
| Evaluate | AutomationValue | 1 | 1 | Metadata Only | Reviewed | Not Tested | |
| RegisterXLL | bool | 1 | 30 | Metadata Only | Reviewed | Not Tested | |
| _WSFunction | AutomationValue | 30 | 169 | Metadata Only | Reviewed | Not Tested | |
| SaveWorkspace | Unknown | 1 | 212 | Metadata Only | Reviewed | Not Tested | |
| AddChartAutoFormat | Unknown | 3 | 216 | Metadata Only | Reviewed | Not Tested | |
| DeleteChartAutoFormat | Unknown | 1 | 217 | Metadata Only | Reviewed | Not Tested | |
| SetDefaultChart | Unknown | 2 | 219 | Metadata Only | Reviewed | Not Tested | |
| Run | AutomationValue | 31 | 259 | Metadata Only | Reviewed | Not Tested | |
| Calculate | Unknown | 0 | 279 | Metadata Only | Reviewed | Not Tested | |
| Save | Unknown | 1 | 283 | Metadata Only | Reviewed | Not Tested | |
| Repeat | Unknown | 0 | 301 | Metadata Only | Reviewed | Not Tested | |
| Quit | Unknown | 0 | 302 | Implemented | Reviewed | Live Tested | |
| Undo | Unknown | 0 | 303 | Metadata Only | Reviewed | Not Tested | |
| ConvertFormula | AutomationValue | 5 | 325 | Metadata Only | Reviewed | Not Tested | |
| DDEExecute | Unknown | 2 | 333 | Metadata Only | Reviewed | Not Tested | |
| DDEInitiate | i32 | 2 | 334 | Metadata Only | Reviewed | Not Tested | |
| DDEPoke | Unknown | 3 | 335 | Metadata Only | Reviewed | Not Tested | |
| DDERequest | AutomationValue | 2 | 336 | Metadata Only | Reviewed | Not Tested | |
| DDETerminate | Unknown | 1 | 337 | Metadata Only | Reviewed | Not Tested | |
| DoubleClick | Unknown | 0 | 349 | Metadata Only | Reviewed | Not Tested | |
| ExecuteExcel4Macro | AutomationValue | 1 | 350 | Metadata Only | Reviewed | Not Tested | |
| Help | Unknown | 2 | 354 | Metadata Only | Reviewed | Not Tested | |
| InputBox | AutomationValue | 8 | 357 | Metadata Only | Reviewed | Not Tested | |
| SendKeys | Unknown | 2 | 383 | Metadata Only | Reviewed | Not Tested | |
| _Wait | Unknown | 1 | 393 | Metadata Only | Reviewed | Not Tested | |
| Goto | Unknown | 2 | 475 | Metadata Only | Reviewed | Not Tested | |
| CheckSpelling | bool | 3 | 505 | Metadata Only | Reviewed | Not Tested | |
| OnTime | Unknown | 4 | 624 | Metadata Only | Reviewed | Not Tested | |
| OnKey | Unknown | 2 | 626 | Metadata Only | Reviewed | Not Tested | |
| Intersect | Range | 30 | 766 | Metadata Only | Reviewed | Not Tested | |
| OnRepeat | Unknown | 2 | 769 | Metadata Only | Reviewed | Not Tested | |
| OnUndo | Unknown | 2 | 770 | Metadata Only | Reviewed | Not Tested | |
| RecordMacro | Unknown | 2 | 773 | Metadata Only | Reviewed | Not Tested | |
| Union | Range | 30 | 779 | Metadata Only | Reviewed | Not Tested | |
| AddCustomList | Unknown | 2 | 780 | Metadata Only | Reviewed | Not Tested | |
| DeleteCustomList | Unknown | 1 | 783 | Metadata Only | Reviewed | Not Tested | |
| GetCustomListNum | i32 | 1 | 785 | Metadata Only | Reviewed | Not Tested | |
| GetCustomListContents | AutomationValue | 1 | 786 | Metadata Only | Reviewed | Not Tested | |
| Volatile | Unknown | 1 | 788 | Metadata Only | Reviewed | Not Tested | |
| _Run2 | AutomationValue | 31 | 806 | Metadata Only | Reviewed | Not Tested | |
| ResetTipWizard | Unknown | 0 | 928 | Metadata Only | Reviewed | Not Tested | |
| MailLogon | Unknown | 3 | 943 | Metadata Only | Reviewed | Not Tested | |
| MailLogoff | Unknown | 0 | 945 | Metadata Only | Reviewed | Not Tested | |
| NextLetter | Workbook | 0 | 972 | Metadata Only | Reviewed | Not Tested | |
| _FindFile | Unknown | 0 | 1068 | Metadata Only | Reviewed | Not Tested | |
| GetOpenFilename | AutomationValue | 5 | 1075 | Metadata Only | Reviewed | Not Tested | |
| GetSaveAsFilename | AutomationValue | 5 | 1076 | Metadata Only | Reviewed | Not Tested | |
| CentimetersToPoints | f64 | 1 | 1086 | Metadata Only | Reviewed | Not Tested | |
| InchesToPoints | f64 | 1 | 1087 | Metadata Only | Reviewed | Not Tested | |
| ActivateMicrosoftApp | Unknown | 1 | 1095 | Metadata Only | Reviewed | Not Tested | |
| _MacroOptions | Unknown | 10 | 1135 | Metadata Only | Reviewed | Not Tested | |
| Wait | bool | 1 | 1770 | Metadata Only | Reviewed | Not Tested | |
| FindFile | bool | 0 | 1771 | Metadata Only | Reviewed | Not Tested | |
| Dummy1 | AutomationValue | 4 | 1782 | Metadata Only | Reviewed | Not Tested | |
| Dummy2 | AutomationValue | 8 | 1783 | Metadata Only | Reviewed | Not Tested | |
| Dummy3 | AutomationValue | 0 | 1784 | Metadata Only | Reviewed | Not Tested | |
| Dummy4 | AutomationValue | 15 | 1785 | Metadata Only | Reviewed | Not Tested | |
| Dummy5 | AutomationValue | 13 | 1786 | Metadata Only | Reviewed | Not Tested | |
| Dummy6 | AutomationValue | 0 | 1787 | Metadata Only | Reviewed | Not Tested | |
| Dummy7 | AutomationValue | 0 | 1788 | Metadata Only | Reviewed | Not Tested | |
| Dummy8 | AutomationValue | 1 | 1789 | Metadata Only | Reviewed | Not Tested | |
| Dummy9 | AutomationValue | 0 | 1790 | Metadata Only | Reviewed | Not Tested | |
| Dummy10 | bool | 1 | 1791 | Metadata Only | Reviewed | Not Tested | |
| Dummy11 | Unknown | 0 | 1792 | Metadata Only | Reviewed | Not Tested | |
| GetPhonetic | String | 1 | 1795 | Metadata Only | Reviewed | Not Tested | |
| Dummy12 | Unknown | 2 | 1803 | Metadata Only | Reviewed | Not Tested | |
| CalculateFull | Unknown | 0 | 1805 | Metadata Only | Reviewed | Not Tested | |
| Dummy13 | AutomationValue | 30 | 1933 | Metadata Only | Reviewed | Not Tested | |
| Dummy14 | Unknown | 0 | 1944 | Metadata Only | Reviewed | Not Tested | |
| CalculateFullRebuild | Unknown | 0 | 1945 | Metadata Only | Reviewed | Not Tested | |
| CheckAbort | Unknown | 1 | 1952 | Metadata Only | Reviewed | Not Tested | |
| DisplayXMLSourcePane | Unknown | 1 | 2252 | Metadata Only | Reviewed | Not Tested | |
| Support | AutomationValue | 3 | 2255 | Metadata Only | Reviewed | Not Tested | |
| Dummy20 | AutomationValue | 1 | 2373 | Metadata Only | Reviewed | Not Tested | |
| CalculateUntilAsyncQueriesDone | Unknown | 0 | 2387 | Metadata Only | Reviewed | Not Tested | |
| SharePointVersion | i32 | 1 | 2392 | Metadata Only | Reviewed | Not Tested | |
| MacroOptions | Unknown | 11 | 2770 | Metadata Only | Reviewed | Not Tested | |
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
