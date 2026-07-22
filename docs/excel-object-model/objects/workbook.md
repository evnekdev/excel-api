# Workbook

## Summary

An Excel workbook object. The initial crate supports basic identity, saved-state, and explicit close-without-saving operations.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `_Workbook` |
| GUID | `{000208da-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4160 |
| Crate type | `excel_com::Workbook` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Worksheets` | `excel.worksheets` | Implemented |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| _CodeName | PROPERTYGET/PROPERTYPUT | String | declared | -2147418112 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET | String | declared | 110 | Implemented | Reviewed | Live Tested | |
| Charts | PROPERTYGET | Sheets | declared | 121 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| ActiveChart | PROPERTYGET | Chart | declared | 183 | Metadata Only | Reviewed | Not Tested | |
| Title | PROPERTYGET/PROPERTYPUT | String | declared | 199 | Metadata Only | Reviewed | Not Tested | |
| Colors | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 286 | Metadata Only | Reviewed | Not Tested | |
| CreateBackup | PROPERTYGET | bool | declared | 287 | Metadata Only | Reviewed | Not Tested | |
| FileFormat | PROPERTYGET | XlFileFormat | declared | 288 | Implemented | Reviewed | Live Tested | |
| FullName | PROPERTYGET | String | declared | 289 | Implemented | Reviewed | Live Tested | |
| HasPassword | PROPERTYGET | bool | declared | 290 | Metadata Only | Reviewed | Not Tested | |
| Path | PROPERTYGET | String | declared | 291 | Implemented | Reviewed | Live Tested | |
| ProtectWindows | PROPERTYGET | bool | declared | 295 | Metadata Only | Reviewed | Not Tested | |
| ReadOnly | PROPERTYGET | bool | declared | 296 | Implemented | Reviewed | Live Tested | |
| _ReadOnlyRecommended | PROPERTYGET | bool | declared | 297 | Metadata Only | Reviewed | Not Tested | |
| Saved | PROPERTYGET/PROPERTYPUT | bool | declared | 298 | Implemented | Reviewed | Live Tested | |
| WriteReserved | PROPERTYGET | bool | declared | 299 | Metadata Only | Reviewed | Not Tested | |
| WriteReservedBy | PROPERTYGET | String | declared | 300 | Metadata Only | Reviewed | Not Tested | |
| ActiveSheet | PROPERTYGET | Object | declared | 307 | Metadata Only | Reviewed | Not Tested | |
| Date1904 | PROPERTYGET/PROPERTYPUT | bool | declared | 403 | Metadata Only | Reviewed | Not Tested | |
| DisplayDrawingObjects | PROPERTYGET/PROPERTYPUT | XlDisplayDrawingObjects | declared | 404 | Metadata Only | Reviewed | Not Tested | |
| PrecisionAsDisplayed | PROPERTYGET/PROPERTYPUT | bool | declared | 405 | Metadata Only | Reviewed | Not Tested | |
| SaveLinkValues | PROPERTYGET/PROPERTYPUT | bool | declared | 406 | Metadata Only | Reviewed | Not Tested | |
| UpdateRemoteReferences | PROPERTYGET/PROPERTYPUT | bool | declared | 411 | Metadata Only | Reviewed | Not Tested | |
| Password | PROPERTYGET/PROPERTYPUT | String | declared | 429 | Metadata Only | Reviewed | Not Tested | |
| Windows | PROPERTYGET | Windows | declared | 430 | Metadata Only | Reviewed | Not Tested | |
| Names | PROPERTYGET | Names | declared | 442 | Metadata Only | Reviewed | Not Tested | |
| Sheets | PROPERTYGET | Sheets | declared | 485 | Metadata Only | Reviewed | Not Tested | |
| Styles | PROPERTYGET | Styles | declared | 493 | Metadata Only | Reviewed | Not Tested | |
| Worksheets | PROPERTYGET | Sheets | declared | 494 | Implemented | Reviewed | Live Tested | |
| Author | PROPERTYGET/PROPERTYPUT | String | declared | 574 | Metadata Only | Reviewed | Not Tested | |
| Comments | PROPERTYGET/PROPERTYPUT | String | declared | 575 | Metadata Only | Reviewed | Not Tested | |
| Keywords | PROPERTYGET/PROPERTYPUT | String | declared | 577 | Metadata Only | Reviewed | Not Tested | |
| Excel4MacroSheets | PROPERTYGET | Sheets | declared | 579 | Metadata Only | Reviewed | Not Tested | |
| Excel4IntlMacroSheets | PROPERTYGET | Sheets | declared | 581 | Metadata Only | Reviewed | Not Tested | |
| Modules | PROPERTYGET | Sheets | declared | 582 | Metadata Only | Reviewed | Not Tested | |
| ProtectStructure | PROPERTYGET | bool | declared | 588 | Metadata Only | Reviewed | Not Tested | |
| PivotTables | PROPERTYGET | Object | declared | 690 | Metadata Only | Reviewed | Not Tested | |
| DialogSheets | PROPERTYGET | Sheets | declared | 764 | Metadata Only | Reviewed | Not Tested | |
| UpdateLinks | PROPERTYGET/PROPERTYPUT | XlUpdateLinks | declared | 864 | Metadata Only | Reviewed | Not Tested | |
| RoutingSlip | PROPERTYGET | RoutingSlip | declared | 949 | Metadata Only | Reviewed | Not Tested | |
| HasRoutingSlip | PROPERTYGET/PROPERTYPUT | bool | declared | 950 | Metadata Only | Reviewed | Not Tested | |
| Routed | PROPERTYGET | bool | declared | 951 | Metadata Only | Reviewed | Not Tested | |
| Subject | PROPERTYGET/PROPERTYPUT | String | declared | 953 | Metadata Only | Reviewed | Not Tested | |
| HasMailer | PROPERTYGET/PROPERTYPUT | bool | declared | 976 | Metadata Only | Reviewed | Not Tested | |
| Mailer | PROPERTYGET | Mailer | declared | 979 | Metadata Only | Reviewed | Not Tested | |
| OnSheetActivate | PROPERTYGET/PROPERTYPUT | String | declared | 1031 | Metadata Only | Reviewed | Not Tested | |
| OnSheetDeactivate | PROPERTYGET/PROPERTYPUT | String | declared | 1081 | Metadata Only | Reviewed | Not Tested | |
| WritePassword | PROPERTYGET/PROPERTYPUT | String | declared | 1128 | Metadata Only | Reviewed | Not Tested | |
| MultiUserEditing | PROPERTYGET | bool | declared | 1169 | Metadata Only | Reviewed | Not Tested | |
| ShowConflictHistory | PROPERTYGET/PROPERTYPUT | bool | declared | 1171 | Metadata Only | Reviewed | Not Tested | |
| RevisionNumber | PROPERTYGET | i32 | declared | 1172 | Metadata Only | Reviewed | Not Tested | |
| UserStatus | PROPERTYGET | AutomationValue | declared | 1173 | Metadata Only | Reviewed | Not Tested | |
| ConflictResolution | PROPERTYGET/PROPERTYPUT | XlSaveConflictResolution | declared | 1175 | Metadata Only | Reviewed | Not Tested | |
| BuiltinDocumentProperties | PROPERTYGET | Object | declared | 1176 | Metadata Only | Reviewed | Not Tested | |
| CustomDocumentProperties | PROPERTYGET | Object | declared | 1177 | Metadata Only | Reviewed | Not Tested | |
| OnSave | PROPERTYGET/PROPERTYPUT | String | declared | 1178 | Metadata Only | Reviewed | Not Tested | |
| Container | PROPERTYGET | Object | declared | 1190 | Metadata Only | Reviewed | Not Tested | |
| UserControl | PROPERTYGET/PROPERTYPUT | bool | declared | 1210 | Metadata Only | Reviewed | Not Tested | |
| CodeName | PROPERTYGET | String | declared | 1373 | Metadata Only | Reviewed | Not Tested | |
| CommandBars | PROPERTYGET | CommandBars | declared | 1439 | Metadata Only | Reviewed | Not Tested | |
| AcceptLabelsInFormulas | PROPERTYGET/PROPERTYPUT | bool | declared | 1441 | Metadata Only | Reviewed | Not Tested | |
| AutoUpdateFrequency | PROPERTYGET/PROPERTYPUT | i32 | declared | 1442 | Metadata Only | Reviewed | Not Tested | |
| AutoUpdateSaveChanges | PROPERTYGET/PROPERTYPUT | bool | declared | 1443 | Metadata Only | Reviewed | Not Tested | |
| ChangeHistoryDuration | PROPERTYGET/PROPERTYPUT | i32 | declared | 1444 | Metadata Only | Reviewed | Not Tested | |
| IsAddin | PROPERTYGET/PROPERTYPUT | bool | declared | 1445 | Metadata Only | Reviewed | Not Tested | |
| PersonalViewListSettings | PROPERTYGET/PROPERTYPUT | bool | declared | 1447 | Metadata Only | Reviewed | Not Tested | |
| PersonalViewPrintSettings | PROPERTYGET/PROPERTYPUT | bool | declared | 1448 | Metadata Only | Reviewed | Not Tested | |
| CustomViews | PROPERTYGET | CustomViews | declared | 1456 | Metadata Only | Reviewed | Not Tested | |
| TemplateRemoveExtData | PROPERTYGET/PROPERTYPUT | bool | declared | 1457 | Metadata Only | Reviewed | Not Tested | |
| HighlightChangesOnScreen | PROPERTYGET/PROPERTYPUT | bool | declared | 1461 | Metadata Only | Reviewed | Not Tested | |
| KeepChangeHistory | PROPERTYGET/PROPERTYPUT | bool | declared | 1462 | Metadata Only | Reviewed | Not Tested | |
| ListChangesOnNewSheet | PROPERTYGET/PROPERTYPUT | bool | declared | 1463 | Metadata Only | Reviewed | Not Tested | |
| VBProject | PROPERTYGET | VBProject | declared | 1469 | Metadata Only | Reviewed | Not Tested | |
| IsInplace | PROPERTYGET | bool | declared | 1769 | Metadata Only | Reviewed | Not Tested | |
| CalculationVersion | PROPERTYGET | i32 | declared | 1806 | Metadata Only | Reviewed | Not Tested | |
| PublishObjects | PROPERTYGET | PublishObjects | declared | 1819 | Metadata Only | Reviewed | Not Tested | |
| WebOptions | PROPERTYGET | WebOptions | declared | 1820 | Metadata Only | Reviewed | Not Tested | |
| HTMLProject | PROPERTYGET | HTMLProject | declared | 1823 | Metadata Only | Reviewed | Not Tested | |
| EnvelopeVisible | PROPERTYGET/PROPERTYPUT | bool | declared | 1824 | Metadata Only | Reviewed | Not Tested | |
| VBASigned | PROPERTYGET | bool | declared | 1828 | Metadata Only | Reviewed | Not Tested | |
| FullNameURLEncoded | PROPERTYGET | String | declared | 1927 | Metadata Only | Reviewed | Not Tested | |
| ReadOnlyRecommended | PROPERTYGET/PROPERTYPUT | bool | declared | 2005 | Metadata Only | Reviewed | Not Tested | |
| ShowPivotTableFieldList | PROPERTYGET/PROPERTYPUT | bool | declared | 2046 | Metadata Only | Reviewed | Not Tested | |
| EnableAutoRecover | PROPERTYGET/PROPERTYPUT | bool | declared | 2049 | Metadata Only | Reviewed | Not Tested | |
| RemovePersonalInformation | PROPERTYGET/PROPERTYPUT | bool | declared | 2050 | Metadata Only | Reviewed | Not Tested | |
| PasswordEncryptionProvider | PROPERTYGET | String | declared | 2059 | Metadata Only | Reviewed | Not Tested | |
| PasswordEncryptionAlgorithm | PROPERTYGET | String | declared | 2060 | Metadata Only | Reviewed | Not Tested | |
| PasswordEncryptionKeyLength | PROPERTYGET | i32 | declared | 2061 | Metadata Only | Reviewed | Not Tested | |
| PasswordEncryptionFileProperties | PROPERTYGET | bool | declared | 2063 | Metadata Only | Reviewed | Not Tested | |
| SmartTagOptions | PROPERTYGET | SmartTagOptions | declared | 2064 | Metadata Only | Reviewed | Not Tested | |
| Permission | PROPERTYGET | Permission | declared | 2264 | Metadata Only | Reviewed | Not Tested | |
| SharedWorkspace | PROPERTYGET | SharedWorkspace | declared | 2265 | Metadata Only | Reviewed | Not Tested | |
| Sync | PROPERTYGET | Sync | declared | 2266 | Metadata Only | Reviewed | Not Tested | |
| XmlNamespaces | PROPERTYGET | XmlNamespaces | declared | 2268 | Metadata Only | Reviewed | Not Tested | |
| XmlMaps | PROPERTYGET | XmlMaps | declared | 2269 | Metadata Only | Reviewed | Not Tested | |
| SmartDocument | PROPERTYGET | SmartDocument | declared | 2273 | Metadata Only | Reviewed | Not Tested | |
| DocumentLibraryVersions | PROPERTYGET | DocumentLibraryVersions | declared | 2274 | Metadata Only | Reviewed | Not Tested | |
| InactiveListBorderVisible | PROPERTYGET/PROPERTYPUT | bool | declared | 2275 | Metadata Only | Reviewed | Not Tested | |
| DisplayInkComments | PROPERTYGET/PROPERTYPUT | bool | declared | 2276 | Metadata Only | Reviewed | Not Tested | |
| ContentTypeProperties | PROPERTYGET | MetaProperties | declared | 2512 | Metadata Only | Reviewed | Not Tested | |
| Connections | PROPERTYGET | Connections | declared | 2513 | Metadata Only | Reviewed | Not Tested | |
| Signatures | PROPERTYGET | SignatureSet | declared | 2516 | Metadata Only | Reviewed | Not Tested | |
| ServerPolicy | PROPERTYGET | ServerPolicy | declared | 2519 | Metadata Only | Reviewed | Not Tested | |
| DocumentInspectors | PROPERTYGET | DocumentInspectors | declared | 2521 | Metadata Only | Reviewed | Not Tested | |
| ServerViewableItems | PROPERTYGET | ServerViewableItems | declared | 2524 | Metadata Only | Reviewed | Not Tested | |
| TableStyles | PROPERTYGET | TableStyles | declared | 2525 | Metadata Only | Reviewed | Not Tested | |
| DefaultTableStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2526 | Metadata Only | Reviewed | Not Tested | |
| DefaultPivotTableStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2527 | Metadata Only | Reviewed | Not Tested | |
| CheckCompatibility | PROPERTYGET/PROPERTYPUT | bool | declared | 2528 | Metadata Only | Reviewed | Not Tested | |
| HasVBProject | PROPERTYGET | bool | declared | 2529 | Metadata Only | Reviewed | Not Tested | |
| CustomXMLParts | PROPERTYGET | CustomXMLParts | declared | 2530 | Metadata Only | Reviewed | Not Tested | |
| Final | PROPERTYGET/PROPERTYPUT | bool | declared | 2531 | Metadata Only | Reviewed | Not Tested | |
| Research | PROPERTYGET | Research | declared | 2532 | Metadata Only | Reviewed | Not Tested | |
| Theme | PROPERTYGET | OfficeTheme | declared | 2533 | Metadata Only | Reviewed | Not Tested | |
| Excel8CompatibilityMode | PROPERTYGET | bool | declared | 2535 | Metadata Only | Reviewed | Not Tested | |
| ConnectionsDisabled | PROPERTYGET | bool | declared | 2536 | Metadata Only | Reviewed | Not Tested | |
| ShowPivotChartActiveFields | PROPERTYGET/PROPERTYPUT | bool | declared | 2538 | Metadata Only | Reviewed | Not Tested | |
| IconSets | PROPERTYGET | IconSets | declared | 2539 | Metadata Only | Reviewed | Not Tested | |
| EncryptionProvider | PROPERTYGET/PROPERTYPUT | String | declared | 2540 | Metadata Only | Reviewed | Not Tested | |
| DoNotPromptForConvert | PROPERTYGET/PROPERTYPUT | bool | declared | 2541 | Metadata Only | Reviewed | Not Tested | |
| ForceFullCalculation | PROPERTYGET/PROPERTYPUT | bool | declared | 2542 | Metadata Only | Reviewed | Not Tested | |
| SlicerCaches | PROPERTYGET | SlicerCaches | declared | 2866 | Metadata Only | Reviewed | Not Tested | |
| ActiveSlicer | PROPERTYGET | Slicer | declared | 2867 | Metadata Only | Reviewed | Not Tested | |
| DefaultSlicerStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 2868 | Metadata Only | Reviewed | Not Tested | |
| AccuracyVersion | PROPERTYGET/PROPERTYPUT | i32 | declared | 2871 | Metadata Only | Reviewed | Not Tested | |
| ChartDataPointTrack | PROPERTYGET/PROPERTYPUT | bool | declared | 2998 | Metadata Only | Reviewed | Not Tested | |
| CaseSensitive | PROPERTYGET | bool | declared | 3056 | Metadata Only | Reviewed | Not Tested | |
| UseWholeCellCriteria | PROPERTYGET | bool | declared | 3057 | Metadata Only | Reviewed | Not Tested | |
| UseWildcards | PROPERTYGET | bool | declared | 3058 | Metadata Only | Reviewed | Not Tested | |
| Model | PROPERTYGET | Model | declared | 3059 | Metadata Only | Reviewed | Not Tested | |
| DefaultTimelineStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 3060 | Metadata Only | Reviewed | Not Tested | |
| WorkIdentity | PROPERTYGET/PROPERTYPUT | String | declared | 3173 | Metadata Only | Reviewed | Not Tested | |
| Queries | PROPERTYGET | Queries | declared | 3186 | Metadata Only | Reviewed | Not Tested | |
| AutoSaveOn | PROPERTYGET/PROPERTYPUT | bool | declared | 3232 | Metadata Only | Reviewed | Not Tested | |
| SensitivityLabel | PROPERTYGET | ISensitivityLabel | declared | 3379 | Metadata Only | Reviewed | Not Tested | |
| ExternalCodeServiceTimeout | PROPERTYGET/PROPERTYPUT | i32 | declared | 3406 | Metadata Only | Reviewed | Not Tested | |
| CompatibilityVersion | PROPERTYGET/PROPERTYPUT | i32 | declared | 3416 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| SaveCopyAs | Unknown | 1 | declared | 175 | Implemented | Reviewed | Live Tested | |
| Close | Unknown | 3 | declared | 277 | Implemented | Reviewed | Live Tested | |
| NewWindow | Window | 0 | declared | 280 | Metadata Only | Reviewed | Not Tested | |
| PrintPreview | Unknown | 1 | declared | 281 | Metadata Only | Reviewed | Not Tested | |
| _Protect | Unknown | 3 | declared | 282 | Metadata Only | Reviewed | Not Tested | |
| Save | Unknown | 0 | declared | 283 | Implemented | Reviewed | Live Tested | |
| __SaveAs | Unknown | 11 | declared | 284 | Metadata Only | Reviewed | Not Tested | |
| Unprotect | Unknown | 1 | declared | 285 | Metadata Only | Reviewed | Not Tested | |
| Activate | Unknown | 0 | declared | 304 | Metadata Only | Reviewed | Not Tested | |
| DeleteNumberFormat | Unknown | 1 | declared | 397 | Metadata Only | Reviewed | Not Tested | |
| RunAutoMacros | Unknown | 1 | declared | 634 | Metadata Only | Reviewed | Not Tested | |
| PivotTableWizard | Unknown | 16 | declared | 684 | Metadata Only | Reviewed | Not Tested | |
| ChangeLink | Unknown | 3 | declared | 802 | Metadata Only | Reviewed | Not Tested | |
| OpenLinks | Unknown | 3 | declared | 803 | Metadata Only | Reviewed | Not Tested | |
| UpdateLink | Unknown | 2 | declared | 804 | Metadata Only | Reviewed | Not Tested | |
| LinkInfo | AutomationValue | 4 | declared | 807 | Metadata Only | Reviewed | Not Tested | |
| LinkSources | AutomationValue | 1 | declared | 808 | Metadata Only | Reviewed | Not Tested | |
| SetLinkOnData | Unknown | 2 | declared | 809 | Metadata Only | Reviewed | Not Tested | |
| __PrintOut | Unknown | 7 | declared | 905 | Metadata Only | Reviewed | Not Tested | |
| Route | Unknown | 0 | declared | 946 | Metadata Only | Reviewed | Not Tested | |
| SendMail | Unknown | 3 | declared | 947 | Metadata Only | Reviewed | Not Tested | |
| ForwardMailer | Unknown | 0 | declared | 973 | Metadata Only | Reviewed | Not Tested | |
| Reply | Unknown | 0 | declared | 977 | Metadata Only | Reviewed | Not Tested | |
| ReplyAll | Unknown | 0 | declared | 978 | Metadata Only | Reviewed | Not Tested | |
| SendMailer | Unknown | 2 | declared | 980 | Metadata Only | Reviewed | Not Tested | |
| ChangeFileAccess | Unknown | 3 | declared | 989 | Metadata Only | Reviewed | Not Tested | |
| UpdateFromFile | Unknown | 0 | declared | 995 | Metadata Only | Reviewed | Not Tested | |
| Post | Unknown | 1 | declared | 1166 | Metadata Only | Reviewed | Not Tested | |
| ExclusiveAccess | bool | 0 | declared | 1168 | Metadata Only | Reviewed | Not Tested | |
| MergeWorkbook | Unknown | 1 | declared | 1446 | Metadata Only | Reviewed | Not Tested | |
| PivotCaches | PivotCaches | 0 | declared | 1449 | Metadata Only | Reviewed | Not Tested | |
| _ProtectSharing | Unknown | 6 | declared | 1450 | Metadata Only | Reviewed | Not Tested | |
| RefreshAll | Unknown | 0 | declared | 1452 | Metadata Only | Reviewed | Not Tested | |
| RemoveUser | Unknown | 1 | declared | 1453 | Metadata Only | Reviewed | Not Tested | |
| UnprotectSharing | Unknown | 1 | declared | 1455 | Metadata Only | Reviewed | Not Tested | |
| HighlightChangesOptions | Unknown | 3 | declared | 1458 | Metadata Only | Reviewed | Not Tested | |
| PurgeChangeHistoryNow | Unknown | 2 | declared | 1464 | Metadata Only | Reviewed | Not Tested | |
| AcceptAllChanges | Unknown | 3 | declared | 1466 | Metadata Only | Reviewed | Not Tested | |
| RejectAllChanges | Unknown | 3 | declared | 1467 | Metadata Only | Reviewed | Not Tested | |
| ResetColors | Unknown | 0 | declared | 1468 | Metadata Only | Reviewed | Not Tested | |
| FollowHyperlink | Unknown | 7 | declared | 1470 | Metadata Only | Reviewed | Not Tested | |
| AddToFavorites | Unknown | 0 | declared | 1476 | Metadata Only | Reviewed | Not Tested | |
| _PrintOut | Unknown | 8 | declared | 1772 | Metadata Only | Reviewed | Not Tested | |
| WebPagePreview | Unknown | 0 | declared | 1818 | Metadata Only | Reviewed | Not Tested | |
| ReloadAs | Unknown | 1 | declared | 1821 | Metadata Only | Reviewed | Not Tested | |
| sblt | Unknown | 1 | declared | 1826 | Metadata Only | Reviewed | Not Tested | |
| _SaveAs | Unknown | 12 | declared | 1925 | Metadata Only | Reviewed | Not Tested | |
| Protect | Unknown | 3 | declared | 2029 | Metadata Only | Reviewed | Not Tested | |
| Dummy17 | Unknown | 1 | declared | 2044 | Metadata Only | Reviewed | Not Tested | |
| BreakLink | Unknown | 2 | declared | 2047 | Metadata Only | Reviewed | Not Tested | |
| Dummy16 | Unknown | 0 | declared | 2048 | Metadata Only | Reviewed | Not Tested | |
| CheckIn | Unknown | 3 | declared | 2051 | Metadata Only | Reviewed | Not Tested | |
| CanCheckIn | bool | 0 | declared | 2053 | Metadata Only | Reviewed | Not Tested | |
| SendForReview | Unknown | 4 | declared | 2054 | Metadata Only | Reviewed | Not Tested | |
| ReplyWithChanges | Unknown | 1 | declared | 2057 | Metadata Only | Reviewed | Not Tested | |
| EndReview | Unknown | 0 | declared | 2058 | Metadata Only | Reviewed | Not Tested | |
| SetPasswordEncryptionOptions | Unknown | 4 | declared | 2062 | Metadata Only | Reviewed | Not Tested | |
| RecheckSmartTags | Unknown | 0 | declared | 2065 | Metadata Only | Reviewed | Not Tested | |
| SendFaxOverInternet | Unknown | 3 | declared | 2267 | Metadata Only | Reviewed | Not Tested | |
| XmlImport | XlXmlImportResult | 4 | declared | 2270 | Metadata Only | Reviewed | Not Tested | |
| XmlImportXml | XlXmlImportResult | 4 | declared | 2277 | Metadata Only | Reviewed | Not Tested | |
| SaveAsXMLData | Unknown | 2 | declared | 2278 | Metadata Only | Reviewed | Not Tested | |
| ToggleFormsDesign | Unknown | 0 | declared | 2279 | Metadata Only | Reviewed | Not Tested | |
| PrintOut | Unknown | 9 | declared | 2361 | Metadata Only | Reviewed | Not Tested | |
| _ExportAsFixedFormat | Unknown | 9 | declared | 2493 | Metadata Only | Reviewed | Not Tested | |
| RemoveDocumentInformation | Unknown | 1 | declared | 2514 | Metadata Only | Reviewed | Not Tested | |
| CheckInWithVersion | Unknown | 4 | declared | 2517 | Metadata Only | Reviewed | Not Tested | |
| LockServerFile | Unknown | 0 | declared | 2520 | Metadata Only | Reviewed | Not Tested | |
| GetWorkflowTasks | WorkflowTasks | 0 | declared | 2522 | Metadata Only | Reviewed | Not Tested | |
| GetWorkflowTemplates | WorkflowTemplates | 0 | declared | 2523 | Metadata Only | Reviewed | Not Tested | |
| ApplyTheme | Unknown | 1 | declared | 2534 | Metadata Only | Reviewed | Not Tested | |
| EnableConnections | Unknown | 0 | declared | 2537 | Metadata Only | Reviewed | Not Tested | |
| ProtectSharing | Unknown | 7 | declared | 2543 | Metadata Only | Reviewed | Not Tested | |
| Dummy26 | Unknown | 0 | declared | 2869 | Metadata Only | Reviewed | Not Tested | |
| Dummy27 | Unknown | 0 | declared | 2870 | Metadata Only | Reviewed | Not Tested | |
| CreateForecastSheet | Unknown | 10 | declared | 3167 | Metadata Only | Reviewed | Not Tested | |
| SaveAs | Unknown | 13 | declared | 3174 | Implemented | Reviewed | Live Tested | |
| ExportAsFixedFormat | Unknown | 10 | declared | 3175 | Metadata Only | Reviewed | Not Tested | |
| PublishToPBI | String | 3 | declared | 3257 | Metadata Only | Reviewed | Not Tested | |
| ConvertComments | Unknown | 0 | declared | 3279 | Metadata Only | Reviewed | Not Tested | |
| PublishToDocs | String | 3 | declared | 3334 | Metadata Only | Reviewed | Not Tested | |
| LookUpInDocs | PublishedDocs | 1 | declared | 3335 | Metadata Only | Reviewed | Not Tested | |
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
