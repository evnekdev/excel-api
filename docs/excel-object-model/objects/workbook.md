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
| Crate type | `excel_com::Workbook` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Live Tested |

## Relationships

| Relationship | Target | Status |
|---|---|---|
| `Application` | `excel.application` | Metadata Only |
| `Worksheets` | `excel.worksheets` | Metadata Only |

## Properties

| Property | Access | Type | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---:|---|---|---|---|
| _CodeName | PROPERTYGET/PROPERTYPUT | String | -2147418112 | Metadata Only | Reviewed | Not Tested | |
| Name | PROPERTYGET | String | 110 | Implemented | Reviewed | Live Tested | |
| Charts | PROPERTYGET | Sheets | 121 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | 150 | Metadata Only | Reviewed | Not Tested | |
| ActiveChart | PROPERTYGET | Chart | 183 | Metadata Only | Reviewed | Not Tested | |
| Title | PROPERTYGET/PROPERTYPUT | String | 199 | Metadata Only | Reviewed | Not Tested | |
| Colors | PROPERTYGET/PROPERTYPUT | AutomationValue | 286 | Metadata Only | Reviewed | Not Tested | |
| CreateBackup | PROPERTYGET | bool | 287 | Metadata Only | Reviewed | Not Tested | |
| FileFormat | PROPERTYGET | XlFileFormat | 288 | Metadata Only | Reviewed | Not Tested | |
| FullName | PROPERTYGET | String | 289 | Metadata Only | Reviewed | Not Tested | |
| HasPassword | PROPERTYGET | bool | 290 | Metadata Only | Reviewed | Not Tested | |
| Path | PROPERTYGET | String | 291 | Metadata Only | Reviewed | Not Tested | |
| ProtectWindows | PROPERTYGET | bool | 295 | Metadata Only | Reviewed | Not Tested | |
| ReadOnly | PROPERTYGET | bool | 296 | Metadata Only | Reviewed | Not Tested | |
| _ReadOnlyRecommended | PROPERTYGET | bool | 297 | Metadata Only | Reviewed | Not Tested | |
| Saved | PROPERTYGET/PROPERTYPUT | bool | 298 | Implemented | Reviewed | Live Tested | |
| WriteReserved | PROPERTYGET | bool | 299 | Metadata Only | Reviewed | Not Tested | |
| WriteReservedBy | PROPERTYGET | String | 300 | Metadata Only | Reviewed | Not Tested | |
| ActiveSheet | PROPERTYGET | Object | 307 | Metadata Only | Reviewed | Not Tested | |
| Date1904 | PROPERTYGET/PROPERTYPUT | bool | 403 | Metadata Only | Reviewed | Not Tested | |
| DisplayDrawingObjects | PROPERTYGET/PROPERTYPUT | XlDisplayDrawingObjects | 404 | Metadata Only | Reviewed | Not Tested | |
| PrecisionAsDisplayed | PROPERTYGET/PROPERTYPUT | bool | 405 | Metadata Only | Reviewed | Not Tested | |
| SaveLinkValues | PROPERTYGET/PROPERTYPUT | bool | 406 | Metadata Only | Reviewed | Not Tested | |
| UpdateRemoteReferences | PROPERTYGET/PROPERTYPUT | bool | 411 | Metadata Only | Reviewed | Not Tested | |
| Password | PROPERTYGET/PROPERTYPUT | String | 429 | Metadata Only | Reviewed | Not Tested | |
| Windows | PROPERTYGET | Windows | 430 | Metadata Only | Reviewed | Not Tested | |
| Names | PROPERTYGET | Names | 442 | Metadata Only | Reviewed | Not Tested | |
| Sheets | PROPERTYGET | Sheets | 485 | Metadata Only | Reviewed | Not Tested | |
| Styles | PROPERTYGET | Styles | 493 | Metadata Only | Reviewed | Not Tested | |
| Worksheets | PROPERTYGET | Sheets | 494 | Metadata Only | Reviewed | Not Tested | |
| Author | PROPERTYGET/PROPERTYPUT | String | 574 | Metadata Only | Reviewed | Not Tested | |
| Comments | PROPERTYGET/PROPERTYPUT | String | 575 | Metadata Only | Reviewed | Not Tested | |
| Keywords | PROPERTYGET/PROPERTYPUT | String | 577 | Metadata Only | Reviewed | Not Tested | |
| Excel4MacroSheets | PROPERTYGET | Sheets | 579 | Metadata Only | Reviewed | Not Tested | |
| Excel4IntlMacroSheets | PROPERTYGET | Sheets | 581 | Metadata Only | Reviewed | Not Tested | |
| Modules | PROPERTYGET | Sheets | 582 | Metadata Only | Reviewed | Not Tested | |
| ProtectStructure | PROPERTYGET | bool | 588 | Metadata Only | Reviewed | Not Tested | |
| PivotTables | PROPERTYGET | Object | 690 | Metadata Only | Reviewed | Not Tested | |
| DialogSheets | PROPERTYGET | Sheets | 764 | Metadata Only | Reviewed | Not Tested | |
| UpdateLinks | PROPERTYGET/PROPERTYPUT | XlUpdateLinks | 864 | Metadata Only | Reviewed | Not Tested | |
| RoutingSlip | PROPERTYGET | RoutingSlip | 949 | Metadata Only | Reviewed | Not Tested | |
| HasRoutingSlip | PROPERTYGET/PROPERTYPUT | bool | 950 | Metadata Only | Reviewed | Not Tested | |
| Routed | PROPERTYGET | bool | 951 | Metadata Only | Reviewed | Not Tested | |
| Subject | PROPERTYGET/PROPERTYPUT | String | 953 | Metadata Only | Reviewed | Not Tested | |
| HasMailer | PROPERTYGET/PROPERTYPUT | bool | 976 | Metadata Only | Reviewed | Not Tested | |
| Mailer | PROPERTYGET | Mailer | 979 | Metadata Only | Reviewed | Not Tested | |
| OnSheetActivate | PROPERTYGET/PROPERTYPUT | String | 1031 | Metadata Only | Reviewed | Not Tested | |
| OnSheetDeactivate | PROPERTYGET/PROPERTYPUT | String | 1081 | Metadata Only | Reviewed | Not Tested | |
| WritePassword | PROPERTYGET/PROPERTYPUT | String | 1128 | Metadata Only | Reviewed | Not Tested | |
| MultiUserEditing | PROPERTYGET | bool | 1169 | Metadata Only | Reviewed | Not Tested | |
| ShowConflictHistory | PROPERTYGET/PROPERTYPUT | bool | 1171 | Metadata Only | Reviewed | Not Tested | |
| RevisionNumber | PROPERTYGET | i32 | 1172 | Metadata Only | Reviewed | Not Tested | |
| UserStatus | PROPERTYGET | AutomationValue | 1173 | Metadata Only | Reviewed | Not Tested | |
| ConflictResolution | PROPERTYGET/PROPERTYPUT | XlSaveConflictResolution | 1175 | Metadata Only | Reviewed | Not Tested | |
| BuiltinDocumentProperties | PROPERTYGET | Object | 1176 | Metadata Only | Reviewed | Not Tested | |
| CustomDocumentProperties | PROPERTYGET | Object | 1177 | Metadata Only | Reviewed | Not Tested | |
| OnSave | PROPERTYGET/PROPERTYPUT | String | 1178 | Metadata Only | Reviewed | Not Tested | |
| Container | PROPERTYGET | Object | 1190 | Metadata Only | Reviewed | Not Tested | |
| UserControl | PROPERTYGET/PROPERTYPUT | bool | 1210 | Metadata Only | Reviewed | Not Tested | |
| CodeName | PROPERTYGET | String | 1373 | Metadata Only | Reviewed | Not Tested | |
| CommandBars | PROPERTYGET | CommandBars | 1439 | Metadata Only | Reviewed | Not Tested | |
| AcceptLabelsInFormulas | PROPERTYGET/PROPERTYPUT | bool | 1441 | Metadata Only | Reviewed | Not Tested | |
| AutoUpdateFrequency | PROPERTYGET/PROPERTYPUT | i32 | 1442 | Metadata Only | Reviewed | Not Tested | |
| AutoUpdateSaveChanges | PROPERTYGET/PROPERTYPUT | bool | 1443 | Metadata Only | Reviewed | Not Tested | |
| ChangeHistoryDuration | PROPERTYGET/PROPERTYPUT | i32 | 1444 | Metadata Only | Reviewed | Not Tested | |
| IsAddin | PROPERTYGET/PROPERTYPUT | bool | 1445 | Metadata Only | Reviewed | Not Tested | |
| PersonalViewListSettings | PROPERTYGET/PROPERTYPUT | bool | 1447 | Metadata Only | Reviewed | Not Tested | |
| PersonalViewPrintSettings | PROPERTYGET/PROPERTYPUT | bool | 1448 | Metadata Only | Reviewed | Not Tested | |
| CustomViews | PROPERTYGET | CustomViews | 1456 | Metadata Only | Reviewed | Not Tested | |
| TemplateRemoveExtData | PROPERTYGET/PROPERTYPUT | bool | 1457 | Metadata Only | Reviewed | Not Tested | |
| HighlightChangesOnScreen | PROPERTYGET/PROPERTYPUT | bool | 1461 | Metadata Only | Reviewed | Not Tested | |
| KeepChangeHistory | PROPERTYGET/PROPERTYPUT | bool | 1462 | Metadata Only | Reviewed | Not Tested | |
| ListChangesOnNewSheet | PROPERTYGET/PROPERTYPUT | bool | 1463 | Metadata Only | Reviewed | Not Tested | |
| VBProject | PROPERTYGET | VBProject | 1469 | Metadata Only | Reviewed | Not Tested | |
| IsInplace | PROPERTYGET | bool | 1769 | Metadata Only | Reviewed | Not Tested | |
| CalculationVersion | PROPERTYGET | i32 | 1806 | Metadata Only | Reviewed | Not Tested | |
| PublishObjects | PROPERTYGET | PublishObjects | 1819 | Metadata Only | Reviewed | Not Tested | |
| WebOptions | PROPERTYGET | WebOptions | 1820 | Metadata Only | Reviewed | Not Tested | |
| HTMLProject | PROPERTYGET | HTMLProject | 1823 | Metadata Only | Reviewed | Not Tested | |
| EnvelopeVisible | PROPERTYGET/PROPERTYPUT | bool | 1824 | Metadata Only | Reviewed | Not Tested | |
| VBASigned | PROPERTYGET | bool | 1828 | Metadata Only | Reviewed | Not Tested | |
| FullNameURLEncoded | PROPERTYGET | String | 1927 | Metadata Only | Reviewed | Not Tested | |
| ReadOnlyRecommended | PROPERTYGET/PROPERTYPUT | bool | 2005 | Metadata Only | Reviewed | Not Tested | |
| ShowPivotTableFieldList | PROPERTYGET/PROPERTYPUT | bool | 2046 | Metadata Only | Reviewed | Not Tested | |
| EnableAutoRecover | PROPERTYGET/PROPERTYPUT | bool | 2049 | Metadata Only | Reviewed | Not Tested | |
| RemovePersonalInformation | PROPERTYGET/PROPERTYPUT | bool | 2050 | Metadata Only | Reviewed | Not Tested | |
| PasswordEncryptionProvider | PROPERTYGET | String | 2059 | Metadata Only | Reviewed | Not Tested | |
| PasswordEncryptionAlgorithm | PROPERTYGET | String | 2060 | Metadata Only | Reviewed | Not Tested | |
| PasswordEncryptionKeyLength | PROPERTYGET | i32 | 2061 | Metadata Only | Reviewed | Not Tested | |
| PasswordEncryptionFileProperties | PROPERTYGET | bool | 2063 | Metadata Only | Reviewed | Not Tested | |
| SmartTagOptions | PROPERTYGET | SmartTagOptions | 2064 | Metadata Only | Reviewed | Not Tested | |
| Permission | PROPERTYGET | Permission | 2264 | Metadata Only | Reviewed | Not Tested | |
| SharedWorkspace | PROPERTYGET | SharedWorkspace | 2265 | Metadata Only | Reviewed | Not Tested | |
| Sync | PROPERTYGET | Sync | 2266 | Metadata Only | Reviewed | Not Tested | |
| XmlNamespaces | PROPERTYGET | XmlNamespaces | 2268 | Metadata Only | Reviewed | Not Tested | |
| XmlMaps | PROPERTYGET | XmlMaps | 2269 | Metadata Only | Reviewed | Not Tested | |
| SmartDocument | PROPERTYGET | SmartDocument | 2273 | Metadata Only | Reviewed | Not Tested | |
| DocumentLibraryVersions | PROPERTYGET | DocumentLibraryVersions | 2274 | Metadata Only | Reviewed | Not Tested | |
| InactiveListBorderVisible | PROPERTYGET/PROPERTYPUT | bool | 2275 | Metadata Only | Reviewed | Not Tested | |
| DisplayInkComments | PROPERTYGET/PROPERTYPUT | bool | 2276 | Metadata Only | Reviewed | Not Tested | |
| ContentTypeProperties | PROPERTYGET | MetaProperties | 2512 | Metadata Only | Reviewed | Not Tested | |
| Connections | PROPERTYGET | Connections | 2513 | Metadata Only | Reviewed | Not Tested | |
| Signatures | PROPERTYGET | SignatureSet | 2516 | Metadata Only | Reviewed | Not Tested | |
| ServerPolicy | PROPERTYGET | ServerPolicy | 2519 | Metadata Only | Reviewed | Not Tested | |
| DocumentInspectors | PROPERTYGET | DocumentInspectors | 2521 | Metadata Only | Reviewed | Not Tested | |
| ServerViewableItems | PROPERTYGET | ServerViewableItems | 2524 | Metadata Only | Reviewed | Not Tested | |
| TableStyles | PROPERTYGET | TableStyles | 2525 | Metadata Only | Reviewed | Not Tested | |
| DefaultTableStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | 2526 | Metadata Only | Reviewed | Not Tested | |
| DefaultPivotTableStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | 2527 | Metadata Only | Reviewed | Not Tested | |
| CheckCompatibility | PROPERTYGET/PROPERTYPUT | bool | 2528 | Metadata Only | Reviewed | Not Tested | |
| HasVBProject | PROPERTYGET | bool | 2529 | Metadata Only | Reviewed | Not Tested | |
| CustomXMLParts | PROPERTYGET | CustomXMLParts | 2530 | Metadata Only | Reviewed | Not Tested | |
| Final | PROPERTYGET/PROPERTYPUT | bool | 2531 | Metadata Only | Reviewed | Not Tested | |
| Research | PROPERTYGET | Research | 2532 | Metadata Only | Reviewed | Not Tested | |
| Theme | PROPERTYGET | OfficeTheme | 2533 | Metadata Only | Reviewed | Not Tested | |
| Excel8CompatibilityMode | PROPERTYGET | bool | 2535 | Metadata Only | Reviewed | Not Tested | |
| ConnectionsDisabled | PROPERTYGET | bool | 2536 | Metadata Only | Reviewed | Not Tested | |
| ShowPivotChartActiveFields | PROPERTYGET/PROPERTYPUT | bool | 2538 | Metadata Only | Reviewed | Not Tested | |
| IconSets | PROPERTYGET | IconSets | 2539 | Metadata Only | Reviewed | Not Tested | |
| EncryptionProvider | PROPERTYGET/PROPERTYPUT | String | 2540 | Metadata Only | Reviewed | Not Tested | |
| DoNotPromptForConvert | PROPERTYGET/PROPERTYPUT | bool | 2541 | Metadata Only | Reviewed | Not Tested | |
| ForceFullCalculation | PROPERTYGET/PROPERTYPUT | bool | 2542 | Metadata Only | Reviewed | Not Tested | |
| SlicerCaches | PROPERTYGET | SlicerCaches | 2866 | Metadata Only | Reviewed | Not Tested | |
| ActiveSlicer | PROPERTYGET | Slicer | 2867 | Metadata Only | Reviewed | Not Tested | |
| DefaultSlicerStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | 2868 | Metadata Only | Reviewed | Not Tested | |
| AccuracyVersion | PROPERTYGET/PROPERTYPUT | i32 | 2871 | Metadata Only | Reviewed | Not Tested | |
| ChartDataPointTrack | PROPERTYGET/PROPERTYPUT | bool | 2998 | Metadata Only | Reviewed | Not Tested | |
| CaseSensitive | PROPERTYGET | bool | 3056 | Metadata Only | Reviewed | Not Tested | |
| UseWholeCellCriteria | PROPERTYGET | bool | 3057 | Metadata Only | Reviewed | Not Tested | |
| UseWildcards | PROPERTYGET | bool | 3058 | Metadata Only | Reviewed | Not Tested | |
| Model | PROPERTYGET | Model | 3059 | Metadata Only | Reviewed | Not Tested | |
| DefaultTimelineStyle | PROPERTYGET/PROPERTYPUT | AutomationValue | 3060 | Metadata Only | Reviewed | Not Tested | |
| WorkIdentity | PROPERTYGET/PROPERTYPUT | String | 3173 | Metadata Only | Reviewed | Not Tested | |
| Queries | PROPERTYGET | Queries | 3186 | Metadata Only | Reviewed | Not Tested | |
| AutoSaveOn | PROPERTYGET/PROPERTYPUT | bool | 3232 | Metadata Only | Reviewed | Not Tested | |
| SensitivityLabel | PROPERTYGET | ISensitivityLabel | 3379 | Metadata Only | Reviewed | Not Tested | |
| ExternalCodeServiceTimeout | PROPERTYGET/PROPERTYPUT | i32 | 3406 | Metadata Only | Reviewed | Not Tested | |
| CompatibilityVersion | PROPERTYGET/PROPERTYPUT | i32 | 3416 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---:|---|---|---|---|
| SaveCopyAs | Unknown | 1 | 175 | Metadata Only | Reviewed | Not Tested | |
| Close | Unknown | 3 | 277 | Implemented | Reviewed | Live Tested | |
| NewWindow | Window | 0 | 280 | Metadata Only | Reviewed | Not Tested | |
| PrintPreview | Unknown | 1 | 281 | Metadata Only | Reviewed | Not Tested | |
| _Protect | Unknown | 3 | 282 | Metadata Only | Reviewed | Not Tested | |
| Save | Unknown | 0 | 283 | Metadata Only | Reviewed | Not Tested | |
| __SaveAs | Unknown | 11 | 284 | Metadata Only | Reviewed | Not Tested | |
| Unprotect | Unknown | 1 | 285 | Metadata Only | Reviewed | Not Tested | |
| Activate | Unknown | 0 | 304 | Metadata Only | Reviewed | Not Tested | |
| DeleteNumberFormat | Unknown | 1 | 397 | Metadata Only | Reviewed | Not Tested | |
| RunAutoMacros | Unknown | 1 | 634 | Metadata Only | Reviewed | Not Tested | |
| PivotTableWizard | Unknown | 16 | 684 | Metadata Only | Reviewed | Not Tested | |
| ChangeLink | Unknown | 3 | 802 | Metadata Only | Reviewed | Not Tested | |
| OpenLinks | Unknown | 3 | 803 | Metadata Only | Reviewed | Not Tested | |
| UpdateLink | Unknown | 2 | 804 | Metadata Only | Reviewed | Not Tested | |
| LinkInfo | AutomationValue | 4 | 807 | Metadata Only | Reviewed | Not Tested | |
| LinkSources | AutomationValue | 1 | 808 | Metadata Only | Reviewed | Not Tested | |
| SetLinkOnData | Unknown | 2 | 809 | Metadata Only | Reviewed | Not Tested | |
| __PrintOut | Unknown | 7 | 905 | Metadata Only | Reviewed | Not Tested | |
| Route | Unknown | 0 | 946 | Metadata Only | Reviewed | Not Tested | |
| SendMail | Unknown | 3 | 947 | Metadata Only | Reviewed | Not Tested | |
| ForwardMailer | Unknown | 0 | 973 | Metadata Only | Reviewed | Not Tested | |
| Reply | Unknown | 0 | 977 | Metadata Only | Reviewed | Not Tested | |
| ReplyAll | Unknown | 0 | 978 | Metadata Only | Reviewed | Not Tested | |
| SendMailer | Unknown | 2 | 980 | Metadata Only | Reviewed | Not Tested | |
| ChangeFileAccess | Unknown | 3 | 989 | Metadata Only | Reviewed | Not Tested | |
| UpdateFromFile | Unknown | 0 | 995 | Metadata Only | Reviewed | Not Tested | |
| Post | Unknown | 1 | 1166 | Metadata Only | Reviewed | Not Tested | |
| ExclusiveAccess | bool | 0 | 1168 | Metadata Only | Reviewed | Not Tested | |
| MergeWorkbook | Unknown | 1 | 1446 | Metadata Only | Reviewed | Not Tested | |
| PivotCaches | PivotCaches | 0 | 1449 | Metadata Only | Reviewed | Not Tested | |
| _ProtectSharing | Unknown | 6 | 1450 | Metadata Only | Reviewed | Not Tested | |
| RefreshAll | Unknown | 0 | 1452 | Metadata Only | Reviewed | Not Tested | |
| RemoveUser | Unknown | 1 | 1453 | Metadata Only | Reviewed | Not Tested | |
| UnprotectSharing | Unknown | 1 | 1455 | Metadata Only | Reviewed | Not Tested | |
| HighlightChangesOptions | Unknown | 3 | 1458 | Metadata Only | Reviewed | Not Tested | |
| PurgeChangeHistoryNow | Unknown | 2 | 1464 | Metadata Only | Reviewed | Not Tested | |
| AcceptAllChanges | Unknown | 3 | 1466 | Metadata Only | Reviewed | Not Tested | |
| RejectAllChanges | Unknown | 3 | 1467 | Metadata Only | Reviewed | Not Tested | |
| ResetColors | Unknown | 0 | 1468 | Metadata Only | Reviewed | Not Tested | |
| FollowHyperlink | Unknown | 7 | 1470 | Metadata Only | Reviewed | Not Tested | |
| AddToFavorites | Unknown | 0 | 1476 | Metadata Only | Reviewed | Not Tested | |
| _PrintOut | Unknown | 8 | 1772 | Metadata Only | Reviewed | Not Tested | |
| WebPagePreview | Unknown | 0 | 1818 | Metadata Only | Reviewed | Not Tested | |
| ReloadAs | Unknown | 1 | 1821 | Metadata Only | Reviewed | Not Tested | |
| sblt | Unknown | 1 | 1826 | Metadata Only | Reviewed | Not Tested | |
| _SaveAs | Unknown | 12 | 1925 | Metadata Only | Reviewed | Not Tested | |
| Protect | Unknown | 3 | 2029 | Metadata Only | Reviewed | Not Tested | |
| Dummy17 | Unknown | 1 | 2044 | Metadata Only | Reviewed | Not Tested | |
| BreakLink | Unknown | 2 | 2047 | Metadata Only | Reviewed | Not Tested | |
| Dummy16 | Unknown | 0 | 2048 | Metadata Only | Reviewed | Not Tested | |
| CheckIn | Unknown | 3 | 2051 | Metadata Only | Reviewed | Not Tested | |
| CanCheckIn | bool | 0 | 2053 | Metadata Only | Reviewed | Not Tested | |
| SendForReview | Unknown | 4 | 2054 | Metadata Only | Reviewed | Not Tested | |
| ReplyWithChanges | Unknown | 1 | 2057 | Metadata Only | Reviewed | Not Tested | |
| EndReview | Unknown | 0 | 2058 | Metadata Only | Reviewed | Not Tested | |
| SetPasswordEncryptionOptions | Unknown | 4 | 2062 | Metadata Only | Reviewed | Not Tested | |
| RecheckSmartTags | Unknown | 0 | 2065 | Metadata Only | Reviewed | Not Tested | |
| SendFaxOverInternet | Unknown | 3 | 2267 | Metadata Only | Reviewed | Not Tested | |
| XmlImport | XlXmlImportResult | 4 | 2270 | Metadata Only | Reviewed | Not Tested | |
| XmlImportXml | XlXmlImportResult | 4 | 2277 | Metadata Only | Reviewed | Not Tested | |
| SaveAsXMLData | Unknown | 2 | 2278 | Metadata Only | Reviewed | Not Tested | |
| ToggleFormsDesign | Unknown | 0 | 2279 | Metadata Only | Reviewed | Not Tested | |
| PrintOut | Unknown | 9 | 2361 | Metadata Only | Reviewed | Not Tested | |
| _ExportAsFixedFormat | Unknown | 9 | 2493 | Metadata Only | Reviewed | Not Tested | |
| RemoveDocumentInformation | Unknown | 1 | 2514 | Metadata Only | Reviewed | Not Tested | |
| CheckInWithVersion | Unknown | 4 | 2517 | Metadata Only | Reviewed | Not Tested | |
| LockServerFile | Unknown | 0 | 2520 | Metadata Only | Reviewed | Not Tested | |
| GetWorkflowTasks | WorkflowTasks | 0 | 2522 | Metadata Only | Reviewed | Not Tested | |
| GetWorkflowTemplates | WorkflowTemplates | 0 | 2523 | Metadata Only | Reviewed | Not Tested | |
| ApplyTheme | Unknown | 1 | 2534 | Metadata Only | Reviewed | Not Tested | |
| EnableConnections | Unknown | 0 | 2537 | Metadata Only | Reviewed | Not Tested | |
| ProtectSharing | Unknown | 7 | 2543 | Metadata Only | Reviewed | Not Tested | |
| Dummy26 | Unknown | 0 | 2869 | Metadata Only | Reviewed | Not Tested | |
| Dummy27 | Unknown | 0 | 2870 | Metadata Only | Reviewed | Not Tested | |
| CreateForecastSheet | Unknown | 10 | 3167 | Metadata Only | Reviewed | Not Tested | |
| SaveAs | Unknown | 13 | 3174 | Metadata Only | Reviewed | Not Tested | |
| ExportAsFixedFormat | Unknown | 10 | 3175 | Metadata Only | Reviewed | Not Tested | |
| PublishToPBI | String | 3 | 3257 | Metadata Only | Reviewed | Not Tested | |
| ConvertComments | Unknown | 0 | 3279 | Metadata Only | Reviewed | Not Tested | |
| PublishToDocs | String | 3 | 3334 | Metadata Only | Reviewed | Not Tested | |
| LookUpInDocs | PublishedDocs | 1 | 3335 | Metadata Only | Reviewed | Not Tested | |
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
