# PageSetup

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `PageSetup` |
| GUID | `{000208b4-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::PageSetup` |
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
| Orientation | PROPERTYGET/PROPERTYPUT | XlPageOrientation | declared | 134 | Implemented | Reviewed | Blocked | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| Order | PROPERTYGET/PROPERTYPUT | XlOrder | declared | 192 | Implemented | Reviewed | Blocked | |
| Zoom | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 663 | Implemented | Reviewed | Blocked | |
| LeftMargin | PROPERTYGET/PROPERTYPUT | f64 | declared | 999 | Implemented | Reviewed | Blocked | |
| RightMargin | PROPERTYGET/PROPERTYPUT | f64 | declared | 1000 | Implemented | Reviewed | Blocked | |
| TopMargin | PROPERTYGET/PROPERTYPUT | f64 | declared | 1001 | Implemented | Reviewed | Blocked | |
| BottomMargin | PROPERTYGET/PROPERTYPUT | f64 | declared | 1002 | Implemented | Reviewed | Blocked | |
| PrintHeadings | PROPERTYGET/PROPERTYPUT | bool | declared | 1003 | Implemented | Reviewed | Blocked | |
| PrintGridlines | PROPERTYGET/PROPERTYPUT | bool | declared | 1004 | Implemented | Reviewed | Blocked | |
| CenterHorizontally | PROPERTYGET/PROPERTYPUT | bool | declared | 1005 | Implemented | Reviewed | Blocked | |
| CenterVertically | PROPERTYGET/PROPERTYPUT | bool | declared | 1006 | Implemented | Reviewed | Blocked | |
| PaperSize | PROPERTYGET/PROPERTYPUT | XlPaperSize | declared | 1007 | Implemented | Reviewed | Blocked | |
| FirstPageNumber | PROPERTYGET/PROPERTYPUT | i32 | declared | 1008 | Implemented | Reviewed | Blocked | |
| BlackAndWhite | PROPERTYGET/PROPERTYPUT | bool | declared | 1009 | Implemented | Reviewed | Blocked | |
| CenterFooter | PROPERTYGET/PROPERTYPUT | String | declared | 1010 | Implemented | Reviewed | Blocked | |
| CenterHeader | PROPERTYGET/PROPERTYPUT | String | declared | 1011 | Implemented | Reviewed | Blocked | |
| ChartSize | PROPERTYGET/PROPERTYPUT | XlObjectSize | declared | 1012 | Metadata Only | Reviewed | Not Tested | |
| FitToPagesTall | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1013 | Implemented | Reviewed | Blocked | |
| FitToPagesWide | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1014 | Implemented | Reviewed | Blocked | |
| FooterMargin | PROPERTYGET/PROPERTYPUT | f64 | declared | 1015 | Implemented | Reviewed | Blocked | |
| HeaderMargin | PROPERTYGET/PROPERTYPUT | f64 | declared | 1016 | Implemented | Reviewed | Blocked | |
| LeftFooter | PROPERTYGET/PROPERTYPUT | String | declared | 1017 | Implemented | Reviewed | Blocked | |
| LeftHeader | PROPERTYGET/PROPERTYPUT | String | declared | 1018 | Implemented | Reviewed | Blocked | |
| PrintArea | PROPERTYGET/PROPERTYPUT | String | declared | 1019 | Implemented | Reviewed | Blocked | |
| Draft | PROPERTYGET/PROPERTYPUT | bool | declared | 1020 | Implemented | Reviewed | Blocked | |
| PrintNotes | PROPERTYGET/PROPERTYPUT | bool | declared | 1021 | Metadata Only | Reviewed | Not Tested | |
| PrintQuality | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 1022 | Implemented | Reviewed | Blocked | |
| PrintTitleColumns | PROPERTYGET/PROPERTYPUT | String | declared | 1023 | Implemented | Reviewed | Blocked | |
| PrintTitleRows | PROPERTYGET/PROPERTYPUT | String | declared | 1024 | Implemented | Reviewed | Blocked | |
| RightFooter | PROPERTYGET/PROPERTYPUT | String | declared | 1025 | Implemented | Reviewed | Blocked | |
| RightHeader | PROPERTYGET/PROPERTYPUT | String | declared | 1026 | Implemented | Reviewed | Blocked | |
| PrintComments | PROPERTYGET/PROPERTYPUT | XlPrintLocation | declared | 1524 | Implemented | Reviewed | Blocked | |
| PrintErrors | PROPERTYGET/PROPERTYPUT | XlPrintErrors | declared | 2149 | Implemented | Reviewed | Blocked | |
| CenterHeaderPicture | PROPERTYGET | Graphic | declared | 2150 | Metadata Only | Reviewed | Not Tested | |
| CenterFooterPicture | PROPERTYGET | Graphic | declared | 2151 | Metadata Only | Reviewed | Not Tested | |
| LeftHeaderPicture | PROPERTYGET | Graphic | declared | 2152 | Metadata Only | Reviewed | Not Tested | |
| LeftFooterPicture | PROPERTYGET | Graphic | declared | 2153 | Metadata Only | Reviewed | Not Tested | |
| RightHeaderPicture | PROPERTYGET | Graphic | declared | 2154 | Metadata Only | Reviewed | Not Tested | |
| RightFooterPicture | PROPERTYGET | Graphic | declared | 2155 | Metadata Only | Reviewed | Not Tested | |
| OddAndEvenPagesHeaderFooter | PROPERTYGET/PROPERTYPUT | bool | declared | 2600 | Metadata Only | Reviewed | Not Tested | |
| DifferentFirstPageHeaderFooter | PROPERTYGET/PROPERTYPUT | bool | declared | 2601 | Metadata Only | Reviewed | Not Tested | |
| ScaleWithDocHeaderFooter | PROPERTYGET/PROPERTYPUT | bool | declared | 2602 | Metadata Only | Reviewed | Not Tested | |
| AlignMarginsHeaderFooter | PROPERTYGET/PROPERTYPUT | bool | declared | 2603 | Metadata Only | Reviewed | Not Tested | |
| Pages | PROPERTYGET | Pages | declared | 2604 | Metadata Only | Reviewed | Not Tested | |
| EvenPage | PROPERTYGET | Page | declared | 2605 | Metadata Only | Reviewed | Not Tested | |
| FirstPage | PROPERTYGET | Page | declared | 2606 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
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
