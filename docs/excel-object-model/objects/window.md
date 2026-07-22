# Window

## Summary

This type-library object is structurally inventoried for future wrapper planning.

## Sources

- registered Excel type library
- official Microsoft documentation URL recorded in metadata
<!-- BEGIN GENERATED MEMBERS -->
## Identity

| Field | Value |
|---|---|
| Interface | `Window` |
| GUID | `{00020893-0000-0000-c000-000000000046}` |
| Object kind | dispatch-interface |
| Surface class | primary-object-model |
| Roadmap class | implemented-wrapper |
| Type flags | 4096 |
| Crate type | `excel_com::Window` |
| Implementation | Partial |
| Documentation | Reviewed |
| Tests | Blocked |

## Capabilities

No capability metadata is recorded for this surface.


## Relationships

| Relationship | Target | Status |
|---|---|---|
| `ActiveCell` | `excel.range` | Metadata Only |
| `Application` | `excel.application` | Metadata Only |
| `RangeSelection` | `excel.range` | Metadata Only |
| `VisibleRange` | `excel.range` | Metadata Only |

## Properties

| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---|---|---:|---|---|---|---|
| Type | PROPERTYGET | XlWindowType | declared | 108 | Metadata Only | Reviewed | Not Tested | |
| Width | PROPERTYGET/PROPERTYPUT | f64 | declared | 122 | Metadata Only | Reviewed | Not Tested | |
| Height | PROPERTYGET/PROPERTYPUT | f64 | declared | 123 | Metadata Only | Reviewed | Not Tested | |
| Top | PROPERTYGET/PROPERTYPUT | f64 | declared | 126 | Metadata Only | Reviewed | Not Tested | |
| Left | PROPERTYGET/PROPERTYPUT | f64 | declared | 127 | Metadata Only | Reviewed | Not Tested | |
| Caption | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 139 | Metadata Only | Reviewed | Not Tested | |
| Selection | PROPERTYGET | Object | declared | 147 | Metadata Only | Reviewed | Not Tested | |
| Application | PROPERTYGET | Application | declared | 148 | Metadata Only | Reviewed | Not Tested | |
| Creator | PROPERTYGET | XlCreator | declared | 149 | Metadata Only | Reviewed | Not Tested | |
| Parent | PROPERTYGET | Object | declared | 150 | Metadata Only | Reviewed | Not Tested | |
| ActiveChart | PROPERTYGET | Chart | declared | 183 | Metadata Only | Reviewed | Not Tested | |
| ActiveCell | PROPERTYGET | Range | declared | 305 | Metadata Only | Reviewed | Not Tested | |
| ActiveSheet | PROPERTYGET | Object | declared | 307 | Metadata Only | Reviewed | Not Tested | |
| UsableHeight | PROPERTYGET | f64 | declared | 389 | Metadata Only | Reviewed | Not Tested | |
| UsableWidth | PROPERTYGET | f64 | declared | 390 | Metadata Only | Reviewed | Not Tested | |
| WindowState | PROPERTYGET/PROPERTYPUT | XlWindowState | declared | 396 | Metadata Only | Reviewed | Not Tested | |
| Index | PROPERTYGET | i32 | declared | 486 | Implemented | Reviewed | Blocked | |
| Visible | PROPERTYGET/PROPERTYPUT | bool | declared | 558 | Metadata Only | Reviewed | Not Tested | |
| OnWindow | PROPERTYGET/PROPERTYPUT | String | declared | 623 | Metadata Only | Reviewed | Not Tested | |
| ActivePane | PROPERTYGET | Pane | declared | 642 | Metadata Only | Reviewed | Not Tested | |
| DisplayFormulas | PROPERTYGET/PROPERTYPUT | bool | declared | 644 | Metadata Only | Reviewed | Not Tested | |
| DisplayGridlines | PROPERTYGET/PROPERTYPUT | bool | declared | 645 | Implemented | Reviewed | Blocked | |
| DisplayHeadings | PROPERTYGET/PROPERTYPUT | bool | declared | 646 | Implemented | Reviewed | Blocked | |
| DisplayOutline | PROPERTYGET/PROPERTYPUT | bool | declared | 647 | Metadata Only | Reviewed | Not Tested | |
| _DisplayRightToLeft | PROPERTYGET/PROPERTYPUT | bool | declared | 648 | Metadata Only | Reviewed | Not Tested | |
| DisplayZeros | PROPERTYGET/PROPERTYPUT | bool | declared | 649 | Implemented | Reviewed | Blocked | |
| FreezePanes | PROPERTYGET/PROPERTYPUT | bool | declared | 650 | Implemented | Reviewed | Blocked | |
| GridlineColor | PROPERTYGET/PROPERTYPUT | i32 | declared | 651 | Metadata Only | Reviewed | Not Tested | |
| GridlineColorIndex | PROPERTYGET/PROPERTYPUT | XlColorIndex | declared | 652 | Metadata Only | Reviewed | Not Tested | |
| Panes | PROPERTYGET | Panes | declared | 653 | Metadata Only | Reviewed | Not Tested | |
| ScrollColumn | PROPERTYGET/PROPERTYPUT | i32 | declared | 654 | Implemented | Reviewed | Blocked | |
| ScrollRow | PROPERTYGET/PROPERTYPUT | i32 | declared | 655 | Implemented | Reviewed | Blocked | |
| SelectedSheets | PROPERTYGET | Sheets | declared | 656 | Metadata Only | Reviewed | Not Tested | |
| Split | PROPERTYGET/PROPERTYPUT | bool | declared | 657 | Metadata Only | Reviewed | Not Tested | |
| SplitColumn | PROPERTYGET/PROPERTYPUT | i32 | declared | 658 | Implemented | Reviewed | Blocked | |
| SplitHorizontal | PROPERTYGET/PROPERTYPUT | f64 | declared | 659 | Implemented | Reviewed | Blocked | |
| SplitRow | PROPERTYGET/PROPERTYPUT | i32 | declared | 660 | Implemented | Reviewed | Blocked | |
| SplitVertical | PROPERTYGET/PROPERTYPUT | f64 | declared | 661 | Implemented | Reviewed | Blocked | |
| Zoom | PROPERTYGET/PROPERTYPUT | AutomationValue | declared | 663 | Implemented | Reviewed | Blocked | |
| TabRatio | PROPERTYGET/PROPERTYPUT | f64 | declared | 673 | Metadata Only | Reviewed | Not Tested | |
| DisplayHorizontalScrollBar | PROPERTYGET/PROPERTYPUT | bool | declared | 921 | Metadata Only | Reviewed | Not Tested | |
| DisplayVerticalScrollBar | PROPERTYGET/PROPERTYPUT | bool | declared | 922 | Metadata Only | Reviewed | Not Tested | |
| DisplayWorkbookTabs | PROPERTYGET/PROPERTYPUT | bool | declared | 923 | Metadata Only | Reviewed | Not Tested | |
| VisibleRange | PROPERTYGET | Range | declared | 1118 | Metadata Only | Reviewed | Not Tested | |
| WindowNumber | PROPERTYGET | i32 | declared | 1119 | Metadata Only | Reviewed | Not Tested | |
| RangeSelection | PROPERTYGET | Range | declared | 1189 | Metadata Only | Reviewed | Not Tested | |
| EnableResize | PROPERTYGET/PROPERTYPUT | bool | declared | 1192 | Metadata Only | Reviewed | Not Tested | |
| View | PROPERTYGET/PROPERTYPUT | XlWindowView | declared | 1194 | Implemented | Reviewed | Blocked | |
| DisplayRightToLeft | PROPERTYGET/PROPERTYPUT | bool | declared | 1774 | Metadata Only | Reviewed | Not Tested | |
| Hwnd | PROPERTYGET | i32 | declared | 1950 | Metadata Only | Reviewed | Not Tested | |
| SheetViews | PROPERTYGET | SheetViews | declared | 2368 | Metadata Only | Reviewed | Not Tested | |
| ActiveSheetView | PROPERTYGET | Object | declared | 2369 | Metadata Only | Reviewed | Not Tested | |
| DisplayRuler | PROPERTYGET/PROPERTYPUT | bool | declared | 2370 | Metadata Only | Reviewed | Not Tested | |
| AutoFilterDateGrouping | PROPERTYGET/PROPERTYPUT | bool | declared | 2371 | Metadata Only | Reviewed | Not Tested | |
| DisplayWhitespace | PROPERTYGET/PROPERTYPUT | bool | declared | 2372 | Metadata Only | Reviewed | Not Tested | |
| DisplayDataTypeIcons | PROPERTYGET/PROPERTYPUT | bool | declared | 3418 | Metadata Only | Reviewed | Not Tested | |

## Methods

| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |
|---|---|---:|---|---:|---|---|---|---|
| Close | bool | 3 | declared | 277 | Metadata Only | Reviewed | Not Tested | |
| NewWindow | Window | 0 | declared | 280 | Metadata Only | Reviewed | Not Tested | |
| PrintPreview | AutomationValue | 1 | declared | 281 | Metadata Only | Reviewed | Not Tested | |
| Activate | AutomationValue | 0 | declared | 304 | Implemented | Reviewed | Blocked | |
| LargeScroll | AutomationValue | 4 | declared | 547 | Metadata Only | Reviewed | Not Tested | |
| SmallScroll | AutomationValue | 4 | declared | 548 | Metadata Only | Reviewed | Not Tested | |
| ScrollWorkbookTabs | AutomationValue | 2 | declared | 662 | Metadata Only | Reviewed | Not Tested | |
| ActivateNext | AutomationValue | 0 | declared | 1115 | Metadata Only | Reviewed | Not Tested | |
| ActivatePrevious | AutomationValue | 0 | declared | 1116 | Metadata Only | Reviewed | Not Tested | |
| _PrintOut | AutomationValue | 8 | declared | 1772 | Metadata Only | Reviewed | Not Tested | |
| PointsToScreenPixelsX | i32 | 1 | declared | 1776 | Metadata Only | Reviewed | Not Tested | |
| PointsToScreenPixelsY | i32 | 1 | declared | 1777 | Metadata Only | Reviewed | Not Tested | |
| RangeFromPoint | Object | 2 | declared | 1778 | Metadata Only | Reviewed | Not Tested | |
| ScrollIntoView | Unknown | 5 | declared | 1781 | Metadata Only | Reviewed | Not Tested | |
| PrintOut | AutomationValue | 8 | declared | 2361 | Metadata Only | Reviewed | Not Tested | |
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
