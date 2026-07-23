# Charts, shapes, pictures, and sparklines

## 1. Scope and baseline

This change adds the experimental, apartment-bound Excel drawing surface. It
starts from `5bb683a8e7d7adbdd358b4d9bf71ae9b8b8412b6`. A single fresh-process
`Workbooks.Add` baseline succeeded and Excel exited normally.

## 2. Prompt 16 runtime status

The Prompt 16 baseline was retained as historical evidence. Its unrelated live
test failures are not used as evidence for drawing members.

## 3. Drawing module architecture

`excel-com/src/excel/drawing/mod.rs` is the dedicated drawing subsystem. It
uses typed Excel and Office-role wrappers over the existing private dispatch
layer; it exposes neither raw COM pointers nor a generic Office dispatcher.

## 4. ChartObjects

`Worksheet::chart_objects` supplies one-based lookup, fallible iteration, and
point-based `ChartObjects.Add` with finite, positive dimensions.

## 5. Chart creation

`Worksheet::add_chart` creates the object, sets its Excel source data and
chart type, then applies optional orientation, title, and legend. It attempts
to delete a partially-created object on a later failure.

## 6. Chart sheets

`Workbook::charts`, `Charts`, and `ChartSheet` model the native chart-sheet
collection, including explicit before/after/end creation destinations.

## 7. Heterogeneous Sheet integration

The existing `Sheet` sum type now distinguishes `Sheet::Chart(ChartSheet)`
from worksheets and other sheet types using Excel's sheet-type value.

## 8. SeriesCollection and Series

Series collections provide count, one-based access, iteration, new-series and
Range-backed add operations. Series wrappers retain their Excel ownership.

## 9. Series values and formulas

`SeriesData` accepts only a Range, a rectangular `AutomationArray`, or an
Excel formula string. Values and XValues remain Excel `VARIANT` values on read
and are converted through the crate's existing Automation conversion layer.

## 10. Axes

`Axes::item` selects a primary or secondary axis and represents an absent
dispatch result as `None`. Axis scale, auto-scale, logarithmic base, tick and
number-format setters validate finite Rust numbers before COM.

## 11. Titles, legend and areas

Chart and axis titles, legend position/layout, chart area, and plot area are
typed wrappers. Text rejects embedded NUL through the existing BSTR helper.

## 12. Data labels

`Series::apply_data_labels` preserves all declared optional Excel positions.
Data-label collection access is intentionally bounded to count and one-based
item access.

## 13. Trendlines

Trendline collection access, iteration, add, equation, R-squared, type, and
deletion are modeled. Excel validates which trendline kinds apply to a series.

## 14. Error bars

Fixed, percentage, standard-deviation, standard-error, and custom source
arguments pass through `Series::set_error_bars`; no statistics are calculated
in Rust.

## 15. Chart formatting

Chart, legend, plot-area, and series formatting return typed `ChartFormat`,
`FillFormat`, and `LineFormat` wrappers. Their broader Office shared-format
operations remain version-sensitive and are not guessed.

## 16. Shapes

`Worksheet::shapes` and `Chart::shapes` provide one-based lookup, iteration,
AutoShape and line creation. Geometry is in Excel points and is finite and
positive where a width and height are required.

## 17. Pictures

`Shapes::add_picture` accepts an existing, reviewed local path only. The
caller chooses linking and saving-with-document; no image is downloaded or
followed from an external location.

## 18. Text boxes and text frames

Text boxes use typed Office `TextFrame2` and `TextRange2` role wrappers. Their
text is explicitly separate from Excel's `Font` interface.

## 19. Positioning and placement

Shape and chart-object coordinates use points. `Range::shape_bounds` derives
anchoring geometry from Excel's Range properties without a Rust layout engine.

## 20. Grouping and z-order

Z-order is exposed. Grouping was structurally investigated through
`Shapes.Range` name SAFEARRAY encoding but is deliberately reported unavailable
in capability metadata until its runtime ownership semantics are verified.

## 21. Chart export

`Chart::export` returns Excel's Boolean result. Export filter availability,
file existence, and image bytes are host observations rather than assumptions.

## 22. Range picture export

`Range::copy_picture` is implemented. The temporary-chart Range-to-image
helper is deferred because its cleanup path has not been independently proven.

## 23. Clipboard state

The API exposes only Excel's `CutCopyMode` and clearing operation. It neither
reads nor persists the operating-system clipboard contents.

## 24. Sparklines

`Range::sparkline_groups` exposes group lookup, iteration, add, type and
source-data operations. Location assignment follows the inspected typelib
contract; feature-level outcomes remain host-dependent.

## 25. Persistence

Persistence is a live-only claim. The corresponding evidence record is updated
only after an observed save, reopen, and object inspection sequence.

## 26. Runtime observations

The fresh `Workbooks.Add` baseline succeeded once. Prompt 17 test observations
are recorded separately in normalized JSON Lines records and distinguish
observed results from structural implementation.

## 27. Version-dependent behaviour

Installed chart-export filters, Office shared-format members, TextFrame2
availability, grouping ownership, and sparkline point-display members may vary
by Excel/Office version.

## 28. Explicit non-decisions

This work does not implement ActiveX, Form controls, OLE objects, embedded
files, SmartArt, 3-D models, Office scripts, VBA editing, PivotCharts, Power
View, events, external downloads, or a generic drawing DSL.

## 29. Rustdoc

The crate-level documentation describes the chart, shape, picture, clipboard,
and sparkline boundaries and includes opt-in live-test commands.

## 30. Inventory

The inventory records typed drawing objects, declared-member mappings,
collection element/index/iteration metadata, and a drawing capability group.
Grouping and Range image export are expressly false.

## 31. Validation

Formatting, workspace compilation, test, lint, Rustdoc, inventory, knowledge
base, and diff checks are run before the change is submitted. Their final
outcomes are reported with the pull request.

## 32. Remaining blockers

Live drawing claims require each actual Excel call. The unimplemented Range
image-export helper and unverified grouping/Office format operations remain
explicitly deferred rather than simulated.

## 33. Recommended Prompt 18 scope

Prompt 18 should first turn the remaining live observations into targeted
runtime matrices, then consider a proven Range-to-image helper, shared Office
format operations, and grouping only after deterministic ownership evidence.
