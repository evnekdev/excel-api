# Prompt 12: Core Range Formatting

## 1. Scope and baseline

Started from `888decf56132a480ba5c7fb035ecf8d7e9f617e3` (`Implement Excel references names and evaluation`). This bounded slice adds Range Font, Interior, Borders, number formats, alignment, wrapping, dimensions, and AutoFit. It excludes conditional formatting, styles, themes, tables, charts, shapes, and printing.

## 2. Formatting object graph

`Range::font`, `Range::interior`, and `Range::borders` return apartment-bound wrappers. `Borders::item` accepts `BorderIndex`, then returns `Border`. No raw COM pointer or generic formatting object is public.

## 3. Mixed-value semantics

`MixedValue<T>` distinguishes `Uniform(T)`, `Mixed`, and `Empty`. Excel `VT_NULL` is mapped to `Mixed` for the selected formatting getters; `VT_EMPTY` maps to `Empty`. Setters accept concrete values only, so callers cannot assign a synthetic mixed state.

## 4. Excel color representation

`ExcelColor` is a signed transparent raw integer to retain Excel sentinel values. Live RGB tests confirmed OLE COLORREF packing: red is the low byte, then green, then blue. Written as hexadecimal this is `0x00BBGGRR`, which is sometimes loosely called BGR; it does **not** mean the API accepts `(blue, green, red)`. `ExcelColor::from_rgb(r, g, b)` is exactly `r | (g << 8) | (b << 16)`. `RGB(12,34,56)` round-tripped as `3678732` (`0x0038220c`). This property is a 24-bit opaque color: the normal high byte is zero and it has no alpha/opacity channel. Theme and tint behavior are separate and intentionally excluded from this slice.

## 5. Color indices

`ExcelColorIndex` preserves raw signed values and exposes `AUTOMATIC = -4105` and `NONE = -4142` from `XlColorIndex`. Indexed colors remain workbook-palette and version dependent.

## 6. Font

`Font` supports name, size, bold, italic, underline, strikethrough, color, and color index. Font name rejects an embedded NUL and size rejects non-finite input before COM. The live workbook applied and read back Arial, 12 points, bold, italic, single underline, no strikeout, and the non-symmetric RGB color.

## 7. Interior

`Interior` supports color, color index, pattern, pattern color, and pattern color index. The live workbook applied solid fill, RGB `(240,200,120)`, and pattern color RGB `(11,22,33)`.

## 8. Borders collection

`Borders` is documented as an enum-keyed collection: `Count`, `Item(BorderIndex)`, and a fallible `_NewEnum` iterator. The observed fresh Range reported six enumerated Borders. Its order is a runtime observation, not a portable ordering guarantee.

## 9. Border

`Border` supports color, color index, line style, and weight. The four outer edges accepted continuous/thin/color formatting and read back as continuous. Assigning `BorderLineStyle::NONE` removed the bottom border and read back as that line-style value.

## 10. NumberFormat

The crate deliberately exposes invariant `NumberFormat`, not `NumberFormatLocal`. The live results were `General`, `0.00`, `0%`, `yyyy-mm-dd`, and `[Green]0.00;[Red]-0.00;0.00;@`; an embedded NUL is rejected before COM.

## 11. Alignment

`HorizontalAlignment` and `VerticalAlignment` are forward-compatible raw `i32` types with curated typelib constants. Center read back as `-4108`; differing horizontal alignments returned mixed state.

## 12. WrapText

`Range::set_wrap_text(true)` read back true on a uniform range. A true/false selection returned `Mixed`, not false.

## 13. RowHeight

`RowHeight` uses points. Distinct 20- and 30-point rows returned `Mixed`; a uniform 22-point setting read back as 22. Non-finite and negative inputs fail before COM. On this runtime, setting an entire row height to zero made it hidden.

## 14. ColumnWidth

`ColumnWidth` is an Excel character-width unit based on the Normal style font, not pixels. Distinct 14 and 18 widths returned `Mixed`; a uniform 16 read back as 16. On this runtime, setting an entire column width to zero made it hidden.

## 15. AutoFit

`Range::auto_fit` invokes Excel without changing the Range. `A1:C4.AutoFit` returned Excel's structured invocation error. `EntireColumn.AutoFit` and `EntireRow.AutoFit` succeeded; each resulting tested dimension was finite and positive, while column widths could differ.

## 16. Physical VARIANT observations

Uniform Font Bold and WrapText were `VT_BOOL`; Font Name and NumberFormat were `VT_BSTR`; Font Size, Font Color, Interior Color, RowHeight, and ColumnWidth were `VT_R8`; HorizontalAlignment and an individual Border LineStyle were `VT_I4`. Mixed Font fields, NumberFormat, HorizontalAlignment, WrapText, RowHeight, and ColumnWidth were observed as `VT_NULL`. These tags stay private implementation evidence.

## 17. Live results

The ignored `range_formatting_live` test starts only with zero Excel processes, makes the crate-created Excel instance visible, uses a fresh unsaved workbook, and naturally exits after close-without-saving plus `Application::quit`. The required existing four live tests and the new formatting test all pass with zero remaining Excel processes.

## 18. Explicit non-decisions

This change does not add conditional formatting, named styles, theme manipulation, gradients, tables, charts, shapes, printing, hidden-row/column APIs, or a style builder. Automatic colors are recorded as observations rather than elevated to special public variants.

## 19. Rustdoc

Every new public wrapper, raw formatting type, getter, setter, and iterator is documented. Crate documentation includes a compact formatting API table and a compiling `no_run` example covering Font, fill, number format, alignment, Border, AutoFit, and `MixedValue` matching.

## 20. Inventory

The registered 1.9 typelib maps the implemented Range, Font, Interior, Borders, and Border members with their actual DISPIDs and invocation kinds. Range has controlled formatting capability metadata. Mixed-capable members have `mixed_value_possible = true`. Borders is recorded as `Border`, `enum-key`, and iterator implemented.

## 21. Validation

Validation runs formatting, clippy, workspace tests, doctests, strict Rustdoc, inventory extraction/generation/check/diff, knowledge-base checking, all five public live tests, and the private physical-VARIANT diagnostic. Evidence omits pointers, PIDs, HWNDs, and machine paths.

## 22. Remaining blockers

No implementation blocker remains. Font availability, workbook palette behavior, automatic-color values, and other host-specific formatting results remain environment observations.

## 23. Recommended Prompt 13 scope

Build the next bounded slice around non-formatting Range behavior or workbook/worksheet operations, after fresh typelib and runtime investigation. Do not broaden this formatting layer into styles or conditional formatting without a separate design prompt.
