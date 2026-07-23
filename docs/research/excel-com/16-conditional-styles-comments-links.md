# Prompt 16: conditional formatting, styles, comments, and hyperlinks

## 1. Scope and baseline

This prompt adds the bounded Excel presentation surface to `excel-com` from
`implementation/excel-com-16-conditional-styles-comments-links`. The recorded
fresh `Workbooks.Add` baseline passed and the owned workbook exited naturally.

## 2. Prompt 15 runtime status

Prompt 15's historical runtime-blocked metadata is retained separately in the
inventory model. Prompt 16 does not reinterpret those historical observations.

## 3. Presentation-module split

The former `presentation.rs` façade now re-exports focused sheets, window,
layout, output, macro-runtime, operations, conditional, styles, comments, and
hyperlinks modules. Existing public names remain at the façade.

## 4. Conditional-format object model

`FormatConditions` returns a typed `ConditionalFormat` enum without exposing
raw dispatch. Known `Type` values select ordinary, scale, bar, icon, top,
average, and unique-value wrappers; unknown values remain explicit.

## 5. Rule collections and subtype detection

Collection access is one-based and rejects zero before COM. Enumeration owns
an `IEnumVARIANT` cursor and fuses after an error. Subtyping uses Excel's
stable `Type` property, not raw-pointer exposure.

## 6. Cell-value and expression rules

Typed option records retain the `Add` positions. Formula text is not parsed;
Excel validates it and interprets relative formulas from the upper-left cell
of `AppliesTo`.

## 7. Priority and StopIfTrue

Rules expose Excel priority mutation and `StopIfTrue`. A clean-state retest is
still required to answer ordering-after-deletion and persistence questions.

## 8. Colour scales

Two- and three-criterion scales are supported. Criteria have typed threshold
type/value access and `FormatColor` access.

## 9. Data bars

Data bars expose typed min/max threshold objects, colour, direction, axis,
fill, visibility, priority, and applies-to access.

## 10. Icon sets

The first controlled call established that `AddIconSetCondition` takes no
argument in this host. The wrapper now creates it without arguments and sets
`IconSet` afterward; retest remains pending.

## 11. Top/bottom and average rules

Top/bottom rank, percent, and direction plus above/below-average settings are
represented. Excel remains responsible for unsupported combinations.

## 12. Duplicate and text rules

Unique/duplicate and typed text rule construction are present. Time-period
construction is mapped through Excel's generic `Add` argument positions.

## 13. DisplayFormat

`DisplayFormat` is read-only and distinct from underlying Range formatting.
Its host-specific timing and mixed-value behavior remain pending live retest.

## 14. Styles collection

`Workbook::styles` has one-based/name lookup, fallible enumeration, and an
optional Range `BasedOn` argument for custom-style creation.

## 15. Style assignment

Range style names preserve mixed results. Assignment supports a style name and
the Style dispatch object; host behavior is recorded only after live evidence.

## 16. Theme colours

`ThemeColor` is a transparent forward-compatible value. Theme-aware members
were added where exposed by Font, Interior, Border, Tab, and FormatColor.

## 17. Tint and shade

Every new tint/shade setter rejects non-finite values and values outside
`-1.0..=1.0`. RGB/theme precedence remains Excel-defined and unobserved here.

## 18. Advanced Font/Interior/Border properties

The report-oriented Font, Interior, and Border members include script,
shadow, outline, theme-font, theme-colour, tint, and pattern tint access.

## 19. Legacy comments and Notes

`Comment` represents Excel's legacy Comment object, which modern Excel labels
a Note. It is intentionally not conflated with threaded comments.

## 20. Threaded comments

Read-only text, author identity, date, and replies are exposed structurally.
Creation is intentionally absent because it is identity and service dependent.

## 21. Hyperlinks

Collections, optional `Address`/`SubAddress`, display text, screen tip, range,
and deletion are represented. The API neither validates nor follows targets.

## 22. Persistence

Save/reopen verification is pending the clean-state controlled live rerun and
is not inferred from setters.

## 23. Runtime observations

`Workbooks.Add` passed once. The first conditional run reached cell, scale,
and data-bar construction before an incorrect argument-bearing icon-set call.
After the visible instance was closed, the corrected retest's fresh
`Workbooks.Add` returned `0x800A03EC`; no repeat or host-configuration
investigation was performed.

## 24. Explicit non-decisions

No chart, shape, picture, OLE, control, VBA-editing, external-data, PivotTable,
event, raw-COM, `Send`, or `Sync` API was added.

## 25. Rustdoc

The crate overview documents rule ordering, formula-relative behavior,
`StopIfTrue`, DisplayFormat, styles, theme/RGB interaction, Notes, threaded
comments, hyperlink behavior, and apartment affinity.

## 26. Inventory

Every invoked wrapper member has a registry ID. The generator records typed
collection element metadata, iterator status, heterogeneous rule policy, and
Range advanced-presentation capability metadata.

## 27. Validation

Non-live compilation is run during implementation. The final validation set
includes workspace formatting, check, test, Clippy, Rustdoc, inventory, KB,
and diff checks.

## 28. Remaining blockers

The fresh-add anomaly blocks further controlled live coverage in this prompt.
No process is terminated by the task.

## 29. Recommended Prompt 17 scope

Use the clean-state persistence observations to close runtime questions, then
move to the next bounded object-model family rather than broad UI automation.
