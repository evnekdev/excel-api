# excel-api user guide

This guide is the user-facing companion to the Rustdoc for the three published
crates. Architecture documents explain maintainer decisions; this guide explains
how an XLL author uses the supported surface.

## Status vocabulary

| Status | Meaning |
| --- | --- |
| **Stable target** | Included in the intended 1.0 XLL core, subject to the release checklist. |
| **Preview** | Implemented and tested deterministically, but still awaiting the listed real-Excel lifecycle validation. |
| **Experimental** | Retained research/prototype work with no production support promise. |
| **Deferred** | Deliberately outside the initial release. |
| **Unsupported** | Do not rely on this configuration or behavior. |

The stable target is 64-bit Windows Excel using the Excel 12 C API. RTD,
general COM, Ribbon, task panes, `xlcOnTime`, and autonomous notifications are
experimental or deferred; they are not core XLL features.

## Pages

- [Getting started](getting-started.md)
- [Project layout](project-layout.md)
- [Worksheet functions](worksheet-functions.md)
- [Values and conversions](values-and-conversions.md)
- [Strings](strings.md)
- [Arrays](arrays.md)
- [References](references.md)
- [Optional, missing, empty, and nil](optional-missing-empty-and-nil.md)
- [Errors](errors.md)
- [Commands](commands.md)
- [Lifecycle](lifecycle.md)
- [Excel calls and contexts](excel-calls-and-contexts.md)
- [Thread-safe functions](thread-safe-functions.md)
- [Async UDFs](async-udfs.md)
- [Cooperative dispatcher](cooperative-dispatcher.md)
- [Packaging](packaging.md)
- [Signing and trust](signing-and-trust.md)
- [Testing](testing.md)
- [Troubleshooting](troubleshooting.md)
- [Migration from Excel-DNA](migration-from-excel-dna.md)
- [Macro reference](macro-reference.md)
