# `excel-com` 0.1.0 release preparation

## Release scope and feature freeze

This document records release preparation for `excel-com` 0.1.0. The release
scope is the existing Windows desktop Excel COM Automation implementation; it
does not add an Excel object-model domain. A feature freeze is in effect.
Permitted changes are API and safety corrections, lifecycle regularization,
tests, documentation, packaging, compatibility work, and performance fixes
that do not expand semantics.

`excel-com` is an experimental Windows-only Excel Automation crate. Version
0.1 establishes the initial public API, but breaking changes remain possible
before 1.0.

## Public namespace policy

The crate root remains a reviewed 0.1 convenience facade for common workflows
and established Excel names. Specialized domains have canonical, discoverable
homes under `excel_com::drawing`, `external_data`, `pivot`, `presentation`,
`data`, `formatting`, and `tables`; new documentation and examples use those
paths. The complete root facade is deliberately frozen by the API snapshot so
that subsequent minor releases can reduce accidental duplication deliberately.

The 0.1 API review removed `Application::new`; callers create a clearly owned
session with `OwnedApplication::new(&apartment)` and use `application()`.
`ComMessageFilterGuard`, COM retry classifications, and the obsolete
`XlSheetVisibility` alias are internal. `ComRetryPolicy` and structured
`InvocationError` accessors remain public.

## Semantics fixed for 0.1

* Numeric Excel collection indexes are one-based. Zero is rejected before a
  COM call. Collection `count`, `item`/`item_by_*`, and fallible `_NewEnum`
  iteration preserve this policy.
* Wrapper `Clone` performs COM `AddRef` and creates another Rust handle to the
  same Excel object. It never duplicates an Excel workbook, worksheet, range,
  chart, or external-data object. All wrappers remain apartment-bound.
* `OwnedApplication` alone has `quit` and `quit_and_wait`. Its `Drop` never
  calls Excel `Quit`, never accepts unsaved-workbook prompts, and never force
  terminates a process; it releases COM references and best-effort restores a
  locally installed message filter. `AttachedApplication` only releases its
  references and has no shutdown API.
* `ComRetryPolicy` is opt-in, STA-local, bounded to three attempts and one
  second by default, and retries only safe reads or idempotent property writes.
  Methods, including `Workbooks.Add`, are not replayed after ambiguous COM
  delivery. Message-filter implementation details remain private.
* One optional Excel parameter is represented by `Option<T>`; multiple
  positions use typed `*Options`. Defaults retain Excel `Missing` arguments,
  and password/connection-string debug representations are redacted.
* `AutomationValue` distinguishes Empty, Null, text (including the empty
  string), number, bool, error SCODE, OLE Automation date, currency, and a
  rectangular `AutomationArray`. Dates remain timezone-neutral serials; Excel
  selects the 1900 or 1904 workbook date system and retains its historical
  1900 leap-year compatibility behavior.
* Formatting uses `MixedValue::Uniform`, `Mixed`, and `Empty`; a mixed Excel
  selection is never silently coerced. Transparent Excel enum wrappers retain
  unknown values and use `from_raw`/`raw` consistently.

## Live-test policy

| Tier | Scope | 0.1 release use |
|---|---|---|
| 0 | No Excel process | Unit, compile-fail, packaging, docs |
| 1 | Copied fixture open/read/close | Required fixture smoke baseline |
| 2 | Owned copied-fixture mutation | Release smoke test when Excel is available |
| 3 | Global Application state | Opt-in, controlled tests only |
| 4 | Provider/account-dependent external data | Opt-in and may remain blocked |

`Workbooks.Add` is structurally covered but not a release acceptance test on
this host. An independent tool reproduced this host's `0x800A03EC` failure.
Release smoke tests therefore copy repository-owned fixtures and use
`Workbooks.Open`; another Windows/Excel installation should verify Add after
release preparation.

## Compatibility and security policy

The declared MSRV is Rust 1.85. Windows desktop Excel through COM Automation
is the sole supported runtime target. x64 is compile- and runtime-tested; x86
is compile-tested only. macOS Excel, Excel Online, Linux-native Excel,
LibreOffice, Office Scripts, and Microsoft Graph workbook APIs are unsupported.

`macro-runtime` is the only opt-in feature. It enables `Application::run_macro`
and is off by default because it can execute workbook-provided code in the
active Excel process. All other current functionality is available in the base
crate; future features will be additive. MSRV increases are significant changes
and will be disclosed in release notes.

High-risk boundaries are documented rather than hidden: attachment does not
confer ownership; external data and links are provider-dependent; passwords,
connection strings, Power Query formulas, and macro arguments must not be
logged; file operations may encounter Excel prompts; PDF export and PageSetup
can be printer dependent.
