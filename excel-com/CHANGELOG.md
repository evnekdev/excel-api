# Changelog

All notable changes to `excel-com` are documented in this file.

The crate follows the spirit of [Keep a Changelog](https://keepachangelog.com/)
and uses pre-1.0 semantic versioning: breaking changes can occur in minor
releases and are called out in the release notes.

## Unreleased

### Added

- Completed the generic workbook-migration surface with typed
  `Application.UserControl` and `Application.EnableEvents` access, an
  `EnableEventsGuard`, public `Range::worksheet`, explicit
  `Name::refers_to_range`, and a typed `Workbooks.Open(UpdateLinks := 0)`
  constant.
- Added real-Excel migration lifecycle and private-session isolation coverage.

### Fixed

- Made `Names::item_by_name` enforce the originating workbook or worksheet
  scope. Excel's ambiguous raw workbook collection lookup can otherwise return
  a worksheet-local name with the same text as a global name.

## 0.1.0 - 2026-07-23

Initial public, experimental release preparation for Windows desktop Excel COM
Automation. The release establishes the ownership model, typed workbook and
range API, formatting, tables, drawings and charts, text/data transformation,
external-data, and PivotTable surfaces. See the repository release notes for
the supported scope and known limitations.
