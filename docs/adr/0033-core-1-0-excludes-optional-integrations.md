# ADR-0033: Core 1.0 excludes optional RTD and COM/Ribbon subsystems

## Status

Accepted.

## Context

The Excel 12/XLL core already provides the safety-critical path for ordinary
native Excel UDFs: ABI bindings, borrowed and owned value layers, stable return
ownership, registration, generated thunks, typed callback contexts/calls,
lifecycle, diagnostics, packaging, asynchronous one-shot UDFs, and cooperative
callback-drained dispatch.

RTD and general COM/Ribbon use a separate lifecycle and deployment model.
The RTD prototype has useful local ABI/direct-activation evidence, but its
Excel-formula lifecycle and policy compatibility remain unverified. Requiring
those optional subsystems would delay review of an independently useful core.

## Decision

The initial stable 1.0 release excludes RTD/streaming, general COM automation,
Ribbon UI, custom task panes, autonomous dispatcher notification, and
`xlcOnTime` research. M18 is parked as experimental research; M19 is deferred;
M20 core stabilization is next.

The RTD prototype remains in the repository as Windows-only, unpublished,
unsupported experimental code. It stays a workspace member for explicit
`--workspace` validation where practical, but it is not a default workspace
member, never enters core package dependency graphs, and is excluded from
normal XLL packaging. The default core workspace members are the three publish
candidate crates plus the minimal XLL example.

## Consequences

- Core 1.0 review focuses on the native Excel C API product boundary.
- Core package publication is limited to `excel-api-sys`, `excel-api`, and
  `excel-api-macros`; prototype packages remain `publish = false`.
- Issue #30 remains optional autonomous-notification research, not a 1.0
  blocker. Issue #37 tracks optional RTD follow-up.
- Real Excel validation still matters for implemented async, dispatcher, and
  stress capabilities, but it does not turn optional RTD/COM/Ribbon into a
  release prerequisite.
- Future optional integrations require separate ADRs and support/deployment
  commitments; they may not manufacture Excel C API callback capability.
