# Extended x64 acceptance: protected and stale targets

Owners #126/#127, investigation #137 under #93. The extended acceptance is not
complete and no further subsystem is closed.

Native artifact: release/native-acceptance build from base `795f17e` plus the
capture/write fixtures under development, SHA-256
`f02edb333a1cd35901e502b0211fbfbd91a089b857be40741c9ece49b1cb53f0`.
This is not an assertion that the final committed source was the exact build
input: the final harness also adds opt-in protection and a checked conversion.

`REF-PROTECTED` passed in run `20261107T062118Z-f42b7e35`: protected scratch J1
remained 123 and the write command returned failure. The original EVENTS probe,
now reading the counter through a fresh invocation, still did not observe a
notification after posted F9. Status BLOCKED, not a verified callback defect.

`REF-STALE` in run `20261107T062204Z-6266b2ac` captured an actual scratch sheet ID,
deleted that sheet, and invoked xlSet with its saved descriptor. The automation
connection failed with -2147023170, then -2147023174 during cleanup. This is a
host-failure observation, not successful rejection. Windows Application event
1001 concurrently recorded OFFICE_MODULE_VERSION_MISMATCH for Excel
16.0.20527.20150; report ID 473c96e8-25fe-4a7c-a846-de0e925080bc. Causation remains
unconfirmed. No retry of the failing scenario or forced Excel termination occurred.
An existing Excel process prevented the subsequent EVENTS attempt; thus the new
ExecuteMso CalculateNow path has not been validated.

Ordinary sweeps now BLOCK REF-STALE. A diagnostic-only explicit
`--allow-stale-probe` opt-in exists, with a regression test proving ordinary
selection never launches its worker. Do not repeat on this host until #137 is
triaged. The fixture takes its saved descriptor once and never owns workbook
lifetime. No production code or ABI was changed in this extension.

Cancellation races, Function Wizard metadata and same-process unload/retry were
not run in this extension. Event delivery is still unresolved. x86 remains
separate and NOT RUN. No CI, screen capture or Office repair was performed.
