# Command registration correction

Tracking: [#124](https://github.com/evnekdev/excel-api/issues/124), native child of
[#96](https://github.com/evnekdev/excel-api/issues/96); related metadata work #95.
Starting source: `80bfdf3caf905d8b3e6bec33f80abfbca0df1d1c`.

## Contract and correction

Microsoft's [xlfRegister Form 1](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1)
defines this positional sequence. The corrected no-argument command supplies:

| Position | Argument | Value / Excel type |
| --- | --- | --- |
| 1 | Module | DLL path / string |
| 2 | Procedure | exported symbol / string |
| 3 | Type text | `I` / string |
| 4 | Command name | Excel-visible name / string |
| 5 | Argument text | empty string |
| 6 | Macro type | `2` / number |
| 7 | Category | `Commands` / string |
| 8 | Shortcut | configured character or empty string |
| 9 | Help topic | empty string |
| 10 | Function help | description or empty string |

Worksheet functions use macro type 1 and may append argument-help strings after
position 10. Commands have no such argument help. Optional omission does not
shift subsequent positions; this implementation supplies explicit empty strings.
Shortcut case has meaning in the documented command contract.

Ten is the library's full-metadata convention, **not Excel's universal minimum**:
the pinned SDK's `SAMPLES/GENERIC/GENERIC.C:150–175,429–438` registers commands with
eight arguments, ending at shortcut. `XLCALL.H` provides the function identifier
and callback signatures, not the positional metadata contract. Thus lowering
the local minimum alone would leave incorrect metadata positions.

The old builder omitted category: shortcut landed in category, the empty help
topic in shortcut, and description in help topic. Its nine roots were rejected
by the library before dispatch. Inserting the fixed category corrects both the
positions and count without a public API or general HelpTopic extension.

Existing counted UTF-16 backing allocations and the roots vector stay live through
the synchronous call. LifecycleContext, CallCapability, ExcelOwnedValue, registration
ID handling, xlFree, rollback and unregister code paths are reused unchanged.

## Regression and acceptance scope

The command-bearing runtime regression failed before the fix with
`InvalidArgumentCount { function: "xlfRegister", count: 9 }`.
Mock snapshots now check every positional value and Excel tag, both populated and
absent optional metadata, actual xlfRegister dispatch and stored command IDs.
Mixed fixtures exercise worksheet registration first and command registration
second. Tests also cover backend failure, a successful callback returning an error
root, rollback of prior commands/functions, failed rollback cleanup followed by
close retry, successful command unregister, invalid metadata and an illegal
worksheet callback context.

Tests use the existing mock backend and runtime. The public registration_ids
accessor reports worksheet IDs only; command IDs are checked in runtime state.
The existing close behavior is preserved: worksheet group first, then command
group, each reversed. The historical architecture's broader wording about global
reverse order does not describe that implementation; any change to that behavior
belongs to #96 and is outside this argument-construction correction.

Native acceptance remains **NOT RUN**, tracked on #124: load an XLL with commands,
invoke the registered command, test configured/absent shortcuts, unload/reload,
and check hidden-name/registration cleanup and stale shortcuts. No host success
follows from mocks. Broader HelpTopic support remains under #95.

The source-to-Rust assessment remains historical at `172b306`; its current-input
validator should request reassessment after these production paths change. Its
pinned defect observation is retained rather than rewriting historical evidence.

Exact verification commands and results are recorded on #124 with the correction
commit. No SDK/ABI layout, canonical identifier, dependency or macro-generator
change is part of this fix; native ABI and Excel automation were not run.

Local verification on Windows x64, Rust 1.99.1:

| Command | Result |
| --- | --- |
| `cargo test -p excel-api --lib` | PASS, 136 tests including six new regressions |
| `cargo test -p excel-api --test macro_metadata` | PASS, 5 tests |
| `cargo test -p excel-api-macros --tests` | PASS, 7 unit tests and 1 trybuild harness (4 fixtures) |
| `cargo test -p minimal-xll --lib --offline` | PASS, 14 tests |
| `cargo check -p excel-api-sys -p excel-api -p excel-api-macros -p minimal-xll --all-targets --all-features` | PASS, existing OnTime deprecation warning |
| `cargo clippy -p excel-api-sys -p excel-api -p excel-api-macros --all-targets --all-features -- -D warnings` | PASS |
| Same Clippy command also selecting `-p minimal-xll` | FAIL, pre-existing `fetch_update` deprecation in optional ontime_experiment.rs:90 |
| `cargo fmt --all` | Completed |
| `git diff --check` | PASS |

The first corrected-suite run exposed two test assumptions about the worksheet-only
registration_ids accessor; those assertions were corrected to inspect command IDs
in runtime state. No additional production change was needed. Broad workspace
validation and CI were not run; focused core/example checks cover this change.
