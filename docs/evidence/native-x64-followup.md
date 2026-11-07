# x64 acceptance follow-up: events, cancellation, lifecycle and Wizard

Scope #126; defects #137 and #138. Branch codex/native-excel-acceptance started
at fabda4d15ec573df16d94f7fd7f494debda1fad4. Remote master remained
26360e10cde0de5cb1d2d0f57121a3d2e6822ea0 at the final fetch. No merge or CI.

## Native results

- Manual F9 event delivery PASS: native-manual-event-delivery.md and JSON.
- Cooperative Escape cancellation PASS: native-cancellation-and-lifecycle.md.
  Full completion/cancellation race coverage remains separate/outstanding.
- Idle same-process deactivation and double registration retry PASS: same file.
- Wizard inspection found final-argument truncation; #138 is a native child of
  #95. Corrective source and regression tests pass. The corrected manual text
  confirmation is still pending; automatic accessibility assertion is BLOCKED.
- #137 remains OPEN. No unsafe stale-reference crash repetition or speculative
  xlSheetNm validity probe. Protected-cell rejection passed again.

Final noninteractive sweep 20261107T071359Z-c12b97bc used
acceptance-wizard-fixed.xll SHA-256
1342cc7e5f5e62897f894c141be586212ecc9d63a36171de0a73fe3f45fcaa11:
14 PASS, 0 FAIL, 6 BLOCKED. Blocked entries: two explicit manual-key opt-ins,
Wizard accessibility, stale crash opt-in, unimplemented autonomous wake,
and QUEUE worker teardown timeout. QUEUE's log had already recorded its
assertion and Quit; its owned PID was absent afterward. Focused retry
20261107T071632Z-f7bc773c at the normal 90-second deadline PASS. Neither run
was overwritten or consolidated into a misleading all-green sweep.

The binary was built from the modified source tree based on fabda4d plus
acceptance cancellation fixtures and #138's sentinel correction. Reports
record exact binary and runner hashes; their git revision alone is not proof
of pristine source provenance. Historical pinned assessments were unchanged.

## Validation

- PASS: cargo fmt --all
- PASS: cargo check --workspace --all-targets --all-features
- PASS: cargo test --workspace --all-features (after updating the existing
  exact-position test from 11 roots to 12 including the documented sentinel).
  Initial run FAILED that stale expectation; the failure log was preserved.
- PASS: cargo clippy -p excel-api -p minimal-xll --features
  minimal-xll/native-acceptance --all-targets -- -D warnings
- FAIL: cargo clippy --workspace --all-targets --all-features -- -D warnings:
  unchanged deprecated fetch_update calls in the RTD example at 599/608/728
  and ON.TIME research example at 90. These files match origin/master.
- PASS: python -B -m unittest discover -s tests/native_excel -p test_runner.py
  (15 tests, including operator opt-in and fresh async-retirement checks).
- PASS: cargo run --manifest-path tools/abi-check/Cargo.toml --offline
  (147 x64 SDK comparisons).
- PASS: python -B tools/excel-c-api-inventory/validate.py check (1,132 records).
  An initial invocation without the required check subcommand exited 2.
- PASS: git diff --check.

User supplied a Wizard image; the agent performed no screen capture. Pre-existing
PID 19684 was never attached to or terminated. No Office setting, installer,
ACL or security configuration changed. No x86 execution or acceptance claim.
The portable harness shares logic; x86 ABI/host acceptance remains #85/#126.
