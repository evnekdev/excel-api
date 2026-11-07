# Native Excel acceptance runner (partial)

Extended probes: REF-PROTECTED checks rejection and preserved cell contents.
REF-STALE is BLOCKED by default after a host RPC failure (#137); only controlled
diagnostic hosts may opt in with `--allow-stale-probe`. Do not repeat the stale
probe until the correlated Office module-version-mismatch report is triaged.

This shared Python runner and catalogue support Windows x64 and x86 as test
infrastructure. It is **not** x86 ABI support or x86 host acceptance. #85 owns
the unsupported x86 XLL ABI and build investigation. #126 owns this suite.
Fourteen bounded probes are implemented against the repository's `minimal-xll`
with its explicit `native-acceptance` feature. These cover basic read/write/clear,
active-workbook information, normal recalculation, HelpTopic-bearing registration,
event observation and fresh-session exit as well as the original seven probes.
They do not establish the full cancellation, protection, rollback, named-document,
Function Wizard, or same-process unload/reload matrix. The F9 event observation
can be BLOCKED when targeted keyboard-message delivery does not raise an event.
A basic reference-kind probe is not an xlCoerce read.
Autonomous wake is not implemented (#30/#101). This is not a complete sweep.

## Requirements and execution (Windows CMD)

- Windows desktop Excel and a matching, trusted XLL. The runner records all
  pre-existing Excel PIDs and never attaches to or closes them. DispatchEx must
  return a window belonging to the sole newly created PID; reused or ambiguous
  instances are rejected. It never changes global Office security.
- Python 3.9 or newer and pywin32. Install in an isolated virtual environment,
  or prepare a wheelhouse on a connected machine and transfer it with the
  bundle. A fresh Python environment can use:

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r tests\native_excel\requirements.txt
cargo build -p minimal-xll --release --features native-acceptance
copy target\release\minimal_xll.dll target\release\acceptance.xll
python tests\native_excel\run_acceptance.py --check-environment --output results
python tests\native_excel\run_acceptance.py --xll C:\ExcelTests\minimal_xll.xll --output results
python tests\native_excel\run_acceptance.py --xll C:\ExcelTests\minimal_xll.xll --scenario REG-ADD --output results
python -m unittest discover -s tests\native_excel -p test_runner.py
```

To use offline, download compatible pywin32 wheels in advance and install with
`python -m pip install --no-index --find-links C:\wheels -r tests\native_excel\requirements.txt`.
Copy this entire directory and its dependencies to the target machine; no GitHub
credentials are required. For x86 supply a separately built, validated x86 XLL.
No supported x86 XLL build is promised: coordinate ABI prerequisites with #85.
Never use the x64 XLL as an x86 substitute. PE machine mismatch is rejected.

Each run creates a unique directory containing `report.json`, `summary.txt`, and
per-probe logs. Transfer that directory for review after checking for sensitive
paths. PASS is a completed live assertion; FAIL is a failed assertion; BLOCKED
is a prerequisite or automation failure; NOT RUN means no implemented probe.
The recorded git revision is the runner checkout when available, **not proof**
of XLL provenance. Record an independently reproducible source/build chain.

Every implemented probe runs in a separate worker with an overall deadline
(`--timeout`, 90 seconds by default). The worker checks process inventory,
uses DispatchEx and verifies the returned window PID against the initial
inventory before touching a scratch workbook or requesting Quit. If
activation hangs, the parent stops only its Python worker and does not kill
Excel; inspect Task Manager and any unsaved work before manually closing an
unknown Excel process. Existing PIDs remain untouched; an ambiguous new-process
inventory blocks the probe. Cleanup restores changed application settings and closes only the
scratch workbook. A timeout or failed cleanup is not a successful test.

The CLEANUP probe checks a fresh session and observes process exit; it does not
prove same-process unregister, retry or reload behavior. The HELP probe verifies
registration with a supplied HelpTopic, not Function Wizard rendering or opening
the resource. The runner does not force unload an XLL. If an
async result is not observed, it does not explicitly close the workbook or
request Quit: inspect the logged owned PID and resolve pending work safely.
Tests requiring pending-work unload remain NOT RUN. Excel security prompts, host UI, callback races,
and internal fault injection are not automatically resolved by this runner.
