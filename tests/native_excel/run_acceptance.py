#!/usr/bin/env python3
"""Bounded native Excel probes. Unimplemented inventory entries never report PASS."""
import argparse
import ctypes
import datetime as dt
import hashlib
import importlib.util
import json
import os
import platform
import struct
import subprocess
import sys
import time
import uuid
from pathlib import Path

# Existing #126 identifiers are retained. Each unimplemented case stays visible.
SCENARIOS = {
    "REG-ADD": ("registration", "RUST.ADD(2,3) equals 5", True),
    "VALUE-TEXT": ("values", "Unicode echo round trips", True),
    "VALUE-OPTION": ("values", "An omitted optional argument is missing", True),
    "REF-KIND": ("references", "A worksheet reference reaches the function", True),
    "COMMAND": ("registration", "The registered ping command executes", True),
    "QUEUE": ("dispatcher", "Worker work waits for an explicit command pump", True),
    "ASYNC": ("async", "Async double completes with 14", True),
    "ASYNC-CANCEL": ("async", "Manual F9 then Escape cancels bounded work and retires its worker", True),
    "REF-READ": ("references", "xlCoerce reads a scratch cell", True),
    "REF-WRITE": ("references", "Command-only xlSet writes and clears J1", True),
    "REF-PROTECTED": ("references", "Protected J1 rejects xlSet without changing its value", True),
    "REF-STALE": ("references", "Deleted-sheet descriptor rejects write and preserves surviving sheet", True),
    "DOC": ("document", "Active workbook name matches scratch workbook", True),
    "CALC": ("calculation", "Command calculates dirty formula in manual mode", True),
    "HELP": ("help", "Function with HelpTopic supplement registers and calculates", True),
    "HELP-WIZARD": ("help", "Function Wizard exposes function and argument descriptions", True),
    "EVENTS": ("events", "Observe calculation-ended callback after F9", True),
    "CLEANUP": ("lifecycle", "Fresh-session registration and owned process exit", True),
    "UNLOAD-RETRY": ("lifecycle", "Idle XLL deactivation and registration retry in the same Excel process", True),
    "WAKE": ("dispatcher", "Autonomous wake (#30/#101; not implemented)", False),
}


def pe_arch(path):
    with open(path, "rb") as f:
        if f.read(2) != b"MZ":
            raise ValueError("missing DOS PE signature")
        f.seek(0x3c)
        raw = f.read(4)
        if len(raw) != 4:
            raise ValueError("truncated DOS header")
        offset = struct.unpack("<I", raw)[0]
        if offset > os.fstat(f.fileno()).st_size - 6:
            raise ValueError("invalid PE offset")
        f.seek(offset)
        if f.read(4) != b"PE\0\0":
            raise ValueError("missing PE signature")
        machine = struct.unpack("<H", f.read(2))[0]
    if machine not in (0x14c, 0x8664):
        raise ValueError(f"unsupported PE machine 0x{machine:04x}")
    return {0x14c: "x86", 0x8664: "x64"}[machine]


def excel_path():
    if os.name != "nt":
        return None
    import winreg
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", 0, winreg.KEY_READ | view) as key:
                    return winreg.QueryValueEx(key, None)[0]
            except OSError:
                pass
    return None


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def environment(xll=None):
    result = {"windows": platform.platform(), "python": sys.version.split()[0],
              "python_arch": "x64" if struct.calcsize("P") == 8 else "x86",
              "windows_arch": platform.machine(), "pywin32": importlib.util.find_spec("win32com") is not None,
              "excel_executable": excel_path(), "excel_arch": None, "xll": None}
    if result["excel_executable"]:
        try:
            result["excel_arch"] = pe_arch(result["excel_executable"])
        except (OSError, ValueError) as exc:
            result["excel_detection_error"] = str(exc)
        try:
            import win32api
            version = win32api.GetFileVersionInfo(result["excel_executable"], "\\")
            ms, ls = version["FileVersionMS"], version["FileVersionLS"]
            result["excel_file_version"] = ".".join(str(x) for x in (ms >> 16, ms & 65535, ls >> 16, ls & 65535))
        except (ImportError, OSError, KeyError) as exc:
            result["excel_file_version_error"] = str(exc)
    try:
        result["source_revision"] = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=Path(__file__).resolve().parents[2], stderr=subprocess.DEVNULL, text=True, timeout=5).strip()
    except (OSError, subprocess.SubprocessError):
        result["source_revision"] = None
    result["source_binary_provenance"] = "not independently established"
    result["runner_sha256"] = sha256(__file__)
    if os.name == "nt":
        try:
            result["existing_excel_pids"] = sorted(process_ids())
        except (OSError, subprocess.SubprocessError) as exc:
            result["process_inventory_error"] = str(exc)
    if xll:
        try:
            result["xll"] = {"path": str(Path(xll).resolve()), "architecture": pe_arch(xll), "sha256": sha256(xll)}
        except (OSError, ValueError) as exc:
            result["xll_error"] = str(exc)
    return result


def validate_environment(env, need_xll=True):
    errors = []
    if os.name != "nt": errors.append("Windows is required")
    if not env["pywin32"]: errors.append("pywin32 is not installed")
    if not env["excel_executable"] or not env["excel_arch"]: errors.append("Excel executable/architecture unavailable")
    if need_xll and not env["xll"]: errors.append("valid XLL required: " + env.get("xll_error", "not supplied"))
    if env["xll"] and env["excel_arch"] and env["xll"]["architecture"] != env["excel_arch"]:
        errors.append("XLL architecture does not match Excel architecture")
    if env.get("process_inventory_error"): errors.append("Excel process inventory unavailable")
    return errors


def process_ids():
    """Use tasklist only for conservative conflict detection, not ownership."""
    output = subprocess.check_output(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"], text=True, timeout=10)
    import csv
    return {int(row[1]) for row in csv.reader(output.splitlines()) if len(row) > 1 and row[0].lower() == "excel.exe" and row[1].isdigit()}


def retire_async_probe(app, sheet):
    # Remove the trigger before a fresh quiescence check. Never authorize Quit
    # using a cached status from before manual input or a rejected COM query.
    sheet.Range("C1").ClearContents()
    status = tuple(int(x) for x in app.Run("RUST.TEST.ASYNC.STATUS").split(","))
    if len(status) != 6 or status[0] != 0 or status[3] != 0:
        raise RuntimeError("Fresh worker retirement check failed; retaining owned instance")
    return status


def worker(scenario, xll, event_file, manual_ui=False):
    def event(message):
        with open(event_file, "a", encoding="utf-8") as out:
            out.write(message + "\n")
    app = book = None
    owned = False
    owned_pid = None
    async_pending = False
    original = {}
    try:
        before = process_ids()
        event(f"activation started; pre-existing Excel PIDs={sorted(before)}; these will not be touched")
        import win32com.client
        app = win32com.client.DispatchEx("Excel.Application")
        hwnd = int(app.Hwnd)
        pid = ctypes.c_ulong()
        if not ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid)) or not isolated_pid(before, process_ids(), pid.value):
            raise RuntimeError("Excel application process ownership could not be established")
        owned = True
        owned_pid = pid.value
        event(f"owned Excel application hwnd={hwnd} pid={pid.value}; version={app.Version}; build={app.Build}")
        for name in ("DisplayAlerts", "ScreenUpdating"):
            original[name] = getattr(app, name)
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        book = app.Workbooks.Add()
        sheet = book.Worksheets(1)
        sheet.Range("B10").Formula = "=SUM(2,3)"
        if sheet.Range("B10").Value != 5:
            raise RuntimeError(f"built-in formula control failed: {sheet.Range('B10').Value!r}")
        event("built-in SUM control returned 5")
        if not app.RegisterXLL(str(Path(xll).resolve())):
            raise AssertionError("RegisterXLL returned false")
        event("scratch workbook created and XLL registered")
        try:
            registered = app.RegisteredFunctions
            matches = [str(row)[:300] for row in registered if "RUST.ADD" in str(row).upper()] if isinstance(registered, tuple) else []
            event(f"RegisteredFunctions rows containing RUST.ADD: {matches[:3]!r}")
        except Exception as exc:
            event(f"RegisteredFunctions inspection unavailable: {exc}")
        if scenario == "HELP-WIZARD":
            original["Visible"] = app.Visible
            app.Visible = True
            app.ScreenUpdating = True
            app.DisplayAlerts = True
            sheet.Range("A1").Formula = "=RUST.TEST.HELP(17)"
            sheet.Range("A1").Select()
            import threading
            import win32gui
            import win32process
            labels = []
            dialog_titles = set()
            def inspect_dialog():
                deadline = time.monotonic() + 10
                while time.monotonic() < deadline:
                    windows = []
                    win32gui.EnumWindows(lambda h, _: windows.append(h), None)
                    for window in windows:
                        if win32process.GetWindowThreadProcessId(window)[1] != owned_pid:
                            continue
                        dialog_titles.add((win32gui.GetClassName(window), win32gui.GetWindowText(window)))
                        if win32gui.GetWindowText(window) != "Function Arguments":
                            continue
                        time.sleep(.3)
                        win32gui.EnumChildWindows(window, lambda h, _: labels.append(win32gui.GetWindowText(h)), None)
                        script = ("[Console]::OutputEncoding=[Text.UTF8Encoding]::new(); "
                                  "Add-Type -AssemblyName UIAutomationClient; "
                                  f"$root=[System.Windows.Automation.AutomationElement]::FromHandle([intptr]{window}); "
                                  "$root.FindAll([System.Windows.Automation.TreeScope]::Subtree,[System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name; $_.Current.HelpText; "
                                  "$legacy=$null; if ($_.TryGetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern,[ref]$legacy)) { $legacy.Current.Name; $legacy.Current.Description; $legacy.Current.Help; $legacy.Current.Value } }")
                        try:
                            ui = subprocess.run(["powershell", "-NoProfile", "-Command", script], capture_output=True, encoding="utf-8", errors="replace", timeout=8)
                            labels.extend(ui.stdout.splitlines())
                            msaa = subprocess.run(["powershell", "-NoProfile", "-File", str(Path(__file__).with_name("inspect_wizard.ps1")), "-WindowHandle", str(window)], capture_output=True, encoding="utf-8", errors="replace", timeout=8)
                            labels.extend(msaa.stdout.splitlines())
                            if msaa.returncode:
                                event("MSAA inspection failed: " + msaa.stderr[:1000])
                        except subprocess.TimeoutExpired:
                            labels.append("UI Automation observation timed out")
                        if manual_ui:
                            event("Wizard held open for up to 180 seconds for operator text observation")
                            hold_until = time.monotonic() + 180
                            while time.monotonic() < hold_until and win32gui.IsWindow(window):
                                time.sleep(.25)
                        if (not win32gui.IsWindow(window)
                                or win32process.GetWindowThreadProcessId(window)[1] != owned_pid
                                or win32gui.GetWindowText(window) != "Function Arguments"):
                            return
                        win32gui.PostMessage(window, 0x0010, 0, 0)
                        return
                    time.sleep(.1)
            watcher = threading.Thread(target=inspect_dialog, daemon=True)
            watcher.start()
            sheet.Range("A1").FunctionWizard()
            watcher.join(timeout=20)
            event("owned wizard-window inventory: " + repr(sorted(dialog_titles)))
            observed = " | ".join(x for x in labels if x)
            event("wizard labels: " + observed[:4000])
            if "Acceptance help metadata" not in observed or "Acceptance argument" not in observed:
                raise RuntimeError("Expected wizard descriptions were not accessible through native child-window text")
        elif scenario == "EVENTS":
            original["Visible"] = app.Visible
            app.Visible = True
            sheet.Range("A1").Formula = "=RUST.TEST.EVENTS()"
            original["Calculation"] = app.Calculation
            app.Calculation = -4135  # xlCalculationManual
            sheet.Range("B1").Value = 20
            sheet.Range("C1").Formula = "=RUST.ADD(B1,1)"
            sheet.Range("B1").Value = 21
            before_events = app.Run("RUST.TEST.EVENTS")
            import win32gui
            import win32process
            app.ScreenUpdating = True
            app.Caption = "Excel API acceptance — focus this scratch window"
            event(f"Waiting up to 60 seconds for user focus of scratch Excel PID {owned_pid}")
            deadline = time.monotonic() + 60
            while time.monotonic() < deadline:
                foreground = win32gui.GetForegroundWindow()
                if foreground and win32process.GetWindowThreadProcessId(foreground)[1] == owned_pid:
                    break
                time.sleep(.1)
            foreground = win32gui.GetForegroundWindow()
            if win32process.GetWindowThreadProcessId(foreground)[1] != owned_pid:
                raise RuntimeError("Scratch Excel did not receive foreground focus for the manual F9 test")
            event("Scratch Excel focused; waiting for user F9 (no injected keyboard input)")
            import pywintypes
            observed = before_events
            deadline = time.monotonic() + 45
            while time.monotonic() < deadline:
                time.sleep(.5)
                try:
                    observed = app.Run("RUST.TEST.EVENTS")
                except pywintypes.com_error as exc:
                    # Excel temporarily rejects automation while processing UI input.
                    nested = exc.excepinfo[5] if exc.excepinfo else None
                    if exc.hresult in (-2146777998, -2147418111) or nested == -2146777998:
                        continue
                    raise
                if observed > before_events: break
            event(f"calculation-ended count before={before_events!r}, after={observed!r}")
            calculated = sheet.Range("C1").Value
            event(f"F9 dirty-formula result={calculated!r}; expected 22")
            if calculated != 22:
                raise RuntimeError("F9 did not calculate the dirty scratch formula")
            if observed <= before_events:
                raise RuntimeError("No calculation-ended notification observed after foreground F9; UI delivery is not guaranteed")
        elif scenario == "UNLOAD-RETRY":
            status = tuple(int(x) for x in app.Run("RUST.TEST.ASYNC.STATUS").split(","))
            if status[0] or status[3]:
                async_pending = True
                raise RuntimeError("Refusing deactivation with pending requests or active workers")
            if app.Run("RUST.ADD", 2, 3) != 5:
                raise AssertionError("Pre-deactivation function control failed")
            module = str(Path(xll).resolve()).replace('"', '""')
            unloaded = app.ExecuteExcel4Macro('UNREGISTER("' + module + '")')
            event(f"Host UNREGISTER module result={unloaded!r}")
            if unloaded is not True:
                raise RuntimeError("Host did not confirm idle XLL deactivation")
            import pywintypes
            try:
                after_unload = app.Run("RUST.ADD", 2, 3)
                event(f"Function result after deactivation={after_unload!r}")
            except pywintypes.com_error:
                event("Function invocation rejected after deactivation")
            else:
                if after_unload == 5:
                    raise AssertionError("Function remained callable after deactivation")
                if after_unload not in (-2146826259, -2146826273):
                    raise RuntimeError(f"Unexpected result after deactivation: {after_unload!r}")
            for attempt in range(2):
                if not app.RegisterXLL(str(Path(xll).resolve())):
                    raise AssertionError(f"Registration retry {attempt+1} rejected")
                if app.Run("RUST.ADD", 2, 3) != 5:
                    raise AssertionError(f"Function failed after registration retry {attempt+1}")
                if not app.Run("RUST.PING.COMMAND"):
                    raise AssertionError(f"Command failed after registration retry {attempt+1}")
            final_pid = ctypes.c_ulong()
            ctypes.windll.user32.GetWindowThreadProcessId(int(app.Hwnd), ctypes.byref(final_pid))
            if final_pid.value != owned_pid:
                raise AssertionError("Excel process identity changed during lifecycle test")
            observed = "host deactivation, function unavailable, two successful registrations in same PID"
        elif scenario == "ASYNC-CANCEL":
            original["Visible"] = app.Visible
            original["Calculation"] = app.Calculation
            app.Calculation = -4135
            app.ScreenUpdating = True
            app.Visible = True
            app.Caption = "Excel API cancellation — press F9 then Escape"
            sheet.Range("B1").Value = 0
            async_pending = True
            sheet.Range("C1").Formula = "=RUST.TEST.DELAY(B1)"
            deadline = time.monotonic() + 10
            while sheet.Range("C1").Value != 42 and time.monotonic() < deadline:
                time.sleep(.1)
            if sheet.Range("C1").Value != 42:
                raise RuntimeError("Zero-delay async control did not complete")
            before_status = tuple(int(x) for x in app.Run("RUST.TEST.ASYNC.STATUS").split(","))
            sheet.Range("B1").Value = 10000
            sheet.Range("A3").Value = "Press F9, wait 1 second, then press Escape."
            event(f"Ready for manual F9 then Escape; initial status={before_status!r}")
            import pywintypes
            deadline = time.monotonic() + 60
            observed = before_status
            while time.monotonic() < deadline:
                time.sleep(.25)
                try:
                    observed = tuple(int(x) for x in app.Run("RUST.TEST.ASYNC.STATUS").split(","))
                except pywintypes.com_error as exc:
                    nested = exc.excepinfo[5] if exc.excepinfo else None
                    if exc.hresult in (-2146777998, -2147418111) or nested == -2146777998:
                        continue
                    raise
                if observed[1] > before_status[1] and observed[0] == 0 and observed[3] == 0:
                    break
            event(f"final async status (pending,cancel_events,end_events,active,canceled_workers,completed)={observed!r}")
            retirement = retire_async_probe(app, sheet)
            event(f"fresh status after removing async formula={retirement!r}")
            async_pending = False
            if observed[1] <= before_status[1] or observed[4] <= before_status[4]:
                raise RuntimeError("Native cancellation event and cooperative worker cancellation were not both observed")
            if async_pending:
                raise RuntimeError("Worker retirement was not observed; retaining owned instance")
        elif scenario == "REF-STALE":
            sheet.Range("J1").Value = 123
            doomed = book.Worksheets.Add()
            doomed.Activate()
            app.Run("RUST.TEST.CAPTURE")
            doomed.Delete()
            doomed = None
            sheet.Activate()
            returned = app.Run("RUST.TEST.WRITE.CAPTURED")
            observed = sheet.Range("J1").Value
            if returned != 0 or observed != 123:
                raise AssertionError(f"expected stale rejection and unchanged 123; command={returned!r}, cell={observed!r}")
        elif scenario == "REF-PROTECTED":
            sheet.Range("J1").Value = 123
            sheet.Protect()
            returned = app.Run("RUST.TEST.WRITE")
            observed = sheet.Range("J1").Value
            sheet.Unprotect()
            if returned != 0 or observed != 123:
                raise AssertionError(f"expected rejected command and unchanged 123; command={returned!r}, cell={observed!r}")
        elif scenario == "REF-WRITE":
            app.Run("RUST.TEST.WRITE")
            observed = sheet.Range("J1").Value
            if observed != 42: raise AssertionError(f"write expected 42, observed {observed!r}")
            app.Run("RUST.TEST.CLEAR")
            observed = sheet.Range("J1").Value
            if observed is not None: raise AssertionError(f"clear expected blank, observed {observed!r}")
        elif scenario == "CALC":
            original["Calculation"] = app.Calculation
            app.Calculation = -4135
            sheet.Range("J2").Value = 1
            sheet.Range("J3").Formula = "=J2+1"
            sheet.Range("J3").Calculate()
            sheet.Range("J2").Value = 10
            if sheet.Range("J3").Value != 2: raise RuntimeError("manual calculation control failed")
            app.Run("RUST.TEST.CALCULATE")
            observed = sheet.Range("J3").Value
            if observed != 11: raise AssertionError(f"expected 11, observed {observed!r}")
        elif scenario == "COMMAND":
            observed = app.Run("RUST.PING.COMMAND")
            event(f"command returned {observed!r}")
        elif scenario == "QUEUE":
            app.Run("RUST.DISPATCH.ENQUEUE")
            sheet.Range("A1").Formula = "=RUST.DISPATCH.STATUS()"
            app.Calculate()
            before_status = str(app.Run("RUST.DISPATCH.STATUS"))
            if not before_status.startswith("queued;"): raise AssertionError(f"before pump: {before_status!r}")
            app.Run("RUST.DISPATCH.PUMP")
            sheet.Range("A1").Formula = "=RUST.DISPATCH.STATUS()"
            sheet.Range("A1").Calculate()
            observed = str(app.Run("RUST.DISPATCH.STATUS"))
            if not observed.startswith("completed;"): raise AssertionError(f"after pump: {observed!r}")
        else:
            sheet.Range("J1").Value = 42
            formulas = {"CLEANUP": ("=RUST.ADD(2,3)", 5), "REF-READ": ("=RUST.TEST.READ(J1)", 42), "DOC": ("=RUST.TEST.DOCUMENT()", book.Name), "HELP": ("=RUST.TEST.HELP(17)", 17), "REG-ADD": ("=RUST.ADD(2,3)", 5), "VALUE-TEXT": ('=RUST.ECHO("A\u00e9\u6c34\U0001f600")', "A\u00e9\u6c34\U0001f600"),
                        "VALUE-OPTION": ("=RUST.OPTION.KIND()", "missing"), "REF-KIND": ("=RUST.REFERENCE.KIND(B1)", ("SRef", "Ref")),
                        "ASYNC": ("=RUST.ASYNC.DOUBLE(7)", 14)}
            formula, expected = formulas[scenario]
            if scenario == "ASYNC": async_pending = True
            sheet.Range("A1").Formula = formula
            app.Calculate()
            deadline = time.monotonic() + 15
            while True:
                observed = sheet.Range("A1").Value
                if observed == expected or isinstance(expected, tuple) and observed in expected:
                    async_pending = False
                    break
                if scenario != "ASYNC" or time.monotonic() >= deadline:
                    raise AssertionError(f"expected {expected!r}, observed {observed!r}")
                time.sleep(.1)
        event(f"assertion passed; observed={observed!r}")
        result = {"status": "PASS", "observed": repr(observed)}
    except Exception as exc:
        event(f"error: {type(exc).__name__}: {exc}")
        result = {"status": "FAIL" if isinstance(exc, AssertionError) else "BLOCKED", "observed": str(exc)}
    finally:
        if owned and async_pending:
            event("async completion not observed; workbook and application intentionally left open for safe inspection")
            result["status"] = "BLOCKED"
            result["observed"] += "; possible async work remains, no explicit unload or Quit attempted"
        elif owned:
            cleanup_errors = []
            for name, value in original.items():
                try:
                    setattr(app, name, value)
                except Exception as exc:
                    cleanup_errors.append(f"restore {name}: {exc}")
            try:
                if book is not None: book.Close(SaveChanges=False)
            except Exception as exc:
                cleanup_errors.append(f"scratch workbook close: {exc}")
            try:
                app.Quit()
                event("owned application Quit requested")
                sheet = book = app = None
                import gc
                gc.collect()
            except Exception as exc:
                cleanup_errors.append(f"Quit: {exc}")
            if cleanup_errors:
                event("cleanup incomplete: " + "; ".join(cleanup_errors))
                if result["status"] == "PASS": result["status"] = "BLOCKED"
                result["observed"] += "; cleanup incomplete: " + "; ".join(cleanup_errors)
    result["owned_pid"] = owned_pid
    return result


def isolated_pid(before, after, pid):
    return bool(pid and pid not in before and after - before == {pid})


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check-environment", action="store_true")
    parser.add_argument("--xll")
    parser.add_argument("--output", default="native-excel-results")
    parser.add_argument("--scenario", action="append", help="Scenario ID or area; repeatable")
    parser.add_argument("--timeout", type=int, default=90, help="Per-scenario worker deadline in seconds")
    parser.add_argument("--manual-ui", action="store_true", help="Operator is present for F9/Escape scenarios; use --timeout 130")
    parser.add_argument("--allow-stale-probe", action="store_true", help="Explicitly enable the known-crashing stale-sheet investigation; isolated diagnostic hosts only")
    parser.add_argument("--worker", choices=SCENARIOS, help=argparse.SUPPRESS)
    parser.add_argument("--event-file", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker:
        print(json.dumps(worker(args.worker, args.xll, args.event_file, args.manual_ui)), flush=True)
        return 0
    if args.timeout < 1 or args.timeout > 600: parser.error("--timeout must be between 1 and 600")
    selected = list(SCENARIOS)
    if args.scenario:
        keys = {s.lower() for s in args.scenario}
        valid = {x.lower() for x in SCENARIOS} | {x[0] for x in SCENARIOS.values()}
        if keys - valid: parser.error("unknown scenarios: " + ", ".join(sorted(keys - valid)))
        selected = [k for k, v in SCENARIOS.items() if k.lower() in keys or v[0] in keys]
    env = environment(args.xll)
    errors = validate_environment(env, need_xll=not args.check_environment)
    run_id = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ-") + uuid.uuid4().hex[:8]
    output = Path(args.output) / run_id
    output.mkdir(parents=True, exist_ok=False)
    report = {"run_id": run_id, "environment": env, "preflight_errors": errors, "scenarios": []}
    if not args.check_environment:
        for sid in selected:
            area, expected, implemented = SCENARIOS[sid]
            entry = {"id": sid, "area": area, "expected": expected, "status": "NOT RUN", "observed": "fixture not implemented", "elapsed_seconds": 0}
            if sid == "REF-STALE" and not args.allow_stale_probe:
                entry.update(status="BLOCKED", observed="Prior stale-sheet probe lost Excel RPC; diagnostic opt-in required")
            elif sid in ("EVENTS", "ASYNC-CANCEL") and not args.manual_ui:
                entry.update(status="BLOCKED", observed="Manual F9/Escape requires an operator and --manual-ui")
            elif sid == "WAKE":
                entry.update(status="BLOCKED", observed="autonomous wake is not implemented (#30/#101)")
            elif implemented and errors:
                entry.update(status="BLOCKED", observed="; ".join(errors))
            elif implemented:
                start = time.monotonic()
                log = output / (sid + ".log")
                command = [sys.executable, __file__, "--worker", sid, "--xll", args.xll, "--event-file", str(log)]
                if args.manual_ui:
                    command.append("--manual-ui")
                try:
                    p = subprocess.run(command, capture_output=True, text=True, timeout=args.timeout)
                    entry.update(json.loads(p.stdout.strip().splitlines()[-1]))
                    # Excel can take a short interval to exit after its last COM
                    # proxy is released. Observe only; never kill a process.
                    deadline = time.monotonic() + 15
                    owned_pid = entry.get("owned_pid")
                    while owned_pid and owned_pid in process_ids() and time.monotonic() < deadline:
                        time.sleep(.25)
                    if owned_pid and owned_pid in process_ids() and entry["status"] == "PASS":
                        entry.update(status="BLOCKED", observed=entry["observed"] + "; Excel process exit not observed")
                except subprocess.TimeoutExpired:
                    entry.update(status="BLOCKED", observed="worker timed out; Excel process ownership may be unknown; no Excel process was terminated")
                except (ValueError, IndexError, OSError, subprocess.SubprocessError) as exc:
                    entry.update(status="BLOCKED", observed="worker failed: " + str(exc))
                entry["elapsed_seconds"] = round(time.monotonic() - start, 3)
                entry["log"] = log.name
            report["scenarios"].append(entry)
    (output / "report.json").write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    lines = [f"Run {run_id}", "Preflight: " + ("; ".join(errors) if errors else "OK")]
    lines += [f"{x['id']}: {x['status']} - {x['observed']}" for x in report["scenarios"]]
    (output / "summary.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))
    print("Evidence:", output)
    return 1 if errors or any(x["status"] in ("FAIL", "BLOCKED") for x in report["scenarios"]) else 0


if __name__ == "__main__":
    sys.exit(main())
