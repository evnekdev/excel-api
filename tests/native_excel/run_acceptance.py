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
    "REF-READ": ("references", "xlCoerce scalar, array and error reads", False),
    "REF-WRITE": ("references", "Command-only xlSet writes and clears", False),
    "DOC": ("document", "Active and named document information", False),
    "CALC": ("calculation", "Command-only normal recalculation", False),
    "HELP": ("help", "HelpTopic and Function Wizard metadata", False),
    "EVENTS": ("events", "Calculation notifications and lifetime", False),
    "CLEANUP": ("lifecycle", "Unregister, retry, unload and reload", False),
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


def worker(scenario, xll, event_file):
    def event(message):
        with open(event_file, "a", encoding="utf-8") as out:
            out.write(message + "\n")
    app = book = None
    owned = False
    async_pending = False
    original = {}
    try:
        before = process_ids()
        if before:
            raise RuntimeError("Existing Excel processes detected; isolation cannot be guaranteed")
        event("activation started; no existing Excel process")
        import win32com.client
        app = win32com.client.DispatchEx("Excel.Application")
        hwnd = int(app.Hwnd)
        pid = ctypes.c_ulong()
        if not ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid)) or not pid.value or pid.value in before or process_ids() != {pid.value}:
            raise RuntimeError("Excel application process ownership could not be established")
        owned = True
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
        if scenario == "COMMAND":
            observed = app.Run("RUST.PING.COMMAND")
            event(f"command returned {observed!r}")
        elif scenario == "QUEUE":
            app.Run("RUST.DISPATCH.ENQUEUE")
            sheet.Range("A1").Formula = "=RUST.DISPATCH.STATUS()"
            app.Calculate()
            before_status = str(sheet.Range("A1").Value)
            if not before_status.startswith("queued;"): raise AssertionError(f"before pump: {before_status!r}")
            app.Run("RUST.DISPATCH.PUMP")
            sheet.Range("A1").Calculate()
            observed = str(sheet.Range("A1").Value)
            if not observed.startswith("completed;"): raise AssertionError(f"after pump: {observed!r}")
        else:
            formulas = {"REG-ADD": ("=RUST.ADD(2,3)", 5), "VALUE-TEXT": ('=RUST.ECHO("Aé水😀")', "Aé水😀"),
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
            try:
                if book is not None: book.Close(SaveChanges=False)
            except Exception as exc:
                cleanup_errors.append(f"scratch workbook close: {exc}")
            for name, value in original.items():
                try:
                    setattr(app, name, value)
                except Exception as exc:
                    cleanup_errors.append(f"restore {name}: {exc}")
            try:
                app.Quit()
                event("owned application Quit requested")
            except Exception as exc:
                cleanup_errors.append(f"Quit: {exc}")
            if cleanup_errors:
                event("cleanup incomplete: " + "; ".join(cleanup_errors))
                if result["status"] == "PASS": result["status"] = "BLOCKED"
                result["observed"] += "; cleanup incomplete: " + "; ".join(cleanup_errors)
    return result


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check-environment", action="store_true")
    parser.add_argument("--xll")
    parser.add_argument("--output", default="native-excel-results")
    parser.add_argument("--scenario", action="append", help="Scenario ID or area; repeatable")
    parser.add_argument("--timeout", type=int, default=90, help="Per-scenario worker deadline in seconds")
    parser.add_argument("--worker", choices=SCENARIOS, help=argparse.SUPPRESS)
    parser.add_argument("--event-file", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker:
        print(json.dumps(worker(args.worker, args.xll, args.event_file)), flush=True)
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
            if sid == "WAKE":
                entry.update(status="BLOCKED", observed="autonomous wake is not implemented (#30/#101)")
            elif implemented and errors:
                entry.update(status="BLOCKED", observed="; ".join(errors))
            elif implemented:
                start = time.monotonic()
                log = output / (sid + ".log")
                command = [sys.executable, __file__, "--worker", sid, "--xll", args.xll, "--event-file", str(log)]
                try:
                    p = subprocess.run(command, capture_output=True, text=True, timeout=args.timeout)
                    entry.update(json.loads(p.stdout.strip().splitlines()[-1]))
                except subprocess.TimeoutExpired:
                    entry.update(status="BLOCKED", observed="worker timed out; Excel process ownership may be unknown; no Excel process was terminated")
                except (ValueError, IndexError, OSError) as exc:
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
