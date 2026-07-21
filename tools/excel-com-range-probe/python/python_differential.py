#!/usr/bin/env python3
"""Bounded, pointer-free Python client observations for Prompt 05I.

The script intentionally records Python-visible conversion results separately
from the raw ``windows-sys`` VARIANT evidence.  It never treats a Python
``datetime``, ``float``, tuple, or ``None`` read as proof of a physical
VARTYPE.  Every live invocation refuses a pre-existing Excel process, creates
one workbook in one new Excel instance, and verifies that the count returns to
zero after requesting ``Quit``.  It never terminates a process.
"""

from __future__ import annotations

import argparse
import datetime as datetime_module
import decimal
import gc
import json
import os
import platform
import subprocess
import sys
import time
import traceback
from typing import Any, Callable


def excel_process_count() -> int:
    completed = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-Command",
            "@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count",
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    return int(completed.stdout.strip())


def wait_for_zero_excel_processes(timeout_seconds: float = 15.0) -> int:
    """Allow COM wrappers to release naturally; never terminate a process."""
    deadline = time.monotonic() + timeout_seconds
    remaining = excel_process_count()
    while remaining != 0 and time.monotonic() < deadline:
        time.sleep(0.25)
        remaining = excel_process_count()
    return remaining


def jsonable(value: Any, depth: int = 0) -> Any:
    """Return a bounded client-visible description without pointer values."""
    if depth > 3:
        return {"kind": "truncated"}
    if value is None:
        return {"kind": "none"}
    if isinstance(value, bool):
        return {"kind": "bool", "value": value}
    if isinstance(value, int):
        return {"kind": "int", "value": value}
    if isinstance(value, float):
        return {"kind": "float", "value": value, "negative_zero": value == 0.0 and str(value).startswith("-")}
    if isinstance(value, str):
        return {
            "kind": "str",
            "utf16_length": len(value.encode("utf-16-le")) // 2,
            "preview": value if len(value) <= 80 else None,
        }
    if isinstance(value, decimal.Decimal):
        return {"kind": "decimal", "value": str(value)}
    if isinstance(value, (datetime_module.datetime, datetime_module.date)):
        return {"kind": type(value).__name__, "iso8601": value.isoformat()}
    if isinstance(value, (tuple, list)):
        items = [jsonable(item, depth + 1) for item in value[:12]]
        return {"kind": type(value).__name__, "length": len(value), "items": items}
    return {"kind": type(value).__name__, "repr": repr(value)[:160]}


def error_record(error: BaseException) -> dict[str, Any]:
    hresult = getattr(error, "hresult", None)
    if hresult is None and getattr(error, "args", None):
        first = error.args[0]
        if isinstance(first, int):
            hresult = first
    return {
        "exception_type": type(error).__name__,
        "hresult": hresult,
        "hresult_hex": f"0x{hresult & 0xFFFFFFFF:08X}" if isinstance(hresult, int) else None,
        "message": str(error)[:500],
    }


def attempt(label: str, operation: Callable[[], Any]) -> dict[str, Any]:
    try:
        return {"label": label, "status": "completed", "value": jsonable(operation())}
    except Exception as error:  # Client failures are evidence, not fatal to the run.
        return {"label": label, "status": "failed", "error": error_record(error)}


class Client:
    def __init__(self, client_name: str, wrapper: str, cache_dir: str):
        self.client_name = client_name
        self.wrapper = wrapper
        self.cache_dir = cache_dir
        self.app: Any = None
        self.input_metadata: Callable[[Any], dict[str, Any]]
        self.wrapper_state: dict[str, Any] = {}

    def create(self) -> None:
        os.makedirs(self.cache_dir, exist_ok=True)
        if self.client_name == "pywin32":
            import pythoncom
            import win32com
            import win32com.client

            def pywin32_metadata(value: Any) -> dict[str, Any]:
                explicit = getattr(value, "varianttype", None)
                return {
                    "source_language_type": type(value).__name__,
                    "explicit_input_vartype": explicit,
                    "raw_return_vartype_observable": False,
                }

            self.input_metadata = pywin32_metadata
            if self.wrapper == "generated":
                # gencache consults this module-level directory; setting it
                # before EnsureDispatch keeps generated wrappers per run.
                win32com.__gen_path__ = self.cache_dir
                import win32com.gen_py

                win32com.gen_py.__path__ = [self.cache_dir]
                self.app = win32com.client.gencache.EnsureDispatch("Excel.Application")
                self.wrapper_state = {
                    "wrapper": "generated",
                    "generation_cache": "isolated-per-run",
                    "wrapper_class": type(self.app).__name__,
                }
            else:
                self.app = win32com.client.DispatchEx("Excel.Application")
                self.wrapper_state = {
                    "wrapper": "dynamic",
                    "generation_cache": "not-used",
                    "wrapper_class": type(self.app).__name__,
                }
            self.pythoncom = pythoncom
            self.win32com_client = win32com.client
            return

        if self.client_name == "comtypes":
            import comtypes
            import comtypes.automation
            import comtypes.client

            def comtypes_metadata(value: Any) -> dict[str, Any]:
                explicit = getattr(value, "vt", None)
                return {
                    "source_language_type": type(value).__name__,
                    "explicit_input_vartype": explicit,
                    "raw_return_vartype_observable": False,
                }

            self.input_metadata = comtypes_metadata
            comtypes.client.gen_dir = self.cache_dir
            if self.wrapper == "generated":
                comtypes.client.GetModule(("{00020813-0000-0000-C000-000000000046}", 1, 9))
                self.app = comtypes.client.CreateObject("Excel.Application")
                self.wrapper_state = {
                    "wrapper": "generated",
                    "generation_cache": "isolated-per-run",
                    "wrapper_class": type(self.app).__name__,
                }
            else:
                self.app = comtypes.client.CreateObject("Excel.Application", dynamic=True)
                self.wrapper_state = {
                    "wrapper": "dynamic",
                    "generation_cache": "not-used",
                    "wrapper_class": type(self.app).__name__,
                }
            self.comtypes = comtypes
            self.comtypes_automation = comtypes.automation
            return

        raise ValueError(f"unsupported client {self.client_name!r}")

    def explicit(self, kind: str, value: Any) -> Any:
        """Return a client wrapper with its requested input VARTYPE visible."""
        if self.client_name == "pywin32":
            tags = {
                "empty": self.pythoncom.VT_EMPTY,
                "null": self.pythoncom.VT_NULL,
                "i4": self.pythoncom.VT_I4,
                "error": self.pythoncom.VT_ERROR,
                "date": self.pythoncom.VT_DATE,
                "currency": self.pythoncom.VT_CY,
            }
            return self.win32com_client.VARIANT(tags[kind], value)
        if kind == "empty":
            # comtypes exposes automatic VT_NULL for None.  It has no
            # portable public constructor for a by-value VT_EMPTY payload.
            return None
        return self.comtypes_automation.VARIANT(value)

    def release(self) -> None:
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
        self.app = None
        gc.collect()


def read_members(target: Any) -> dict[str, Any]:
    return {
        member: attempt(member, lambda member=member: getattr(target, member))
        for member in ("Value", "Value2", "Formula", "Formula2")
    }


def write_and_read(target: Any, member: str, value: Any, client: Client) -> dict[str, Any]:
    clear_before = attempt("ClearContents-before", target.ClearContents)
    write = attempt(f"{member}-write", lambda: setattr(target, member, value))
    return {
        "member": member,
        "source_value": jsonable(value),
        "source_value_metadata": client.input_metadata(value),
        "clear_before": clear_before,
        "write": write,
        "reads": read_members(target),
        "clear_after": attempt("ClearContents-after", target.ClearContents),
        "raw_return_vartype_observable": False,
    }


def mixed_cases(sheet: Any, client: Client) -> list[dict[str, Any]]:
    base = [[1.25, "text", True], [42, 2.5, "more"], [3.5, "tail", False]]
    candidates: list[tuple[str, Any]] = [
        ("M-good-heterogeneous", base),
        ("M-add-empty", client.explicit("empty", None)),
        ("M-add-null", client.explicit("null", None)),
        ("M-add-i4", client.explicit("i4", 42)),
        ("M-add-error", client.explicit("error", 2042)),
        ("M-add-date", client.explicit("date", datetime_module.datetime(2024, 1, 2, 12, 0, 0))),
        ("M-add-currency", client.explicit("currency", decimal.Decimal("123.4500"))),
    ]
    records: list[dict[str, Any]] = []
    target = sheet.Range("A1:C3")
    first_failure: tuple[str, Any] | None = None
    for case_id, replacement in candidates:
        values = [row[:] for row in base]
        if case_id != "M-good-heterogeneous":
            values[2][1] = replacement
        record = write_and_read(target, "Value2", tuple(tuple(row) for row in values), client)
        record.update({"id": case_id, "family": "mixed-array", "attempt": 1})
        records.append(record)
        if first_failure is None and record["write"]["status"] == "failed":
            first_failure = (case_id, replacement)
    if first_failure is not None:
        case_id, replacement = first_failure
        for retry in range(1, 4):
            values = [row[:] for row in base]
            values[2][1] = replacement
            record = write_and_read(target, "Value2", tuple(tuple(row) for row in values), client)
            record.update({"id": f"{case_id}-rerun-{retry}", "family": "mixed-array", "attempt": retry + 1, "reproduces": case_id})
            records.append(record)
    return records


def date_cases(sheet: Any, client: Client) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    cell = sheet.Range("A1")
    for serial in (-1.0, 0.0, 1.0, 0.5, 45292.0, -0.5):
        for format_name, number_format in (("general", "General"), ("date", "m/d/yyyy h:mm")):
            attempt("NumberFormat", lambda: setattr(cell, "NumberFormat", number_format))
            date_value = datetime_module.datetime(1899, 12, 30) + datetime_module.timedelta(days=serial)
            record = write_and_read(cell, "Value", client.explicit("date", date_value), client)
            record.update({"id": f"D-date-{serial:g}-{format_name}", "family": "date", "serial": serial, "number_format": number_format, "input_form": "explicit-or-client-date"})
            records.append(record)
            record = write_and_read(cell, "Value2", serial, client)
            record.update({"id": f"D-r8-{serial:g}-{format_name}", "family": "date", "serial": serial, "number_format": number_format, "input_form": "python-float"})
            records.append(record)
    return records


def shape_cases(sheet: Any, client: Client) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    cases = [
        ("SH-1x2-to-1x3", "A1:C1", ((1, 2),)),
        ("SH-1x3-to-1x2", "A1:B1", ((1, 2, 3),)),
        ("SH-2x2-to-2x3", "A1:C2", ((1, 2), (3, 4))),
        ("SH-2x3-to-2x2", "A1:B2", ((1, 2, 3), (4, 5, 6))),
        ("SH-2d-1x1", "A1", ((1,),)),
        ("SH-rank1-row", "A1:C1", (1, 2, 3)),
        ("SH-rank1-column", "A1:A3", (1, 2, 3)),
        ("SH-rank3", "A1", (((1,),),)),
    ]
    for case_id, address, value in cases:
        record = write_and_read(sheet.Range(address), "Value2", value, client)
        record.update({"id": case_id, "family": "shape-mismatch", "target_range": address})
        records.append(record)
    return records


def dynamic_cases(sheet: Any, client: Client) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for case_id, formula, blocked in [
        ("DA-sequence", "=SEQUENCE(2,3)", False),
        ("DA-text", '=SEQUENCE(2,3)&"x"', False),
        ("DA-blocked", "=SEQUENCE(2,3)", True),
    ]:
        spill = sheet.Range("A1:C2")
        attempt("ClearContents-before", spill.ClearContents)
        if blocked:
            attempt("blocker", lambda: setattr(sheet.Range("B1"), "Value2", "blocked"))
        write = attempt("Formula2-write", lambda: setattr(sheet.Range("A1"), "Formula2", formula))
        record = {
            "id": case_id,
            "family": "dynamic-array",
            "formula2": formula,
            "blocked": blocked,
            "write": write,
            "owner_reads": read_members(sheet.Range("A1")),
            "spill_value2": attempt("spill-Value2", lambda: spill.Value2),
            "clear_after": attempt("ClearContents-after", spill.ClearContents),
            "raw_return_vartype_observable": False,
        }
        records.append(record)
    return records


def execute(args: argparse.Namespace) -> dict[str, Any]:
    preexisting = excel_process_count()
    if preexisting != 0:
        raise RuntimeError(f"safety gate: expected zero pre-existing EXCEL.EXE processes, found {preexisting}")
    started = time.time()
    client = Client(args.client, args.wrapper, args.cache_dir)
    app = None
    book = None
    sheet = None
    records: list[dict[str, Any]] = []
    environment: dict[str, Any] = {
        "schema_version": 1,
        "environment_id": args.environment_id,
        "python_version": sys.version.split()[0],
        "python_architecture": platform.architecture()[0],
        "pointer_bits": 8 * (8 if sys.maxsize > 2**32 else 4),
        "client": args.client,
        "wrapper": args.wrapper,
        "generation_cache": "isolated-per-run" if args.wrapper == "generated" else "not-used",
        "pre_existing_excel_process_count": preexisting,
    }
    setup_error: dict[str, Any] | None = None
    try:
        client.create()
        app = client.app
        environment["wrapper_state"] = client.wrapper_state
        environment["office_version"] = str(app.Version)
        attempt("Visible", lambda: setattr(app, "Visible", False))
        attempt("DisplayAlerts", lambda: setattr(app, "DisplayAlerts", False))
        book = app.Workbooks.Add()
        sheet = app.ActiveSheet
        records.append({"id": "SMOKE-Value2", "family": "smoke", **write_and_read(sheet.Range("A1"), "Value2", 42, client)})
        if args.family in ("all", "mixed"):
            records.extend(mixed_cases(sheet, client))
        if args.family in ("all", "date"):
            records.extend(date_cases(sheet, client))
        if args.family in ("all", "shape"):
            records.extend(shape_cases(sheet, client))
        if args.family in ("all", "dynamic"):
            records.extend(dynamic_cases(sheet, client))
    except Exception as error:
        setup_error = error_record(error)
    finally:
        if book is not None:
            try:
                book.Close(False)
            except Exception:
                pass
        # COM wrapper references keep an out-of-process server alive.  Release
        # the worksheet and workbook before Quit, then wait only for natural
        # exit; this harness never uses process-name or forced termination.
        sheet = None
        book = None
        client.release()
        app = None
        gc.collect()
    remaining = wait_for_zero_excel_processes()
    return {
        "schema_version": 1,
        "environment": environment,
        "setup_error": setup_error,
        "records": records,
        "cleanup": {
            "owned_excel_quit_requested": True,
            "post_test_excel_process_count": remaining,
            "owned_process_exit_verified": remaining == 0,
            "forced_termination": False,
            "elapsed_seconds": round(time.time() - started, 3),
        },
        "success": setup_error is None and remaining == 0,
        "raw_pointer_values_recorded": False,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--client", choices=("pywin32", "comtypes"), required=True)
    parser.add_argument("--wrapper", choices=("dynamic", "generated"), required=True)
    parser.add_argument("--family", choices=("all", "mixed", "date", "shape", "dynamic"), default="all")
    parser.add_argument("--environment-id", required=True)
    parser.add_argument("--cache-dir", required=True)
    return parser.parse_args()


if __name__ == "__main__":
    try:
        print(json.dumps(execute(parse_args()), sort_keys=True, separators=(",", ":")))
    except Exception as error:
        print(json.dumps({"fatal_error": error_record(error), "traceback": traceback.format_exc(limit=3)}, sort_keys=True, separators=(",", ":")))
        raise
