"""Opt-in, fresh-process Excel controls for Prompt 05D.

The harness never emits raw HWND values, local paths, or COM pointer values.
It supports only a bounded create/open fixture sequence and always requests
workbook close plus Excel shutdown before returning one JSON record.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gc
import importlib.metadata
import json
import platform
import sys
from ctypes import Structure, byref, windll, wintypes


PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
SYNCHRONIZE = 0x00100000
WAIT_OBJECT_0 = 0


class FILETIME(Structure):
    _fields_ = (("dwLowDateTime", wintypes.DWORD), ("dwHighDateTime", wintypes.DWORD))


def identity(value):
    return {"class": type(value).__name__, "module": type(value).__module__}


def hresult(error):
    value = getattr(error, "hresult", None)
    if value is None and getattr(error, "args", None):
        candidate = error.args[0]
        if isinstance(candidate, int):
            value = candidate
    return None if not isinstance(value, int) else f"0x{value & 0xFFFFFFFF:08X}"


def safe_property(target, name, redact=False):
    try:
        value = getattr(target, name)
    except Exception as error:  # The error string can disclose local state.
        return {"status": "Not tested", "error_type": type(error).__name__, "hresult": hresult(error)}
    if redact:
        return {"status": "available", "value_recorded": False}
    if isinstance(value, (bool, int, float, str)) or value is None:
        return {"status": "available", "value": value}
    return {"status": "available", "value_type": type(value).__name__}


def filetime_to_utc(value):
    ticks = (value.dwHighDateTime << 32) | value.dwLowDateTime
    return (dt.datetime(1601, 1, 1, tzinfo=dt.timezone.utc) + dt.timedelta(microseconds=ticks // 10)).isoformat()


def process_identity(hwnd):
    process_id = wintypes.DWORD()
    windll.user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), byref(process_id))
    if process_id.value == 0:
        return {"status": "Not tested", "reason": "Hwnd-to-PID lookup returned zero"}
    handle = windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE, False, process_id.value)
    if not handle:
        return {"status": "available", "pid": process_id.value, "start_time_utc": "Not tested"}
    creation, exit_time, kernel, user = FILETIME(), FILETIME(), FILETIME(), FILETIME()
    try:
        if windll.kernel32.GetProcessTimes(handle, byref(creation), byref(exit_time), byref(kernel), byref(user)):
            return {"status": "available", "pid": process_id.value, "start_time_utc": filetime_to_utc(creation)}
        return {"status": "available", "pid": process_id.value, "start_time_utc": "Not tested"}
    finally:
        windll.kernel32.CloseHandle(handle)


def process_exited(identity_record):
    process_id = identity_record.get("pid")
    if not isinstance(process_id, int):
        return {"status": "Not tested", "reason": "PID unavailable"}
    handle = windll.kernel32.OpenProcess(SYNCHRONIZE, False, process_id)
    if not handle:
        return {"status": "exited", "pid": process_id}
    try:
        wait = windll.kernel32.WaitForSingleObject(handle, 5000)
        return {"status": "exited" if wait == WAIT_OBJECT_0 else "not-exited-within-5000ms", "pid": process_id}
    finally:
        windll.kernel32.CloseHandle(handle)


def pywin32_application(mode):
    from win32com import client

    if mode == "dynamic":
        from win32com.client import dynamic

        activated = client.DispatchEx("Excel.Application")
        return dynamic.Dispatch(activated._oleobj_, "Excel.Application"), {
            "api": "DispatchEx then dynamic.Dispatch",
            "clsctx": "CLSCTX_SERVER",
            "requested_iid": "IID_IDispatch",
            "type_info_probing": "dynamic Dispatch wrapper selection",
        }
    from win32com.client import gencache

    return gencache.EnsureDispatch("Excel.Application"), {
        "api": "gencache.EnsureDispatch",
        "clsctx": "client-managed activation",
        "requested_iid": "IID_IDispatch then generated wrapper selection",
        "type_info_probing": "generated cache; cache path is isolated per process and not recorded",
    }


def comtypes_application(mode):
    import comtypes.client

    if mode == "dynamic":
        return comtypes.client.CreateObject("Excel.Application", dynamic=True), {
            "api": "comtypes.client.CreateObject(dynamic=True)",
            "clsctx": "CLSCTX_SERVER default",
            "requested_iid": "IDispatch",
            "type_info_probing": "lazybind when type information is available",
        }
    comtypes.client.GetModule(("{00020813-0000-0000-C000-000000000046}", 1, 9))
    return comtypes.client.CreateObject("Excel.Application"), {
        "api": "GetModule then comtypes.client.CreateObject",
        "clsctx": "CLSCTX_SERVER default",
        "requested_iid": "generated coclass default interface",
        "type_info_probing": "generated dual-interface or dispinterface selection",
    }


def application_state(application, workbooks):
    hwnd = safe_property(application, "Hwnd")
    identity_record = process_identity(hwnd.get("value")) if hwnd.get("status") == "available" else {"status": "Not tested"}
    return {
        "application_version": safe_property(application, "Version"),
        "hwnd": {"status": hwnd.get("status"), "value_recorded": False},
        "owned_process": identity_record,
        "visible": safe_property(application, "Visible"),
        "user_control": safe_property(application, "UserControl"),
        "interactive": safe_property(application, "Interactive"),
        "ready": safe_property(application, "Ready"),
        "workbooks_count": safe_property(workbooks, "Count"),
        "startup_path": safe_property(application, "StartupPath", redact=True),
        "default_file_path": safe_property(application, "DefaultFilePath", redact=True),
        "calculation": safe_property(application, "Calculation"),
        "automation_security": safe_property(application, "AutomationSecurity"),
        "display_alerts": safe_property(application, "DisplayAlerts"),
        "modal_or_error_state": "not directly detectable through the bounded Automation surface",
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--client", choices=("pywin32", "comtypes"), required=True)
    parser.add_argument("--mode", choices=("dynamic", "generated"), required=True)
    parser.add_argument("--operation", choices=("add", "open", "create-fixture"), default="add")
    parser.add_argument("--workbook", help="temporary fixture path; never serialized")
    args = parser.parse_args()
    if args.operation in ("open", "create-fixture") and not args.workbook:
        parser.error("--workbook is required for open and create-fixture")

    application = workbooks = workbook = None
    result = {
        "schema_version": 1,
        "classification": "Inconclusive",
        "client": args.client,
        "mode": args.mode,
        "operation": args.operation,
        "python_version": sys.version.split()[0],
        "interpreter_architecture": platform.architecture()[0],
        "package_architecture": "win_amd64" if args.client == "pywin32" else "pure Python package on the interpreter architecture",
        "com_initialization": {"sys_coinitialization_flags": getattr(sys, "coinit_flags", None), "main_thread": "client import initialization"},
        "raw_hwnd_recorded": False,
        "raw_pointer_values_recorded": False,
        "raw_paths_recorded": False,
    }
    try:
        if args.client == "pywin32":
            result["package_version"] = importlib.metadata.version("pywin32")
            application, activation = pywin32_application(args.mode)
        else:
            result["package_version"] = importlib.metadata.version("comtypes")
            application, activation = comtypes_application(args.mode)
        workbooks = application.Workbooks
        result["activation"] = activation
        result["wrapper_classes"] = {"Application": identity(application), "Workbooks": identity(workbooks)}
        result["session_state_before_operation"] = application_state(application, workbooks)
        if args.operation == "add":
            workbook = workbooks.Add()
        elif args.operation == "open":
            workbook = workbooks.Open(args.workbook)
        else:
            workbook = workbooks.Add()
            workbook.SaveAs(args.workbook)
        result.update({
            "classification": "Control-confirmed",
            "workbook": identity(workbook),
            "workbook_name": workbook.Name,
            "workbook_access": "succeeded",
        })
    except Exception as error:
        result.update({"error_type": type(error).__name__, "hresult": hresult(error), "workbook_access": "failed"})
    finally:
        try:
            if workbook is not None:
                workbook.Close(False)
        except Exception as error:
            result["workbook_close_error"] = {"error_type": type(error).__name__, "hresult": hresult(error)}
        try:
            if application is not None:
                application.Quit()
        except Exception as error:
            result["excel_quit_error"] = {"error_type": type(error).__name__, "hresult": hresult(error)}
        identity_record = result.get("session_state_before_operation", {}).get("owned_process", {})
        workbook = workbooks = application = None
        gc.collect()
        result["cleanup"] = {"owned_process_exit": process_exited(identity_record), "forced_termination": False}
    print(json.dumps(result, sort_keys=True, separators=(",", ":")))


if __name__ == "__main__":
    main()
