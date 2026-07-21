"""Opt-in availability/control diagnostic for Prompt 05C."""

import importlib.util
import json
import sys
import argparse


def identity(value):
    return {"class": type(value).__name__, "module": type(value).__module__}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=("dynamic", "generated"), default="dynamic")
    args = parser.parse_args()
    if importlib.util.find_spec("comtypes") is None:
        print(json.dumps({"schema_version": 1, "client": "comtypes", "mode": args.mode, "status": "Not tested", "reason": "package not installed", "python_version": sys.version.split()[0]}, sort_keys=True, separators=(",", ":")))
        return

    import comtypes
    import comtypes.client

    application = None
    workbook = None
    result = {"schema_version": 1, "client": "comtypes", "mode": args.mode, "comtypes_version": comtypes.__version__, "status": "Inconclusive", "python_version": sys.version.split()[0], "com_initialization": {"sys_coinitialization_flags": getattr(sys, "coinit_flags", None), "module_initialization": "comtypes import initializes the main thread"}}
    try:
        if args.mode == "generated":
            comtypes.client.GetModule(("{00020813-0000-0000-C000-000000000046}", 1, 9))
            application = comtypes.client.CreateObject("Excel.Application")
        else:
            application = comtypes.client.CreateObject("Excel.Application", dynamic=True)
        workbooks = application.Workbooks
        result.update({"application": identity(application), "workbooks": identity(workbooks)})
        workbook = workbooks.Add()
        application.Visible = False
        result.update({"status": "Control-confirmed", "excel_version": application.Version, "workbook": identity(workbook), "workbooks_add_created_name": workbook.Name})
    except Exception as error:
        result.update({"error_type": type(error).__name__, "error": str(error)})
    finally:
        try:
            if workbook is not None:
                workbook.Close(False)
        finally:
            if application is not None:
                application.Quit()
        print(json.dumps(result, sort_keys=True, separators=(",", ":")))


if __name__ == "__main__":
    main()
