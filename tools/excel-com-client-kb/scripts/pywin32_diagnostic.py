"""Opt-in, bounded Excel control for Prompt 05C.

The script emits one JSON object without local paths, raw pointers, or HWND
values.  It creates one unsaved workbook, exercises only non-destructive
members, then requests Excel shutdown in a finally block.
"""

import argparse
import importlib.metadata
import json
import sys


def identity(value):
    return {"class": type(value).__name__, "module": type(value).__module__}


def dynamic_application(pythoncom):
    # Hold the successful, documented DispatchEx activation path constant, then
    # deliberately wrap its IDispatch pointer in the dynamic implementation.
    from win32com import client
    from win32com.client import dynamic

    activated = client.DispatchEx("Excel.Application")
    return dynamic.Dispatch(activated._oleobj_, "Excel.Application")


def generated_application():
    from win32com.client import gencache

    return gencache.EnsureDispatch("Excel.Application")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=("dynamic", "generated"), required=True)
    args = parser.parse_args()

    import pythoncom

    application = None
    workbook = None
    workbooks = None
    result = {
        "schema_version": 1,
        "client": "pywin32",
        "mode": args.mode,
        "python_version": sys.version.split()[0],
        "pywin32_version": importlib.metadata.version("pywin32"),
        "com_initialization": {
            "sys_coinitialization_flags": getattr(sys, "coinit_flags", None),
            "module_initialization": "pythoncom import initializes the main thread",
        },
    }
    try:
        application = dynamic_application(pythoncom) if args.mode == "dynamic" else generated_application()
        workbooks = application.Workbooks
        result.update({"application": identity(application), "workbooks": identity(workbooks)})
        workbook = workbooks.Add()
        application.Visible = False
        worksheets = workbook.Worksheets
        worksheet = worksheets.Item(1)
        cell = worksheet.Range("A1")
        cell.Value2 = 7
        array_value = worksheet.Range("A1:B2").Value2
        result.update(
            {
                "status": "Control-confirmed",
                "excel_version": application.Version,
                "workbook": identity(workbook),
                "worksheets": identity(worksheets),
                "worksheet": identity(worksheet),
                "range_value2_scalar_type": type(cell.Value2).__name__,
                "range_value2_rectangular_shape": [len(array_value), len(array_value[0])],
                "workbooks_add_created_name": workbook.Name,
            }
        )
    except Exception as error:  # Record the class/message without traceback paths.
        result.update({"status": "Inconclusive", "error_type": type(error).__name__, "error": str(error)})
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
