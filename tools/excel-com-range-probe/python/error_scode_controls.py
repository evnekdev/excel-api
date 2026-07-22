"""Opt-in client-visible Prompt 05J controls; no raw VARTYPE inference."""
import argparse
import importlib.metadata
import json
import platform
import sys
import time

CASES = [
    ("null", 2000, "=A1 A2"), ("div0", 2007, "=1/0"),
    ("value", 2015, '=1+"x"'), ("ref", 2023, '=INDIRECT("A0")'),
    ("name", 2029, "=NOT_A_FUNCTION()"), ("num", 2036, "=SQRT(-1)"),
    ("na", 2042, "=NA()"),
]

def hresult(error):
    value = getattr(error, "hresult", None)
    return None if value is None else f"0x{value & 0xffffffff:08X}"

def row(client, case, source, member, action):
    try:
        visible = action()
        return {"schema_version": 1, "id": f"python.{client}.{case[0]}.{source}.{member}", "client": client, "case": case[0], "source_value": source, "member": member, "status": "completed", "visible_result": visible, "physical_vartype": "unknown", "raw_pointer_values_recorded": False}
    except Exception as error:
        return {"schema_version": 1, "id": f"python.{client}.{case[0]}.{source}.{member}", "client": client, "case": case[0], "source_value": source, "member": member, "status": "failed", "hresult": hresult(error), "exception_type": type(error).__name__, "physical_vartype": "unknown", "raw_pointer_values_recorded": False}

def run_pywin32():
    import pythoncom
    from win32com.client import DispatchEx, VARIANT
    app = DispatchEx("Excel.Application")
    app.Visible = False
    book = app.Workbooks.Add()
    sheet = app.ActiveSheet
    rows = []
    try:
        for index, case in enumerate(CASES, 1):
            source_cell = sheet.Range(f"A{index}")
            source_cell.Formula = case[2]
            app.Calculate()
            raw_client_value = source_cell.Value
            for member in ("Value", "Value2"):
                target = sheet.Range("C1")
                rows.append(row("pywin32", case, "formula-returned-client-value", member, lambda target=target, member=member, raw_client_value=raw_client_value: (setattr(target, member, raw_client_value), target.Text)[1]))
                rows.append(row("pywin32", case, "explicit-short-VT_ERROR", member, lambda target=target, member=member, case=case: (setattr(target, member, VARIANT(pythoncom.VT_ERROR, case[1])), target.Text)[1]))
                rows.append(row("pywin32", case, "explicit-full-SCODE-VT_ERROR", member, lambda target=target, member=member, case=case: (setattr(target, member, VARIANT(pythoncom.VT_ERROR, (0x800A0000 | case[1]) - 0x100000000)), target.Text)[1]))
                rows.append({"schema_version":1,"id":f"python.pywin32.{case[0]}.cverr-style.{member}","client":"pywin32","case":case[0],"source_value":"CVErr-style", "member":member,"status":"not-supported","detail":"pywin32 exposes an explicit VT_ERROR constructor, not a VBA CVErr helper","physical_vartype":"unknown","raw_pointer_values_recorded":False})
    finally:
        book.Close(False); app.Quit(); time.sleep(0.25)
    rows.append({"schema_version":1,"id":"python.pywin32.environment","client":"pywin32","python_version":platform.python_version(),"pywin32_version":importlib.metadata.version("pywin32"),"raw_pointer_values_recorded":False})
    return rows

def run_comtypes():
    from comtypes.client import CreateObject
    import comtypes
    app = CreateObject("Excel.Application", dynamic=True)
    app.Visible = False
    book = app.Workbooks.Add()
    sheet = app.ActiveSheet
    rows = []
    try:
        for index, case in enumerate(CASES, 1):
            source_cell = sheet.Range(f"A{index}")
            source_cell.Formula = case[2]
            app.Calculate()
            raw_client_value = source_cell.Value
            for member in ("Value", "Value2"):
                target = sheet.Range("C1")
                rows.append(row("comtypes", case, "formula-returned-client-value", member, lambda target=target, member=member, raw_client_value=raw_client_value: (setattr(target, member, raw_client_value), target.Text)[1]))
                rows.append({"schema_version":1,"id":f"python.comtypes.{case[0]}.explicit-VT_ERROR.{member}","client":"comtypes","case":case[0],"source_value":"explicit VT_ERROR","member":member,"status":"not-supported","detail":"installed comtypes VARIANT union has no exposed SCODE member","physical_vartype":"unknown","raw_pointer_values_recorded":False})
    finally:
        book.Close(False); app.Quit(); time.sleep(0.25)
    rows.append({"schema_version":1,"id":"python.comtypes.environment","client":"comtypes","python_version":platform.python_version(),"comtypes_version":comtypes.__version__,"raw_pointer_values_recorded":False})
    return rows

if __name__ == "__main__":
    parser = argparse.ArgumentParser(); parser.add_argument("--client", choices=["pywin32", "comtypes"], required=True)
    print(json.dumps(run_pywin32() if parser.parse_args().client == "pywin32" else run_comtypes(), separators=(",", ":")))
