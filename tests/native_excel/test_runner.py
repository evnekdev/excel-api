import contextlib
import io
import json
import os
from pathlib import Path
import struct
import sys
import tempfile
import subprocess
import types
import unittest
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import run_acceptance as runner


def pe(path, machine):
    data = bytearray(128)
    data[:2] = b"MZ"
    struct.pack_into("<I", data, 0x3c, 64)
    data[64:68] = b"PE\0\0"
    struct.pack_into("<H", data, 68, machine)
    path.write_bytes(data)


class RunnerTests(unittest.TestCase):
    def test_pe_arch_and_invalid(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "image"
            for machine, expected in [(0x14c, "x86"), (0x8664, "x64")]:
                pe(p, machine)
                self.assertEqual(runner.pe_arch(p), expected)
            p.write_bytes(b"MZ")
            with self.assertRaises(ValueError): runner.pe_arch(p)

    def test_mismatch_and_missing_dependencies(self):
        env = {"pywin32": False, "excel_executable": "excel.exe", "excel_arch": "x64", "xll": {"architecture": "x86"}}
        with mock.patch.object(runner.os, "name", "nt"):
            errors = runner.validate_environment(env)
        self.assertTrue(any("pywin32" in e for e in errors))
        self.assertTrue(any("does not match" in e for e in errors))

    def test_missing_xll(self):
        env = {"pywin32": True, "excel_executable": "excel.exe", "excel_arch": "x64", "xll": None}
        with mock.patch.object(runner.os, "name", "nt"):
            self.assertTrue(any("valid XLL" in e for e in runner.validate_environment(env)))

    def test_conflicting_process_blocks_before_com(self):
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(runner, "process_ids", return_value={42}):
            result = runner.worker("REG-ADD", "dummy", str(Path(tmp) / "event.log"))
        self.assertEqual(result["status"], "BLOCKED")

    def test_process_identity_mismatch_never_quits(self):
        app = mock.Mock()
        app.Hwnd = 123
        client = types.ModuleType("win32com.client")
        client.DispatchEx = mock.Mock(return_value=app)
        package = types.ModuleType("win32com")
        package.client = client
        user32 = mock.Mock()
        user32.GetWindowThreadProcessId.return_value = 0
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(runner, "process_ids", return_value=set()), \
             mock.patch.dict(sys.modules, {"win32com": package, "win32com.client": client}), \
             mock.patch.object(runner.ctypes, "windll", types.SimpleNamespace(user32=user32), create=True):
            result = runner.worker("REG-ADD", "dummy", str(Path(tmp) / "event.log"))
        self.assertEqual(result["status"], "BLOCKED")
        app.Quit.assert_not_called()

    def test_partial_setup_cleans_owned_session(self):
        app = mock.Mock()
        app.Hwnd = 123
        app.DisplayAlerts = True
        app.ScreenUpdating = True
        app.RegisterXLL.return_value = False
        book = app.Workbooks.Add.return_value
        book.Worksheets.return_value.Range.return_value.Value = 5
        client = types.ModuleType("win32com.client")
        client.DispatchEx = mock.Mock(return_value=app)
        package = types.ModuleType("win32com")
        package.client = client
        user32 = mock.Mock()
        def window_pid(hwnd, pointer):
            pointer._obj.value = 77
            return 1
        user32.GetWindowThreadProcessId.side_effect = window_pid
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(runner, "process_ids", side_effect=[set(), {77}]), \
             mock.patch.dict(sys.modules, {"win32com": package, "win32com.client": client}), \
             mock.patch.object(runner.ctypes, "windll", types.SimpleNamespace(user32=user32), create=True):
            result = runner.worker("REG-ADD", "dummy", str(Path(tmp) / "event.log"))
        self.assertEqual(result["status"], "FAIL")
        book.Close.assert_called_once_with(SaveChanges=False)
        app.Quit.assert_called_once()
        self.assertTrue(app.DisplayAlerts)
        self.assertTrue(app.ScreenUpdating)

    def test_malformed_output_path_fails_without_overwriting(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "file"
            output.write_text("preserve", encoding="utf-8")
            with mock.patch.object(sys, "argv", ["runner", "--check-environment", "--output", str(output)]), \
                 mock.patch.object(runner, "environment", return_value={"pywin32": False, "excel_executable": None, "excel_arch": None, "xll": None}):
                with self.assertRaises(OSError): runner.main()
            self.assertEqual(output.read_text(encoding="utf-8"), "preserve")

    def test_timeout_is_reported_and_preserves_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            env = {"pywin32": True, "excel_executable": "x", "excel_arch": "x64", "xll": {"architecture": "x64"}}
            args = ["runner", "--output", tmp, "--xll", "test.xll", "--scenario", "REG-ADD", "--timeout", "1"]
            with mock.patch.object(sys, "argv", args), mock.patch.object(runner, "environment", return_value=env), \
                 mock.patch.object(runner, "validate_environment", return_value=[]), \
                 mock.patch.object(runner.subprocess, "run", side_effect=subprocess.TimeoutExpired("worker", 1)), \
                 contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(runner.main(), 1)
            data = json.loads(next(Path(tmp).glob("*/report.json")).read_text(encoding="utf-8"))
            self.assertEqual(data["scenarios"][0]["status"], "BLOCKED")
            self.assertIn("no Excel process was terminated", data["scenarios"][0]["observed"])

    def test_selection_and_preservation(self):
        with tempfile.TemporaryDirectory() as tmp:
            env = {"pywin32": True, "excel_executable": None, "excel_arch": None, "xll": None}
            args = ["runner", "--output", tmp, "--scenario", "REF-READ"]
            with mock.patch.object(sys, "argv", args), mock.patch.object(runner, "environment", return_value=env), contextlib.redirect_stdout(io.StringIO()):
                runner.main()
                runner.main()
            reports = list(Path(tmp).glob("*/report.json"))
            self.assertEqual(len(reports), 2)
            data = json.loads(reports[0].read_text(encoding="utf-8"))
            self.assertEqual([x["id"] for x in data["scenarios"]], ["REF-READ"])
            self.assertEqual(data["scenarios"][0]["status"], "BLOCKED")

    def test_unknown_selection_rejected(self):
        with mock.patch.object(sys, "argv", ["runner", "--scenario", "nonsense"]), contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit): runner.main()

    def test_autonomous_wake_is_explicitly_blocked(self):
        with tempfile.TemporaryDirectory() as tmp:
            env = {"pywin32": True, "excel_executable": None, "excel_arch": None, "xll": None}
            with mock.patch.object(sys, "argv", ["runner", "--output", tmp, "--scenario", "WAKE"]), \
                 mock.patch.object(runner, "environment", return_value=env), contextlib.redirect_stdout(io.StringIO()):
                runner.main()
            data = json.loads(next(Path(tmp).glob("*/report.json")).read_text(encoding="utf-8"))
            self.assertEqual(data["scenarios"][0]["status"], "BLOCKED")
            self.assertIn("not implemented", data["scenarios"][0]["observed"])


if __name__ == "__main__":
    unittest.main()
