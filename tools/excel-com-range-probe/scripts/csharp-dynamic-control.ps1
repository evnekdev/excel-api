Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not ('ExcelRangeProbe.Control' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelRangeProbe {
    public static class Control {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

        private static string Quote(string value) {
            return value == null ? "null" : "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        public static string Run() {
            dynamic app = null;
            dynamic workbook = null;
            dynamic sheet = null;
            uint pid = 0;
            bool workbookClosed = false;
            bool quitRequested = false;
            bool exited = false;
            try {
                var type = Type.GetTypeFromProgID("Excel.Application", true);
                app = Activator.CreateInstance(type);
                app.Visible = false;
                app.DisplayAlerts = false;
                IntPtr hwnd = new IntPtr((int)app.Hwnd);
                GetWindowThreadProcessId(hwnd, out pid);
                if (pid == 0) throw new InvalidOperationException("Hwnd-to-PID returned zero");
                workbook = app.Workbooks.Add();
                sheet = app.ActiveSheet;
                sheet.Range("A1").Value2 = 42;
                sheet.Range("B1:C2").Value2 = new object[,] { { "mixed", 42 }, { true, null } };
                object scalar = sheet.Range("A1").Value2;
                Array matrix = (Array)sheet.Range("B1:C2").Value2;
                string result = "{"
                    + "\"version\":\"csharp-dynamic-automation-control\","
                    + "\"status\":\"completed\","
                    + "\"projection\":\"C# dynamic COM; projection only, not raw VARIANT authority\","
                    + "\"owned_pid\":" + pid + ","
                    + "\"scalar_type\":" + Quote(scalar == null ? "null" : scalar.GetType().FullName) + ","
                    + "\"scalar_value\":" + Quote(Convert.ToString(scalar, System.Globalization.CultureInfo.InvariantCulture)) + ","
                    + "\"matrix_rank\":" + matrix.Rank + ","
                    + "\"matrix_lower_bounds\": [" + matrix.GetLowerBound(0) + "," + matrix.GetLowerBound(1) + "],"
                    + "\"matrix_lengths\": [" + matrix.GetLength(0) + "," + matrix.GetLength(1) + "]"
                    + "}";
                return result;
            }
            catch (Exception error) {
                return "{\"version\":\"csharp-dynamic-automation-control\",\"status\":\"failed\",\"error\":" + Quote(error.GetType().FullName) + "}";
            }
            finally {
                if (workbook != null) {
                    try { workbook.Close(false); workbookClosed = true; } catch { }
                    try { Marshal.FinalReleaseComObject(workbook); } catch { }
                }
                if (sheet != null) {
                    try { Marshal.FinalReleaseComObject(sheet); } catch { }
                }
                if (app != null) {
                    try { app.Quit(); quitRequested = true; } catch { }
                    try { Marshal.FinalReleaseComObject(app); } catch { }
                }
                if (pid != 0) {
                    try { using (var process = Process.GetProcessById((int)pid)) { exited = process.WaitForExit(15000); } } catch { exited = true; }
                }
            }
        }
    }
}
'@
}

[ExcelRangeProbe.Control]::Run()
