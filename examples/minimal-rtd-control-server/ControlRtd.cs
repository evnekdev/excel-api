using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Interop.Excel;

// Test-only control built against Microsoft's installed Office PIA. It is not
// linked by any Rust package and is never registered by an ordinary build.
[ComVisible(true)]
[Guid("F370A35B-7251-49E7-9FB2-3D6655FD1778")]
[ProgId("ExcelApi.ControlRtd")]
[ClassInterface(ClassInterfaceType.None)]
public sealed class ControlRtd : IRtdServer
{
    private readonly object gate = new object();
    private readonly HashSet<int> topics = new HashSet<int>();
    private IRTDUpdateEvent callback;
    private Timer timer;
    private int counter;
    private bool dirty;
    private bool stopping;

    public int ServerStart(IRTDUpdateEvent callbackObject)
    {
        Record("ServerStart");
        if (callbackObject == null) return 0;
        lock (gate)
        {
            callback = callbackObject;
            stopping = false;
            timer = new Timer(Tick, null, 250, 500);
        }
        return 1;
    }

    public object ConnectData(int topicId, ref Array strings, ref bool getNewValues)
    {
        Record("ConnectData");
        string topic = strings != null && strings.Length > 0 ? Convert.ToString(strings.GetValue(0), CultureInfo.InvariantCulture) : null;
        if (!String.Equals(topic, "COUNTER", StringComparison.OrdinalIgnoreCase)) return "#N/A";
        lock (gate) topics.Add(topicId);
        getNewValues = true;
        return Volatile.Read(ref counter);
    }

    public Array RefreshData(ref int topicCount)
    {
        Record("RefreshData");
        lock (gate)
        {
            int[] ids = new int[topics.Count];
            topics.CopyTo(ids);
            object[,] result = new object[2, ids.Length];
            for (int i = 0; i < ids.Length; i++)
            {
                result[0, i] = ids[i];
                result[1, i] = counter;
            }
            topicCount = ids.Length;
            dirty = false;
            return result;
        }
    }

    public void DisconnectData(int topicId)
    {
        Record("DisconnectData");
        lock (gate) topics.Remove(topicId);
    }

    public int Heartbeat()
    {
        Record("Heartbeat");
        lock (gate) return stopping ? 0 : 1;
    }

    public void ServerTerminate()
    {
        Record("ServerTerminate");
        Timer old;
        lock (gate)
        {
            stopping = true;
            old = timer;
            timer = null;
            callback = null;
            topics.Clear();
        }
        if (old != null) old.Dispose();
    }

    private void Tick(object ignored)
    {
        IRTDUpdateEvent current = null;
        lock (gate)
        {
            if (stopping || topics.Count == 0 || dirty) return;
            counter++;
            dirty = true;
            current = callback;
        }
        try { if (current != null) current.UpdateNotify(); }
        catch (COMException) { }
    }

    private static void Record(string method)
    {
        string path = Environment.GetEnvironmentVariable("EXCEL_API_CONTROL_RTD_DIAGNOSTICS");
        if (String.IsNullOrEmpty(path)) return;
        string line = String.Format(CultureInfo.InvariantCulture,
            "{{\"timestamp_utc\":\"{0:o}\",\"process_id\":{1},\"thread_id\":{2},\"method\":\"{3}\"}}",
            DateTime.UtcNow, System.Diagnostics.Process.GetCurrentProcess().Id, Thread.CurrentThread.ManagedThreadId, method);
        try { lock (typeof(ControlRtd)) File.AppendAllText(path, line + Environment.NewLine); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
