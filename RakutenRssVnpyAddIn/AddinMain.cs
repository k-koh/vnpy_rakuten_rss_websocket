using System;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using ExcelDna.Integration;

public class AddinMain : IExcelAddIn
{
    private WebSocketServer _ws;
    private Thread _readLoopThread;
    private static readonly string LogFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "addin.log");

    public void AutoOpen()
    {
        _ws = new WebSocketServer("http://localhost:8765/ws/");
        _ws.Start();
        Log("[AddinMain] WebSocketServer started.");

        _readLoopThread = new Thread(ReadLoop);
        _readLoopThread.IsBackground = true;
        _readLoopThread.Start();
        Log("[AddinMain] WebSocketServer ReadLoop start.");
    }

    public void AutoClose()
    {
        _ws?.Stop();
        Log("[AddinMain] WebSocketServer stopped.");
    }

    private void ReadLoop()
    {
        Log("[AddinMain] ReadLoop started.");
        try
        {
            while (true)
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        var rows = OptionSheetReader.ReadOptionRows();
                        foreach (var row in rows)
                        {
                            string json = row.ToJson();
                            _ws.Broadcast(json);
                            //Log($"[AddinMain] Broadcasted: {json}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"[AddinMain] Exception in Macro: {ex}");
                    }
                });
                Thread.Sleep(5000);
            }
        }
        catch (Exception ex)
        {
            Log($"[AddinMain] Exception in ReadLoop: {ex}");
        }
    }

    public static void Log(string message)
    {
        try
        {
            File.AppendAllText(LogFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}{Environment.NewLine}");
        }
        catch { /* Ignore logging errors */ }
    }
}
