using System;
using System.Net;
using System.Text;
using System.Threading;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json;

public class RestApiServer
{
    private readonly HttpListener _listener = new HttpListener();
    private Thread _serverThread;
    private bool _running;
    private static readonly string LogFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "addin.log");
    public static List<string> RegisteredSymbols { get; } = new List<string>();
    public static Dictionary<string, OptionContract> OptionContracts { get; } = new Dictionary<string, OptionContract>();

    public RestApiServer(string prefix)
    {
        _listener.Prefixes.Add(prefix);
    }

    public void Start()
    {
        _running = true;
        _listener.Start();
        _serverThread = new Thread(ServerLoop);
        _serverThread.IsBackground = true;
        _serverThread.Start();
        Log("[RestApiServer] Started.");
    }

    public void Stop()
    {
        _running = false;
        _listener.Stop();
        Log("[RestApiServer] Stopped.");
    }

    private void ServerLoop()
    {
        while (_running)
        {
            try
            {
                var context = _listener.GetContext();
                ProcessRequest(context);
            }
            catch (Exception ex)
            {
                Log($"[RestApiServer] Exception: {ex}");
                Thread.Sleep(1000);
            }
        }
    }

    private void ProcessRequest(HttpListenerContext context)
    {
        string path = context.Request.Url.AbsolutePath;
        string method = context.Request.HttpMethod;
        Log($"[RestApiServer] {method} {path}");
        string response = "";
        int statusCode = 200;

        if (path == "/rakutenapi/token" && method == "POST")
        {
            response = "{\"Token\":\"dummy-token\",\"ResultCode\":0}";
        }
        else if (path == "/rakutenapi/unregister/all" && method == "PUT")
        {
            RegisteredSymbols.Clear();
            response = "{\"result\":\"all unregistered\"}";
        }
        else if (path == "/rakutenapi/register" && method == "PUT")
        {
            try
            {
                using (var reader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
                {
                    var body = reader.ReadToEnd();
                    var data = JsonConvert.DeserializeObject<RegisterRequest>(body);
                    if (data?.Symbols != null)
                    {
                        foreach (var s in data.Symbols)
                        {
                            RegisteredSymbols.Add(s.Symbol);
                            Log($"[RestApiServer] {method} {path} {s.Symbol}");
                        }
                        var registListJson = JsonConvert.SerializeObject(new { RegistList = RegisteredSymbols });
                        response = registListJson;
                    }
                    else
                    {
                        statusCode = 400;
                        response = "{\"error\":\"invalid symbols\"}";
                    }
                }
            }
            catch (Exception ex)
            {
                statusCode = 400;
                response = $"{{\"error\":\"exception\",\"message\":\"{ex.Message}\"}}";
            }
        }
        else if (path.StartsWith("/rakutenapi/symbol/") && method == "GET")
        {
            string symbol = path.Substring("/rakutenapi/symbol/".Length);
            if (OptionContracts.ContainsKey(symbol))
            {
                response = OptionContracts[symbol].ToJson();
            }
            else
            {
                statusCode = 404;
                response = "{\"error\":\"symbol not found\"}";
            }
        }
        else
        {
            statusCode = 404;
            response = "{\"error\":\"not found\"}";
        }

        context.Response.StatusCode = statusCode;
        context.Response.ContentType = "application/json";
        byte[] buffer = Encoding.UTF8.GetBytes(response);
        context.Response.OutputStream.Write(buffer, 0, buffer.Length);
        context.Response.OutputStream.Close();
    }

    private static void Log(string message)
    {
        try
        {
            File.AppendAllText(LogFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}{Environment.NewLine}");
        }
        catch { }
    }
}

public class RegisterRequest
{
    public List<SymbolInfo> Symbols { get; set; }
}
