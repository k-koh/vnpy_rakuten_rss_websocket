using System;
using System.Collections.Generic;
using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Threading;

public class WebSocketServer
{
    private readonly HttpListener _listener = new HttpListener();
    private readonly List<WebSocket> _clients = new List<WebSocket>();
    private CancellationTokenSource _cts = new CancellationTokenSource();

    public WebSocketServer(string prefix)
    {
        _listener.Prefixes.Add(prefix);
    }

    public void Start()
    {
        _listener.Start();
        ListenLoop();
    }

    public void Stop()
    {
        _cts.Cancel();
        _listener.Stop();
    }

    private async void ListenLoop()
    {
        while (!_cts.IsCancellationRequested)
        {
            HttpListenerContext ctx = await _listener.GetContextAsync();

            if (!ctx.Request.IsWebSocketRequest)
            {
                ctx.Response.StatusCode = 400;
                ctx.Response.Close();
                continue;
            }

            var wsCtx = await ctx.AcceptWebSocketAsync(null);
            var ws = wsCtx.WebSocket;

            lock (_clients) _clients.Add(ws);
            AddinMain.Log($"[WebSocketServer] Client connected. Total clients: {_clients.Count}");
        }
    }

    public void Broadcast(string msg)
    {
        byte[] data = Encoding.UTF8.GetBytes(msg);

        lock (_clients)
        {
            int clientIndex = 0;
            foreach (var ws in _clients.ToArray())
            {
                if (ws.State == WebSocketState.Open)
                {
                    ws.SendAsync(new ArraySegment<byte>(data),
                                 WebSocketMessageType.Text,
                                 true,
                                 CancellationToken.None);
                    //AddinMain.Log($"[WebSocketServer] Sent data to client #{clientIndex} (State: {ws.State}, SubProtocol: {ws.SubProtocol ?? "none"}): {msg}");
                }
                clientIndex++;
            }
        }
    }
}
