# vnpy_rakuten_rss_websocket
楽天RSS → Excel → WebSocket → vn.py を実現する

このプロジェクトは以下を実現します：

- Excel（楽天RSS アドイン）で日経225オプションのデータを取得
- Excel XLL アドイン（C#、Excel-DNA）でセルを監視
- WebSocket サーバーを内蔵してリアルタイム配信
- vn.py または Python の WebSocket クライアントが寿信

# ExcelAddin-RakutenRSS-WebSocket

This project connects:

- 楽天RSS (MarketSpeed RSS)
- Excel (RSS data inside cells)
- Excel XLL Add-in (C#)
- Built-in WebSocket Server
- Python / vn.py WebSocket Client

## Features
- Reads Nikkei225 option quotes from Excel (via Rakuten RSS)
- Monitors cells in real-time
- Broadcasts updates to WebSocket clients as JSON

## Main Project Files

- `AddinMain.cs` - Add-in entry point, WebSocket server, and Excel integration
- `OptionSheetReader.cs` - Reads option data from Excel cells
- `WebSocketServer.cs` - Simple WebSocket server implementation
- `JsonModels.cs` - Data models for JSON serialization
- `Properties/AssemblyInfo.cs` - Assembly metadata

## Build Tools
- Microsoft Visual Studio Community 2026 (18.1.0)
- Microsoft .NET Framework 4.8.09032

## Build
1. Visual Studio → Class Library (.NET Framework 4.8)
2. Install NuGet:
    - ExcelDna.AddIn
    - Newtonsoft.Json

## Run
1. Start Excel (with this add-in installed)
2. WebSocket server auto-starts on port 8765
3. Connect a WebSocket client (e.g., Python, vn.py)

## File Structure

```
RakutenRssVnpyAddIn/
├── AddinMain.cs
├── OptionSheetReader.cs
├── WebSocketServer.cs
├── JsonModels.cs
├── Properties/
│   └── AssemblyInfo.cs
└── README.md
```
