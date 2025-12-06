# vnpy_rakuten_rss_websocket
楽天RSS → Excel → WebSocket → vn.py を実現する

このプロジェクトは以下を実現します：

- Excel（楽天RSS アドイン）で日経225オプションのデータを取得
- VSTO Excel アドインでセルを監視
- WebSocket サーバーを内蔵してリアルタイム配信
- vn.py または Python の WebSocket クライアントが受信

````

# ExcelAddin-RakutenRSS-WebSocket

This project connects:

- 楽天RSS (MarketSpeed RSS)
- Excel (RSS data inside cells)
- VSTO Excel Add-in (C#)
- Built-in WebSocket Server
- Python / vn.py WebSocket Client

## Features
- Reads Nikkei225 option quotes from Excel (via Rakuten RSS)
- Monitors cells in real-time
- Broadcasts updates to WebSocket clients as JSON
- Python client included

## Project Structure
See folder tree.

## Build
1. Visual Studio → Create VSTO Excel Add-in (.NET Framework 4.8)
2. Install NuGet:
    - Fleck
    - Newtonsoft.Json

## Run
1. Start Excel (with this add-in installed)
2. WebSocket server auto-starts on port 18080
3. Run Python client:


ExcelAddin-RakutenRSS-WebSocket/
├── README.md
├── ExcelAddin/
│   ├── ExcelAddin.csproj
│   ├── ThisAddin.cs
│   ├── ExcelAddin.vsto
│   ├── Models/
│   │   └── OptionQuote.cs
│   ├── Services/
│   │   ├── RssReader.cs
│   │   └── OptionMonitor.cs
│   ├── WebSocket/
│   │   └── WsServer.cs
│   ├── Utils/
│   │   └── Logger.cs
│   └── Properties/
│       ├── AssemblyInfo.cs
│       └── Resources.resx
└── PythonClient/
    ├── client.py
    └── requirements.txt