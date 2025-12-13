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
- Excel XLL Add-in (C#)
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
1. Visual Studio → Class Library(.NET Framework) Add-in (.NET Framework 4.8)
2. Install NuGet:
    - ExcelDna.AddIn
    - Newtonsoft.Json

## Release
Put build files into Microsoft Addin directory
C:\Users\YourName\AppData\Roaming\Microsoft\AddIns
- RakutenRssVnpyAddIn-AddIn64.xll
- RakutenRssVnpyAddIn-AddIn64.dna
- RakutenRssVnpyAddIn.dll
- Newtonsoft.Json.dll

## Register Add-in in Excel
1. File -> Options -> Add-ins -> Setting 
2. Check the RakutenRssVnpyAddIn

## Run
1. Start Excel (with this add-in installed)
2. WebSocket server auto-starts on port 8765
3. Run Python client:


ExcelAddin-RakutenRSS-WebSocket/
├── README.md
├── RakutenRssVnpyAddIn
│   ├── RakutenRssVnpyAddIn.csproj
│   ├── RakutenRssVnpyAddIn.slnx
│   ├── RakutenRssVnpyAddIn-AddIn.dna
│   ├── AddinMain.cs
│   ├── JsonModels.cs
│   ├── OptionSheetReader.cs
│   ├── WebSocketServer.cs
│   ├──Properties
│   │   ├──AssemblyInfo.cs
│   │   ├──ExcelDna.Build.props
└── client.py
