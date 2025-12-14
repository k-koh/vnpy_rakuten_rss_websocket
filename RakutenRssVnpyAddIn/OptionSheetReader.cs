using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
//using ExcelDna.Util;

public static class OptionSheetReader
{
    public static List<OptionRow> ReadOptionRows()
    {
        List<OptionRow> list = new List<OptionRow>();
        string sheet = "設定";
        // Call data: AN~BF (columns 40~58), Put data: BH~BZ (columns 60~78)
        // We'll use Excel column letters for mapping
        string[] callCols = { "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF" };
        string[] putCols  = { "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ" };

        for (int row = 10; row <= 174; row++)
        {
            // --- Call data ---
            var callSymbol = Get(sheet, callCols[0], row);
            if (!string.IsNullOrWhiteSpace(callSymbol))
            {
                var obj = new OptionRow
                {
                    Symbol = callSymbol,
                    SymbolName = Get(sheet, callCols[1], row),
                    TradingUnit = Get(sheet, callCols[2], row),
                    DerivMonth = Get(sheet, callCols[3], row),
                    PutOrCall = Get(sheet, callCols[4], row),
                    StrikePrice = Get(sheet, callCols[5], row),
                    TradeStart = Get(sheet, callCols[6], row),
                    TradeEnd = Get(sheet, callCols[7], row),
                    Sell1_Price = Get(sheet, callCols[8], row),
                    Sell1_Qty = Get(sheet, callCols[9], row),
                    Buy1_Price = Get(sheet, callCols[10], row),
                    Buy1_Qty = Get(sheet, callCols[11], row),
                    CurrentPrice = Get(sheet, callCols[12], row),
                    TradingVolume = Get(sheet, callCols[13], row),
                    TradingValue = Get(sheet, callCols[14], row),
                    OpeningPrice = Get(sheet, callCols[15], row),
                    HighPrice = Get(sheet, callCols[16], row),
                    LowPrice = Get(sheet, callCols[17], row),
                    PreviousClose = Get(sheet, callCols[18], row)
                };
                list.Add(obj);
            }
            // --- Put data ---
            var putSymbol = Get(sheet, putCols[0], row);
            if (!string.IsNullOrWhiteSpace(putSymbol))
            {
                var obj = new OptionRow
                {
                    Symbol = putSymbol,
                    SymbolName = Get(sheet, putCols[1], row),
                    TradingUnit = Get(sheet, putCols[2], row),
                    DerivMonth = Get(sheet, putCols[3], row),
                    PutOrCall = Get(sheet, putCols[4], row),
                    StrikePrice = Get(sheet, putCols[5], row),
                    TradeStart = Get(sheet, putCols[6], row),
                    TradeEnd = Get(sheet, putCols[7], row),
                    Sell1_Price = Get(sheet, putCols[8], row),
                    Sell1_Qty = Get(sheet, putCols[9], row),
                    Buy1_Price = Get(sheet, putCols[10], row),
                    Buy1_Qty = Get(sheet, putCols[11], row),
                    CurrentPrice = Get(sheet, putCols[12], row),
                    TradingVolume = Get(sheet, putCols[13], row),
                    TradingValue = Get(sheet, putCols[14], row),
                    OpeningPrice = Get(sheet, putCols[15], row),
                    HighPrice = Get(sheet, putCols[16], row),
                    LowPrice = Get(sheet, putCols[17], row),
                    PreviousClose = Get(sheet, putCols[18], row)
                };
                list.Add(obj);
            }
        }
        return list;
    }

    public static Dictionary<string, OptionContract> ReadOptionContracts()
    {
        Dictionary<string, OptionContract> contracts = new Dictionary<string, OptionContract>();
        string sheet = "設定";
        string[] callCols = { "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU" };
        string[] putCols = { "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO" };

        for (int row = 10; row <= 174; row++)
        {
            // --- Call data ---
            var callSymbol = Get(sheet, callCols[0], row);
            if (!string.IsNullOrWhiteSpace(callSymbol))
            {
                var contract = new OptionContract
                {
                    Symbol = callSymbol,
                    SymbolName = Get(sheet, callCols[1], row),
                    TradingUnit = Get(sheet, callCols[2], row),
                    DerivMonth = Get(sheet, callCols[3], row),
                    PutOrCall = Get(sheet, callCols[4], row),
                    StrikePrice = Get(sheet, callCols[5], row),
                    TradeStart = Get(sheet, callCols[6], row),
                    TradeEnd = Get(sheet, callCols[7], row)
                };
                contracts[callSymbol] = contract;
            }
            // --- Put data ---
            var putSymbol = Get(sheet, putCols[0], row);
            if (!string.IsNullOrWhiteSpace(putSymbol))
            {
                var contract = new OptionContract
                {
                    Symbol = putSymbol,
                    SymbolName = Get(sheet, putCols[1], row),
                    TradingUnit = Get(sheet, putCols[2], row),
                    DerivMonth = Get(sheet, putCols[3], row),
                    PutOrCall = Get(sheet, putCols[4], row),
                    StrikePrice = Get(sheet, putCols[5], row),
                    TradeStart = Get(sheet, putCols[6], row),
                    TradeEnd = Get(sheet, putCols[7], row)
                };
                contracts[putSymbol] = contract;
            }
        }
        return contracts;
    }


    public static Dictionary<string, string> ReadSymbolNames()
    {
        Dictionary<string, string> names = new Dictionary<string, string>();
        string sheet = "設定";
        string[] callCols = { "AN", "AQ", "AR", "AS" };
        string[] putCols  = { "BH", "BK", "BL", "BM" };

        for (int row = 10; row <= 174; row++)
        {
            // --- Call data ---
            var callSymbol = Get(sheet, callCols[0], row);
            if (!string.IsNullOrWhiteSpace(callSymbol))
            {
                string DerivMonth  = Get(sheet, callCols[1], row);
                string PutOrCall   = Get(sheet, callCols[2], row);
                string StrikePrice = Get(sheet, callCols[3], row);

                string name = $"{DerivMonth}-{PutOrCall}-{StrikePrice}";
                names[name] = callSymbol;
            }
            // --- Put data ---
            var putSymbol = Get(sheet, putCols[0], row);
            if (!string.IsNullOrWhiteSpace(putSymbol))
            {
                string DerivMonth  = Get(sheet, putCols[1], row);
                string PutOrCall   = Get(sheet, putCols[2], row);
                string StrikePrice = Get(sheet, putCols[3], row);
                string name = $"{DerivMonth}-{PutOrCall}-{StrikePrice}";
                names[name] = putSymbol;
            }
        }
        return names;
    }

    private static string Get(string sheet, string col, int row)
    {
        try
        {
            // dynamic を使って Excel.Application 型のキャストを不要に
            dynamic app = ExcelDnaUtil.Application;
            dynamic ws = app.ActiveWorkbook.Worksheets[sheet];
            dynamic r = ws.Range[$"{col}{row}"];
            object v = r.Value;

            // --- Null（空白セル） ---
            if (v == null)
            {
                return "";
            }

            // --- String but empty ---
            if (v is string str && string.IsNullOrWhiteSpace(str))
            {
                return "";
            }


            // Excel の COM エラーはすべて 0x800A0000 (-2146828288) 付近
            if (v is int iv && (uint)iv >= 0x800A0000)
                return "";

            // --- Normal value ---
            string s = v.ToString();
            return s;
        }
        catch (Exception ex)
        {
            AddinMain.Log($"[OptionSheetReader] ERROR at {sheet}!{col}{row}: {ex.Message}");
            return null;
        }
    }
}
