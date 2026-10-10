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
    // 設定 sheet layout (ATM ±22000, 500 step => 89 strikes per month)
    // Row 8-9: NVI/VIX, Row 10-11: Month 1/2 Future
    private const int FirstRow = 8;
    private const int Month1StartRow = 12;
    private const int Month1EndRow = 100;
    private const int Month2StartRow = 102;
    private const int Month2EndRow = 190;
    private const int LastRow = Month2EndRow;

    public static List<OptionRow> ReadOptionRows()
    {
        List<OptionRow> list = new List<OptionRow>();
        dynamic sheet = FindSettingsSheet();
        if (sheet == null) return list;
        // Call data: AN~BF (columns 40~58), Put data: BH~BZ (columns 60~78)
        // We'll use Excel column letters for mapping
        string[] callCols = { "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF" };
        string[] putCols  = { "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ" };

        for (int row = FirstRow; row <= LastRow; row++)
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
        // Add Symbol for VICalculation index
        var viSymbol = "NVI-E";
        if (!string.IsNullOrWhiteSpace(viSymbol))
        {
            double viValue = OptionSheetReader.CalculateNikkeiVI();
            string vi = viValue.ToString("F2"); // Convert to string with 2 decimal places
            var obj = new OptionRow
            {
                Symbol = viSymbol,
                SymbolName  = "NVI-E",
                TradingUnit = "1",
                DerivMonth  = "vi-n1",
                PutOrCall   = "V",
                StrikePrice = "0",
                TradeStart  = "",
                TradeEnd    = "2026/02/12",
                Sell1_Price = vi,
                Sell1_Qty   = "1",
                Buy1_Price  = vi,
                Buy1_Qty    = "1",
                CurrentPrice= vi,
                TradingVolume="0",
                TradingValue="0",
                OpeningPrice= vi,
                HighPrice   = vi,
                LowPrice    = vi,
                PreviousClose= vi
            };
            list.Add(obj);
        }
        return list;
    }

    public static Dictionary<string, OptionContract> ReadOptionContracts()
    {
        Dictionary<string, OptionContract> contracts = new Dictionary<string, OptionContract>();
        dynamic sheet = FindSettingsSheet();
        if (sheet == null) return contracts;
        string[] callCols = { "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU" };
        string[] putCols = { "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO" };

        for (int row = FirstRow; row <= LastRow; row++)
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
        // Add Symbol for VICalculation index
        var viSymbol = "NVI-E";
        if (!string.IsNullOrWhiteSpace(viSymbol))
        {
            var contract = new OptionContract
            {
                Symbol = viSymbol,
                SymbolName  = "NVI-E",
                TradingUnit = "1",
                DerivMonth  = "vi-n1",
                PutOrCall   = "V",
                StrikePrice = "0",
                TradeStart  = "",
                TradeEnd    = "2026/02/12"
            };
            contracts[viSymbol] = contract;
        }
        return contracts;
    }


    public static Dictionary<string, string> ReadSymbolNames()
    {
        Dictionary<string, string> names = new Dictionary<string, string>();
        dynamic sheet = FindSettingsSheet();
        if (sheet == null) return names;
        string[] callCols = { "AN", "AQ", "AR", "AS" };
        string[] putCols  = { "BH", "BK", "BL", "BM" };

        for (int row = FirstRow; row <= LastRow; row++)
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
        // Add Symbol for VICalculation index
        var viSymbol = "NVI-E";
        if (!string.IsNullOrWhiteSpace(viSymbol))
        {
            string DerivMonth  = "vi-n1";
            string PutOrCall   = "V";
            string StrikePrice = "0";

            string name = $"{DerivMonth}-{PutOrCall}-{StrikePrice}";
            names[name] = viSymbol;
        }

        return names;
    }

    /// <summary>
    /// Calculate Nikkei VI (Volatility Index) from option data
    /// Row 10: Month 1 Future, Row 11: Month 2 Future
    /// Rows 12-100: Month 1 Options, Rows 102-190: Month 2 Options
    /// </summary>
    public static double CalculateNikkeiVI()
    {
        try
        {
            dynamic sheet = FindSettingsSheet();
            if (sheet == null) return 0;
            // Column mapping: AN=Symbol, AO=Name, AP=Unit, AQ=DerivMonth, AR=PutOrCall, AS=Strike, AV=Sell1, AX=Buy1, AZ=Current

            // Read underlying futures data
            // Row 10 = Month 1 Future, Row 11 = Month 2 Future
            var month1Future = new VICalculator.UnderlyingData
            {
                Price = ParseDouble(Get(sheet, "AZ", 10)), // CurrentPrice from Call side
                ExpiryDate = VICalculator.ParseExpiryDate(Get(sheet, "AQ", 10), DateTime.Now.Year)
            };

            var month2Future = new VICalculator.UnderlyingData
            {
                Price = ParseDouble(Get(sheet, "AZ", 11)), // CurrentPrice from Call side
                ExpiryDate = VICalculator.ParseExpiryDate(Get(sheet, "AQ", 11), DateTime.Now.Year)
            };

            //AddinMain.Log($"[VI] Month1 Future: Price={month1Future.Price}, Expiry={month1Future.ExpiryDate:yyyy-MM-dd}");
            //AddinMain.Log($"[VI] Month2 Future: Price={month2Future.Price}, Expiry={month2Future.ExpiryDate:yyyy-MM-dd}");

            // Read month 1 options (rows 12-100)
            var month1Options = ReadOptionsForVI(sheet, Month1StartRow, Month1EndRow);
            //AddinMain.Log($"[VI] Month1 Options count: {month1Options.Count}");
            
            // Read month 2 options (rows 102-190)
            var month2Options = ReadOptionsForVI(sheet, Month2StartRow, Month2EndRow);
            //AddinMain.Log($"[VI] Month2 Options count: {month2Options.Count}");

            // Calculate VI
            double vi = VICalculator.CalculateVI(
                month1Future,
                month2Future,
                month1Options,
                month2Options,
                DateTime.Now
            );

            //AddinMain.Log($"[VI] Calculated VI: {vi:F2}");
            return vi;
        }
        catch (Exception ex)
        {
            AddinMain.Log($"[OptionSheetReader] Error calculating VI: {ex.Message}\n{ex.StackTrace}");
            return 0;
        }
    }

    /// <summary>
    /// Read option data for VI calculation from specified row range
    /// </summary>
    private static List<VICalculator.OptionData> ReadOptionsForVI(dynamic sheet, int startRow, int endRow)
    {
        var options = new List<VICalculator.OptionData>();
        // Call columns: AN=Symbol, AS=Strike, AV=Sell1, AX=Buy1, AR=PutOrCall
        // Put columns:  BH=Symbol, BM=Strike, BP=Sell1, BR=Buy1, BL=PutOrCall
        string[] callCols = { "AN", "AS", "AV", "AX", "AR" };
        string[] putCols  = { "BH", "BM", "BP", "BR", "BL" };

        for (int row = startRow; row <= endRow; row++)
        {
            // --- Call options ---
            var callSymbol = Get(sheet, callCols[0], row);
            if (!string.IsNullOrWhiteSpace(callSymbol))
            {
                double strike = ParseDouble(Get(sheet, callCols[1], row));
                double sell = ParseDouble(Get(sheet, callCols[2], row));
                double buy = ParseDouble(Get(sheet, callCols[3], row));
                
                // Calculate mid price: average of bid and ask if both exist, otherwise use whichever exists
                double mid = 0;
                if (sell > 0 && buy > 0)
                    mid = (sell + buy) / 2.0;
                else if (sell > 0)
                    mid = sell;
                else if (buy > 0)
                    mid = buy;

                if (mid > 0 && strike > 0)
                {
                    options.Add(new VICalculator.OptionData
                    {
                        Symbol = callSymbol,
                        StrikePrice = strike,
                        MidPrice = mid,
                        IsCall = true
                    });
                }
            }

            // --- Put options ---
            var putSymbol = Get(sheet, putCols[0], row);
            if (!string.IsNullOrWhiteSpace(putSymbol))
            {
                double strike = ParseDouble(Get(sheet, putCols[1], row));
                double sell = ParseDouble(Get(sheet, putCols[2], row));
                double buy = ParseDouble(Get(sheet, putCols[3], row));
                
                // Calculate mid price
                double mid = 0;
                if (sell > 0 && buy > 0)
                    mid = (sell + buy) / 2.0;
                else if (sell > 0)
                    mid = sell;
                else if (buy > 0)
                    mid = buy;

                if (mid > 0 && strike > 0)
                {
                    options.Add(new VICalculator.OptionData
                    {
                        Symbol = putSymbol,
                        StrikePrice = strike,
                        MidPrice = mid,
                        IsCall = false
                    });
                }
            }
        }

        return options;
    }

    /// <summary>
    /// Safely parse string to double
    /// </summary>
    private static double ParseDouble(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return 0;
        
        double.TryParse(value, out double result);
        return result;
    }

    private const string SettingsSheetName = "設定";
    private static bool _settingsSheetMissingLogged = false;

    /// <summary>
    /// Find the 設定 sheet: ActiveWorkbook first, then any open workbook.
    /// Returns null when no workbook is open yet (e.g. right after AutoOpen) or none has the sheet.
    /// </summary>
    private static dynamic FindSettingsSheet()
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic sheet = FindSheetIn(app.ActiveWorkbook);
        if (sheet == null)
        {
            foreach (dynamic wb in app.Workbooks)
            {
                sheet = FindSheetIn(wb);
                if (sheet != null) break;
            }
        }

        if (sheet == null)
        {
            if (!_settingsSheetMissingLogged)
            {
                AddinMain.Log($"[OptionSheetReader] Waiting for a workbook with sheet '{SettingsSheetName}'.");
                _settingsSheetMissingLogged = true;
            }
        }
        else if (_settingsSheetMissingLogged)
        {
            AddinMain.Log($"[OptionSheetReader] Found sheet '{SettingsSheetName}' in {sheet.Parent.Name}.");
            _settingsSheetMissingLogged = false;
        }
        return sheet;
    }

    private static dynamic FindSheetIn(dynamic wb)
    {
        if (wb == null) return null;
        foreach (dynamic ws in wb.Worksheets)
        {
            if (ws.Name == SettingsSheetName) return ws;
        }
        return null;
    }

    private static string Get(dynamic ws, string col, int row)
    {
        try
        {
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
            AddinMain.Log($"[OptionSheetReader] ERROR at {SettingsSheetName}!{col}{row}: {ex.Message}");
            return null;
        }
    }
}
