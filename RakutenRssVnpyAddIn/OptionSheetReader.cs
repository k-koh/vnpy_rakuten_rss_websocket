using ExcelDna.Integration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
//using ExcelDna.Util;

public static class OptionSheetReader
{
    public static List<OptionRow> ReadOptionRows()
    {
        List<OptionRow> list = new List<OptionRow>();

        for (int row = 7; row <= 87; row++)
        {
            string sheet = "現在値";

            var obj = new OptionRow()
            {
                RowIndex = row,
                CallIV = Get(sheet, "G", row),
                CallDelta = Get(sheet, "H", row),
                CallGamma = Get(sheet, "I", row),
                CallTheta = Get(sheet, "J", row),
                CallVega = Get(sheet, "K", row),
                CallPrice = Get(sheet, "L", row),
                CallMonth = Get(sheet, "M", row),
                CallStrike = Get(sheet, "N", row),
            };

            list.Add(obj);
        }
        return list;
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
            AddinMain.Log($"[OptionSheetReader] ERROR: {ex.Message}");
            return null;
        }
    }
}
