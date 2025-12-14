using Newtonsoft.Json;
using System;

public class OptionRow
{
    public string Symbol { get; set; }
    public string SymbolName { get; set; }
    public string TradingUnit { get; set; }
    public string DerivMonth { get; set; }
    public string PutOrCall { get; set; }
    public string StrikePrice { get; set; }
    public string TradeStart { get; set; }
    public string TradeEnd { get; set; }
    public string Sell1_Price { get; set; }
    public string Sell1_Qty { get; set; }
    public string Buy1_Price { get; set; }
    public string Buy1_Qty { get; set; }
    public string CurrentPrice { get; set; }
    public string TradingVolume { get; set; }
    public string TradingValue { get; set; }
    public string OpeningPrice { get; set; }
    public string HighPrice { get; set; }
    public string LowPrice { get; set; }
    public string PreviousClose { get; set; }

    public string ToJson()
    {
        return JsonConvert.SerializeObject(this);
    }
}


public class OptionContract
{
    public string Symbol { get; set; }
    public string SymbolName { get; set; }
    public string TradingUnit { get; set; }
    public string DerivMonth { get; set; }
    public string PutOrCall { get; set; }
    public string StrikePrice { get; set; }
    public string TradeStart { get; set; }
    public string TradeEnd { get; set; }

    public string ToJson()
    {
        return JsonConvert.SerializeObject(this);
    }
}


public class SymbolInfo
{
    public string Symbol { get; set; }
    public string Exchange { get; set; }
}
