using Newtonsoft.Json;

public class OptionRow
{
    public int RowIndex { get; set; }
    public string CallIV { get; set; }
    public string CallDelta { get; set; }
    public string CallGamma { get; set; }
    public string CallTheta { get; set; }
    public string CallVega { get; set; }
    public string CallPrice { get; set; }
    public string CallMonth { get; set; }
    public string CallStrike { get; set; }

    public string ToJson()
    {
        return JsonConvert.SerializeObject(this);
    }
}
