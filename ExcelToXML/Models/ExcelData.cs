namespace ExcelToXML.Models;

public class ExcelData
{
    public List<ExcelRows> Rows { get; set; }
}

public class ExcelRows
{
    public DateTime Date { get; set; }
    public string Type { get; set; }
    public string EventDescription { get; set; }
    public decimal Amount { get; set; }
    public string Currency { get; set; }
    public DateTime CreatedAt { get; set; }
}