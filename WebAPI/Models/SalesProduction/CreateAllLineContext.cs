namespace WebAPI.Models.SalesProduction;

/// <summary>
/// Internal context for a sales order line that can create a production order.
/// </summary>
public sealed class CreateAllLineContext
{
    public string ItemCode { get; set; } = string.Empty;
    public int LineNum { get; set; }
    public int VisOrder { get; set; }
    public decimal Quantity { get; set; }
    public DateTime DueDate { get; set; }
    public string Project { get; set; } = string.Empty;
    public int DelDays { get; set; }
}
