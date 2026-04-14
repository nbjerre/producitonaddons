namespace WebAPI.Models.SalesProduction;

/// <summary>
/// Input model for recursive creation of sub production orders.
/// </summary>
public sealed class CreateSubProductionOrdersRequest
{
    public string ParentItemCode { get; set; } = string.Empty;
    public decimal ParentQuantity { get; set; }
    public string Project { get; set; } = string.Empty;
    public DateTime ShipDate { get; set; }
    public string CardCode { get; set; } = string.Empty;
    public int OrderDocEntry { get; set; }
    public int OrderLine { get; set; }
    public int VisOrder { get; set; }
    public int ProductionBaseEntry { get; set; }
    public int ProductionBaseLine { get; set; }
    public bool RemoveDelivery { get; set; }
    public int RcsDelDays { get; set; }
    public int Depth { get; set; }
}
