namespace WorksheetAPI.Models;

/// <summary>
/// Stores original sub-BOM values so temporary overrides can be restored.
/// </summary>
public sealed class SubBomSnapshot
{
    public string ItemCode { get; set; } = string.Empty;
    public decimal U_RCS_PQT { get; set; }
    public string U_RCS_ONSTO { get; set; } = "Y";
}
