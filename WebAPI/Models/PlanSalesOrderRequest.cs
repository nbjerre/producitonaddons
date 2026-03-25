namespace WorksheetAPI.Models;

/// <summary>
/// Request til PlanUnlimited-kørsel på salgsordre eller salgsordrelinje.
/// </summary>
public class PlanSalesOrderRequest
{
    /// <summary>
    /// SAP DocEntry for salgsordren.
    /// </summary>
    public long DocEntry { get; set; }

    /// <summary>
    /// Linjenummer på salgsordren. Bruges kun for linjebaserede kørsler.
    /// </summary>
    public int? LineNum { get; set; }

    /// <summary>
    /// Brugernavn der sendes videre til den eksterne runner.
    /// </summary>
    public string User { get; set; } = "manager";
}
