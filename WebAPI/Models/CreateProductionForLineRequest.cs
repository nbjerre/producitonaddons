namespace WorksheetAPI.Models;

/// <summary>
/// Request til oprettelse af en produktionsordre for en enkelt salgsordrelinje.
/// </summary>
public class CreateProductionForLineRequest
{
    /// <summary>
    /// SAP DocEntry for salgsordren.
    /// </summary>
    public int SalesOrderDocEntry { get; set; }

    /// <summary>
    /// Alternativ SAP DocNum hvis DocEntry ikke kendes.
    /// </summary>
    public int? SalesOrderDocNum { get; set; }

    /// <summary>
    /// Varekode til den linje der skal behandles.
    /// </summary>
    public string? ItemCode { get; set; }

    /// <summary>
    /// Linjenummer på salgsordren.
    /// </summary>
    public int? LineNum { get; set; }

    /// <summary>
    /// Angiver om eventuelle understyklister allerede er bekræftet af klienten.
    /// </summary>
    public bool ConfirmSubBoms { get; set; }

    /// <summary>
    /// Midlertidige justeringer for fundne understyklister inden oprettelse.
    /// </summary>
    public List<SubBomAdjustmentDto>? SubBomAdjustments { get; set; }
}

/// <summary>
/// Midlertidige justeringer der anvendes på en understykliste under oprettelsesflowet.
/// </summary>
public class SubBomAdjustmentDto
{
    /// <summary>
    /// Varekoden på understyklisten.
    /// </summary>
    public string ItemCode { get; set; } = string.Empty;

    /// <summary>
    /// Midlertidig værdi til feltet U_RCS_PQT.
    /// </summary>
    public decimal? U_RCS_PQT { get; set; }

    /// <summary>
    /// Midlertidig værdi til feltet U_RCS_ONSTO, typisk Y eller N.
    /// </summary>
    public string? U_RCS_ONSTO { get; set; }
}
