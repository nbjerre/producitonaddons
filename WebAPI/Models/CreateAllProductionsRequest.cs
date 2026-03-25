namespace WorksheetAPI.Models;

/// <summary>
/// Request til oprettelse af produktionsordrer for alle gyldige linjer på en salgsordre.
/// </summary>
public class CreateAllProductionsRequest
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
    /// Angiver om eventuelle understyklister allerede er bekræftet af klienten.
    /// </summary>
    public bool ConfirmSubBoms { get; set; }

    /// <summary>
    /// Midlertidige justeringer for fundne understyklister inden oprettelse.
    /// </summary>
    public List<SubBomAdjustmentDto>? SubBomAdjustments { get; set; }
}
