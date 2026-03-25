namespace WorksheetAPI.Models;

/// <summary>
/// Requestmodel til oprettelse eller opdatering af SAP-brugere.
/// </summary>
public class SapUserCreateDto
{
    /// <summary>
    /// SAP bruger-id.
    /// </summary>
    public string USER_CODE { get; set; } = string.Empty;

    /// <summary>
    /// Password for brugeren. Kræves ved oprettelse.
    /// </summary>
    public string PASSWORD { get; set; } = string.Empty;

    /// <summary>
    /// Vist navn på brugeren.
    /// </summary>
    public string? U_NAME { get; set; }

    /// <summary>
    /// E-mailadresse.
    /// </summary>
    public string? E_Mail { get; set; }

    /// <summary>
    /// Afdeling i SAP.
    /// </summary>
    public string? Department { get; set; }

    /// <summary>
    /// Filial eller branche i SAP.
    /// </summary>
    public string? Branch { get; set; }

    /// <summary>
    /// Angiver om brugeren er superuser, typisk tYES eller tNO.
    /// </summary>
    public string? SUPERUSER { get; set; }

    /// <summary>
    /// Angiver om brugeren er låst, typisk tYES eller tNO.
    /// </summary>
    public string? Locked { get; set; }
}