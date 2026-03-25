namespace WorksheetAPI.Models;

/// <summary>
/// Generisk requestmodel til direkte kald af PlanUnlimited runneren.
/// </summary>
public class PlanUnlimitedRunRequest
{
    /// <summary>
    /// Knap/kommando der skal køres. Tilladte værdier omfatter blandt andet BtnPU, BtnPUL, BtnPR, BtnPl og BtnPlu.
    /// </summary>
    public string Btn { get; set; } = string.Empty;

    /// <summary>
    /// SAP DocEntry for salgsordren.
    /// </summary>
    public long DocEntry { get; set; }

    /// <summary>
    /// Linjenummer på salgsordren. Sæt typisk til -1 hvis det ikke bruges.
    /// </summary>
    public int LineNum { get; set; } = -1;

    /// <summary>
    /// Produktionsordrens dokumentnummer hvis runneren skal køres på en produktionsordre i stedet for en salgsordre.
    /// </summary>
    public int ProdDocNum { get; set; } = -1;

    /// <summary>
    /// Brugernavn der sendes videre til den eksterne runner.
    /// </summary>
    public string User { get; set; } = string.Empty;
}
