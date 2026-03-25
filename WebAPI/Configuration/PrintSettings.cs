namespace WebAPI.Configuration;

public class PrintSettings
{
    public const string SectionName = "PrintSettings";

    public string CrystalBaseUrl { get; set; } = string.Empty;
    public string ProductionReportCode { get; set; } = "WOR10003";
    public int ProductionObjectId { get; set; } = 202;
}