namespace WorksheetAPI.Configuration;

public class PlanUnlimitedSettings
{
    public const string SectionName = "PlanUnlimited";

    public bool Enabled { get; set; } = false;
    public string ExePath { get; set; } = string.Empty;
    public int TimeoutSeconds { get; set; } = 180;
}
