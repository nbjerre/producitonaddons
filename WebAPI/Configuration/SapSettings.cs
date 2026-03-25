namespace WorksheetAPI.Configuration;

/// <summary>
/// SAP configuration settings from appsettings.json
/// </summary>
public class SapSettings
{
    public const string SectionName = "SapSettings";
    
    public bool Debug { get; set; }
    public bool UseSap { get; set; } = true;
    public int RetryCount { get; set; } = 3;
    public List<SapCompanyConfig> Companies { get; set; } = [];
}

/// <summary>
/// Configuration for a single SAP company connection
/// </summary>
public class SapCompanyConfig
{
    public bool Active { get; set; }
    public string Name { get; set; } = string.Empty;
    public string CompanyDb { get; set; } = string.Empty;
    public string UserName { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public string ServiceLayerUrl { get; set; } = string.Empty;
}
