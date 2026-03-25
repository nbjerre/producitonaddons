using B1SLayer;
using WorksheetAPI.Configuration;

namespace WorksheetAPI.Models;

/// <summary>
/// Represents a SAP company with Service Layer connection management
/// </summary>
public class SapCompany
{
    public bool Active { get; set; }
    public string Name { get; set; } = string.Empty;
    public string CompanyDb { get; set; } = string.Empty;
    public string UserName { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public string ServiceLayerUrl { get; set; } = string.Empty;

    private SLConnection? _slConnection;
    private readonly ILogger? _logger;

    public SapCompany() { }

    public SapCompany(SapCompanyConfig config, ILogger? logger = null)
    {
        Active = config.Active;
        Name = config.Name;
        CompanyDb = config.CompanyDb;
        UserName = config.UserName;
        Password = config.Password;
        ServiceLayerUrl = config.ServiceLayerUrl;
        _logger = logger;
    }

    /// <summary>
    /// Get the Service Layer connection (lazy initialization)
    /// </summary>
    public SLConnection Connection
    {
        get
        {
            if (_slConnection == null)
            {
                Connect();
            }
            return _slConnection!;
        }
    }

    /// <summary>
    /// Initialize the Service Layer connection
    /// </summary>
    public void Connect()
    {
        if (_slConnection == null)
        {
            _slConnection = new SLConnection(ServiceLayerUrl, CompanyDb, UserName, Password);
            _slConnection.NumberOfAttempts = 3;
            _slConnection.BatchRequestTimeout = TimeSpan.FromSeconds(600);
            _logger?.LogInformation("[SAP] Connected to {CompanyName} [{CompanyDb}] (new SLConnection created)", Name, CompanyDb);
        }
    }

    /// <summary>
    /// Disconnect from Service Layer
    /// </summary>
    public async Task DisconnectAsync()
    {
        if (_slConnection != null)
        {
            await _slConnection.LogoutAsync();
            _slConnection = null;
            _logger?.LogInformation("[SAP] Disconnected from {CompanyName}", Name);
        }
    }

    /// <summary>
    /// Reconnect to Service Layer
    /// </summary>
    public void Reconnect()
    {
        _slConnection = null;
        Connect();
        _logger?.LogInformation("[SAP] Reconnected to {CompanyName}", Name);
    }

    public override string ToString() => $"{Name} [{CompanyDb}]";
}
