using Microsoft.Extensions.Options;
using WorksheetAPI.Configuration;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;

namespace WorksheetAPI.Services;

/// <summary>
/// Implementation of SAP connection management
/// </summary>
public class SapConnectionService : ISapConnectionService
{
    private readonly SapSettings _settings;
    private readonly ILogger<SapConnectionService> _logger;
    private readonly List<SapCompany> _companies;
    private SapCompany? _mainCompany;

    public SapConnectionService(IOptions<SapSettings> settings, ILogger<SapConnectionService> logger)
    {
        _settings = settings.Value;
        _logger = logger;
        _companies = [];
        InitializeCompanies();
        // Eagerly connect to Service Layer for main company at startup
        _mainCompany?.Connect();
    }

    public bool IsSapEnabled => _settings.UseSap;

    private void InitializeCompanies()
    {
        foreach (var config in _settings.Companies.Where(c => c.Active))
        {
            var company = new SapCompany(config, _logger);
            _companies.Add(company);
            
            // First company named "Main" or first active company becomes the main company
            if (_mainCompany == null || config.Name.Equals("Main", StringComparison.OrdinalIgnoreCase))
            {
                _mainCompany = company;
            }
        }

        if (_mainCompany != null)
        {
            _logger.LogInformation("[SAP] Main company set to: {CompanyName}", _mainCompany.Name);
        }
        else
        {
            _logger.LogWarning("[SAP] No active SAP companies configured");
        }
    }

    public SapCompany GetMainCompany()
    {
        if (_mainCompany == null)
        {
            throw new InvalidOperationException("No SAP companies are configured or active");
        }
        return _mainCompany;
    }

    public SapCompany? GetCompany(string name)
    {
        return _companies.FirstOrDefault(c => 
            c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    public IEnumerable<SapCompany> GetAllCompanies()
    {
        return _companies.AsReadOnly();
    }
}
