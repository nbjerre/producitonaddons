using WorksheetAPI.Models;

namespace WorksheetAPI.Interfaces;

/// <summary>
/// Manages SAP company connections
/// </summary>
public interface ISapConnectionService
{
    /// <summary>
    /// Get the main/default SAP company
    /// </summary>
    SapCompany GetMainCompany();
    
    /// <summary>
    /// Get a SAP company by name
    /// </summary>
    SapCompany? GetCompany(string name);
    
    /// <summary>
    /// Get all active SAP companies
    /// </summary>
    IEnumerable<SapCompany> GetAllCompanies();
    
    /// <summary>
    /// Check if SAP integration is enabled
    /// </summary>
    bool IsSapEnabled { get; }
}
