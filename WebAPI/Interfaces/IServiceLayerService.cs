using WorksheetAPI.Models;

namespace WorksheetAPI.Interfaces;

/// <summary>
/// Service for interacting with SAP Service Layer
/// </summary>
public interface IServiceLayerService
{
    /// <summary>
    /// Get data from Service Layer
    /// </summary>
    Task<ServiceLayerResult<T>> GetAsync<T>(SapCompany company, string resource, string? select = null, string? filter = null, int? top = null);
    
    /// <summary>
    /// Get data as string from Service Layer
    /// </summary>
    Task<ServiceLayerResult<string>> GetStringAsync(SapCompany company, string resource, string? select = null, string? filter = null, int? top = null);
    
    /// <summary>
    /// Post data to Service Layer
    /// </summary>
    Task<ServiceLayerResult<T>> PostAsync<T>(SapCompany company, string resource, object data);
    
    /// <summary>
    /// Update data via Service Layer (PATCH)
    /// </summary>
    Task<ServiceLayerResult<bool>> PatchAsync(SapCompany company, string resource, object data);
    
    /// <summary>
    /// Delete data via Service Layer
    /// </summary>
    Task<ServiceLayerResult<bool>> DeleteAsync(SapCompany company, string resource);
}

/// <summary>
/// Result wrapper for Service Layer operations
/// </summary>
public class ServiceLayerResult<T>
{
    public bool Success { get; set; }
    public T? Data { get; set; }
    public string? Error { get; set; }

    public static ServiceLayerResult<T> Ok(T data) => new() { Success = true, Data = data };
    public static ServiceLayerResult<T> Fail(string error) => new() { Success = false, Error = error };
}
