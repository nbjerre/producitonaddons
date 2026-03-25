using Microsoft.Extensions.Options;
using Newtonsoft.Json;
using WorksheetAPI.Configuration;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;
using B1SLayer;

namespace WorksheetAPI.Services;

/// <summary>
/// Implementation of Service Layer interactions
/// </summary>
public class ServiceLayerService : IServiceLayerService
{
    private readonly ILogger<ServiceLayerService> _logger;
    private readonly SapSettings _settings;

    public ServiceLayerService(ILogger<ServiceLayerService> logger, IOptions<SapSettings> settings)
    {
        _logger = logger;
        _settings = settings.Value;
    }

    public async Task<ServiceLayerResult<T>> GetAsync<T>(SapCompany company, string resource, string? select = null, string? filter = null, int? top = null)
    {
        var retryCount = 0;

        while (retryCount <= _settings.RetryCount)
        {
            try
            {
                var request = company.Connection.Request(resource);
                
                if (!string.IsNullOrEmpty(filter))
                    request = request.Filter(filter);
                if (!string.IsNullOrEmpty(select))
                    request = request.Select(select);

                request = request.WithCaseInsensitive();

                if (top.HasValue)
                {
                    request = request.Top(top.Value)
                        .WithHeader("Prefer", $"odata.maxpagesize={top.Value}");
                }

                var data = await request.GetAsync<T>();

                LogSuccess("GET", company, resource, select, filter);
                return ServiceLayerResult<T>.Ok(data);
            }
            catch (Exception ex)
            {
                if (ShouldRetry(ex) && retryCount < _settings.RetryCount)
                {
                    LogRetry("GET", company, resource, retryCount);
                    await Task.Delay(1000);
                    retryCount++;
                    continue;
                }

                LogError("GET", company, resource, ex, select, filter);
                return ServiceLayerResult<T>.Fail(ex.Message);
            }
        }

        return ServiceLayerResult<T>.Fail("Max retries exceeded");
    }

    public async Task<ServiceLayerResult<string>> GetStringAsync(SapCompany company, string resource, string? select = null, string? filter = null, int? top = null)
    {
        var retryCount = 0;

        while (retryCount <= _settings.RetryCount)
        {
            try
            {
                var slConn = company.Connection;
                var request = slConn.Request(resource);
                if (!string.IsNullOrEmpty(filter))
                    request = request.Filter(filter);
                if (!string.IsNullOrEmpty(select))
                    request = request.Select(select);

                request = request.WithCaseInsensitive();

                if (top.HasValue)
                {
                    request = request.Top(top.Value)
                        .WithHeader("Prefer", $"odata.maxpagesize={top.Value}");
                }

                var data = await request.GetStringAsync();

                LogSuccess("GET", company, resource, select, filter);
                return ServiceLayerResult<string>.Ok(data);
            }
            catch (Exception ex)
            {
                if (ShouldRetry(ex) && retryCount < _settings.RetryCount)
                {
                    LogRetry("GET", company, resource, retryCount);
                    await Task.Delay(1000);
                    retryCount++;
                    continue;
                }

                LogError("GET", company, resource, ex, select, filter);
                return ServiceLayerResult<string>.Fail(ex.Message);
            }
        }

        return ServiceLayerResult<string>.Fail("Max retries exceeded");
    }

    public async Task<ServiceLayerResult<T>> PostAsync<T>(SapCompany company, string resource, object data)
    {
        try
        {
            var result = await company.Connection.Request(resource).PostAsync<T>(data);
            
            _logger.LogInformation("[SL] SUCCESS: POST to {Company} - {Resource}", company.Name, resource);
            if (_settings.Debug)
            {
                _logger.LogDebug("[SL] POST Data: {Data}", JsonConvert.SerializeObject(data));
            }
            
            return ServiceLayerResult<T>.Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[SL] ERROR: POST to {Company} - {Resource}: {Message}", 
                company.Name, resource, ex.Message);
            return ServiceLayerResult<T>.Fail(ex.Message);
        }
    }

    public async Task<ServiceLayerResult<bool>> PatchAsync(SapCompany company, string resource, object data)
    {
        try
        {
            // Log outgoing PATCH payload for debugging
            _logger.LogInformation("[SL] PATCH PAYLOAD to {Company} - {Resource}: {Payload}", company.Name, resource, JsonConvert.SerializeObject(data));
            await company.Connection.Request(resource).PatchAsync(data);
            _logger.LogInformation("[SL] SUCCESS: PATCH to {Company} - {Resource}", company.Name, resource);
            return ServiceLayerResult<bool>.Ok(true);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[SL] ERROR: PATCH to {Company} - {Resource}: {Message}", 
                company.Name, resource, ex.Message);
            return ServiceLayerResult<bool>.Fail(ex.Message);
        }
    }

    public async Task<ServiceLayerResult<bool>> DeleteAsync(SapCompany company, string resource)
    {
        try
        {
            await company.Connection.Request(resource).DeleteAsync();
            
            _logger.LogInformation("[SL] SUCCESS: DELETE in {Company} - {Resource}", company.Name, resource);
            return ServiceLayerResult<bool>.Ok(true);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[SL] ERROR: DELETE in {Company} - {Resource}: {Message}", 
                company.Name, resource, ex.Message);
            return ServiceLayerResult<bool>.Fail(ex.Message);
        }
    }

    private bool ShouldRetry(Exception ex)
    {
        return ex.Message.Contains("Deadlock (-2038) detected during transaction") 
            || ex.Message.Contains("Call timed out:");
    }

    private void LogSuccess(string operation, SapCompany company, string resource, string? select, string? filter)
    {
        _logger.LogInformation("[SL] SUCCESS: {Operation} from {Company} - [{Resource}] {Select} {Filter}",
            operation, company.Name, resource,
            !string.IsNullOrEmpty(select) ? $"Select: [{select}]" : "",
            !string.IsNullOrEmpty(filter) ? $"Filter: [{filter}]" : "");
    }

    private void LogRetry(string operation, SapCompany company, string resource, int attempt)
    {
        _logger.LogWarning("[SL] RETRY ({Attempt}): {Operation} from {Company} - {Resource}",
            attempt + 1, operation, company.Name, resource);
    }

    private void LogError(string operation, SapCompany company, string resource, Exception ex, string? select, string? filter)
    {
        _logger.LogError(ex, "[SL] ERROR: {Operation} from {Company} - [{Resource}] {Select} {Filter}: {Message}",
            operation, company.Name, resource,
            !string.IsNullOrEmpty(select) ? $"Select: [{select}]" : "",
            !string.IsNullOrEmpty(filter) ? $"Filter: [{filter}]" : "",
            ex.Message);
    }
}
