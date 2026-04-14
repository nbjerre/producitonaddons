using Newtonsoft.Json.Linq;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;

namespace WebAPI.Tests.Support;

public sealed class FakeSapConnectionService : ISapConnectionService
{
    private readonly SapCompany _company;

    public FakeSapConnectionService(bool isEnabled = true)
    {
        IsSapEnabled = isEnabled;
        _company = new SapCompany
        {
            Active = true,
            Name = "TEST",
            CompanyDb = "TESTDB",
            UserName = "manager",
            Password = "manager",
            ServiceLayerUrl = "https://sap.test.local"
        };
    }

    public bool IsSapEnabled { get; set; }

    public SapCompany GetMainCompany() => _company;

    public SapCompany? GetCompany(string name)
    {
        return string.Equals(name, _company.Name, StringComparison.OrdinalIgnoreCase) ? _company : null;
    }

    public IEnumerable<SapCompany> GetAllCompanies() => new[] { _company };
}

public sealed class FakePlanUnlimitedRunnerService : IPlanUnlimitedRunnerService
{
    public PlanUnlimitedHealthResult HealthResult { get; set; } = new()
    {
        Enabled = true,
        ExeExists = true,
        ExePath = "C:/plan/runner.exe",
        WorkingDirectory = "C:/plan"
    };

    public Func<PlanUnlimitedRunRequest, CancellationToken, Task<PlanUnlimitedRunResult>> RunHandler { get; set; } =
        (_, _) => Task.FromResult(new PlanUnlimitedRunResult { Success = true, ExitCode = 0, Message = "ok" });

    public PlanUnlimitedRunRequest? LastRunRequest { get; private set; }

    public PlanUnlimitedHealthResult GetHealth() => HealthResult;

    public Task<PlanUnlimitedRunResult> RunAsync(PlanUnlimitedRunRequest request, CancellationToken cancellationToken = default)
    {
        LastRunRequest = request;
        return RunHandler(request, cancellationToken);
    }
}

public sealed class FakeServiceLayerService : IServiceLayerService
{
    public Func<SapCompany, string, string?, string?, int?, Task<ServiceLayerResult<string>>>? GetStringHandler { get; set; }
    public Func<SapCompany, string, object, Task<ServiceLayerResult<JObject>>>? PostJObjectHandler { get; set; }

    public Task<ServiceLayerResult<T>> GetAsync<T>(SapCompany company, string resource, string? select = null, string? filter = null, int? top = null)
    {
        return Task.FromResult(ServiceLayerResult<T>.Fail("Not configured in test"));
    }

    public Task<ServiceLayerResult<string>> GetStringAsync(SapCompany company, string resource, string? select = null, string? filter = null, int? top = null)
    {
        if (GetStringHandler == null)
        {
            return Task.FromResult(ServiceLayerResult<string>.Fail("GetStringAsync not configured"));
        }

        return GetStringHandler(company, resource, select, filter, top);
    }

    public Task<ServiceLayerResult<T>> PostAsync<T>(SapCompany company, string resource, object data)
    {
        if (typeof(T) == typeof(JObject) && PostJObjectHandler != null)
        {
            return PostJObjectHandler(company, resource, data)
                .ContinueWith(task => (ServiceLayerResult<T>)(object)task.Result);
        }

        return Task.FromResult(ServiceLayerResult<T>.Fail("PostAsync not configured"));
    }

    public Task<ServiceLayerResult<bool>> PatchAsync(SapCompany company, string resource, object data)
    {
        return Task.FromResult(ServiceLayerResult<bool>.Fail("PatchAsync not configured"));
    }

    public Task<ServiceLayerResult<bool>> DeleteAsync(SapCompany company, string resource)
    {
        return Task.FromResult(ServiceLayerResult<bool>.Fail("DeleteAsync not configured"));
    }
}
