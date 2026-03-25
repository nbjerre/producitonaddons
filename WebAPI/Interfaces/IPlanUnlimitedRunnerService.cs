using WorksheetAPI.Models;

namespace WorksheetAPI.Interfaces;

public interface IPlanUnlimitedRunnerService
{
    PlanUnlimitedHealthResult GetHealth();
    Task<PlanUnlimitedRunResult> RunAsync(PlanUnlimitedRunRequest request, CancellationToken cancellationToken = default);
}

public class PlanUnlimitedHealthResult
{
    public bool Enabled { get; set; }
    public string ExePath { get; set; } = string.Empty;
    public bool ExeExists { get; set; }
    public string WorkingDirectory { get; set; } = string.Empty;
    public DateTime? ExeLastWriteTime { get; set; }
}

public class PlanUnlimitedRunResult
{
    public bool Success { get; set; }
    public int ExitCode { get; set; }
    public string Output { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
    public string? Message { get; set; }
    public string ExePath { get; set; } = string.Empty;
    public DateTime? ExeLastWriteTime { get; set; }
}
