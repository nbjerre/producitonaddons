using System.Diagnostics;
using Microsoft.Extensions.Options;
using WorksheetAPI.Configuration;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;

namespace WorksheetAPI.Services;

public class PlanUnlimitedRunnerService : IPlanUnlimitedRunnerService
{
    private readonly PlanUnlimitedSettings _settings;
    private readonly ILogger<PlanUnlimitedRunnerService> _logger;

    public PlanUnlimitedRunnerService(
        IOptions<PlanUnlimitedSettings> settings,
        ILogger<PlanUnlimitedRunnerService> logger)
    {
        _settings = settings.Value;
        _logger = logger;
    }

    public PlanUnlimitedHealthResult GetHealth()
    {
        var exePath = _settings.ExePath ?? string.Empty;
        var workingDirectory = string.IsNullOrWhiteSpace(exePath)
            ? string.Empty
            : Path.GetDirectoryName(exePath) ?? string.Empty;
        DateTime? exeLastWrite = null;
        if (!string.IsNullOrWhiteSpace(exePath) && File.Exists(exePath))
        {
            exeLastWrite = File.GetLastWriteTime(exePath);
        }

        return new PlanUnlimitedHealthResult
        {
            Enabled = _settings.Enabled,
            ExePath = exePath,
            ExeExists = !string.IsNullOrWhiteSpace(exePath) && File.Exists(exePath),
            WorkingDirectory = workingDirectory,
            ExeLastWriteTime = exeLastWrite
        };
    }

    public async Task<PlanUnlimitedRunResult> RunAsync(PlanUnlimitedRunRequest request, CancellationToken cancellationToken = default)
    {
        var health = GetHealth();
        if (!health.Enabled)
        {
            return new PlanUnlimitedRunResult
            {
                Success = false,
                Message = "PlanUnlimited integration is disabled",
                ExePath = health.ExePath,
                ExeLastWriteTime = health.ExeLastWriteTime
            };
        }

        if (!health.ExeExists)
        {
            return new PlanUnlimitedRunResult
            {
                Success = false,
                Message = "PlanUnlimited executable was not found",
                ExePath = health.ExePath,
                ExeLastWriteTime = health.ExeLastWriteTime
            };
        }

        var user = string.IsNullOrWhiteSpace(request.User) ? "manager" : request.User.Trim();

        var startInfo = new ProcessStartInfo
        {
            FileName = health.ExePath,
            WorkingDirectory = health.WorkingDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        startInfo.ArgumentList.Add(request.Btn);
        startInfo.ArgumentList.Add(request.DocEntry.ToString());
        startInfo.ArgumentList.Add(request.LineNum.ToString());
        startInfo.ArgumentList.Add(request.ProdDocNum.ToString());
        startInfo.ArgumentList.Add(user);

        using var process = new Process { StartInfo = startInfo };

        try
        {
            _logger.LogInformation("Starting PlanUnlimited: {ExePath} {Btn} {DocEntry} {LineNum} {ProdDocNum} {User}",
                startInfo.FileName,
                request.Btn,
                request.DocEntry,
                request.LineNum,
                request.ProdDocNum,
                user);

            process.Start();

            var outputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var errorTask = process.StandardError.ReadToEndAsync(cancellationToken);

            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(Math.Max(10, _settings.TimeoutSeconds))); 

            await process.WaitForExitAsync(timeoutCts.Token);
            var output = await outputTask;
            var error = await errorTask;

            return new PlanUnlimitedRunResult
            {
                Success = process.ExitCode == 0,
                ExitCode = process.ExitCode,
                Output = output,
                Error = error,
                Message = process.ExitCode == 0 ? "PlanUnlimited completed" : "PlanUnlimited failed",
                ExePath = health.ExePath,
                ExeLastWriteTime = health.ExeLastWriteTime
            };
        }
        catch (OperationCanceledException)
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
            }

            return new PlanUnlimitedRunResult
            {
                Success = false,
                ExitCode = -1,
                Message = "PlanUnlimited timed out",
                ExePath = health.ExePath,
                ExeLastWriteTime = health.ExeLastWriteTime
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "PlanUnlimited execution failed");
            return new PlanUnlimitedRunResult
            {
                Success = false,
                ExitCode = -1,
                Message = ex.Message,
                ExePath = health.ExePath,
                ExeLastWriteTime = health.ExeLastWriteTime
            };
        }
    }
}
