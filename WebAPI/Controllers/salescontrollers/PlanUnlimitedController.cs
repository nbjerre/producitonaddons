using Microsoft.AspNetCore.Mvc;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Models;

namespace WebAPI.Controllers.salescontrollers;

/// <summary>
/// Endpoints til planlægning via den eksterne PlanUnlimited-runner.
/// </summary>
[ApiController]
[Route("api/plan")]
public class PlanUnlimitedController : ControllerBase
{
    private static readonly HashSet<string> AllowedButtons = new(StringComparer.OrdinalIgnoreCase)
    {
        "BtnPU",
        "BtnPUL",
        "BtnPR",
        "BtnPl",
        "BtnPlu"
    };

    private readonly IPlanUnlimitedRunnerService _runner;

    public PlanUnlimitedController(IPlanUnlimitedRunnerService runner)
    {
        _runner = runner;
    }

    /// <summary>
    /// Returnerer health-status for PlanUnlimited-runneren.
    /// </summary>
    [HttpGet("health")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    public IActionResult Health()
    {
        var result = _runner.GetHealth();
        return Ok(result);
    }

    /// <summary>
    /// Kører BtnPR-flowet for en salgsordre.
    /// </summary>
    [HttpPost("order/btnpr")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(object), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> RunBtnPr([FromBody] PlanSalesOrderRequest request, CancellationToken cancellationToken)
    {
        if (request.DocEntry <= 0)
        {
            return BadRequest(new { success = false, error = "DocEntry is required" });
        }

        var runRequest = new PlanUnlimitedRunRequest
        {
            Btn = "BtnPR",
            DocEntry = request.DocEntry,
            LineNum = -1,
            ProdDocNum = -1,
            User = request.User
        };

        var result = await _runner.RunAsync(runRequest, cancellationToken);
        return BuildRunResponse(result);
    }

    /// <summary>
    /// Kører BtnPU-flowet for en salgsordre.
    /// </summary>
    [HttpPost("order/btnpu")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(object), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> RunBtnPu([FromBody] PlanSalesOrderRequest request, CancellationToken cancellationToken)
    {
        if (request.DocEntry <= 0)
        {
            return BadRequest(new { success = false, error = "DocEntry is required" });
        }

        var runRequest = new PlanUnlimitedRunRequest
        {
            Btn = "BtnPU",
            DocEntry = request.DocEntry,
            LineNum = -1,
            ProdDocNum = -1,
            User = request.User
        };

        var result = await _runner.RunAsync(runRequest, cancellationToken);
        return BuildRunResponse(result);
    }

    /// <summary>
    /// Kører BtnPUL-flowet for en bestemt salgsordrelinje.
    /// </summary>
    [HttpPost("order/btnpul")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(object), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> RunBtnPul([FromBody] PlanSalesOrderRequest request, CancellationToken cancellationToken)
    {
        if (request.DocEntry <= 0)
        {
            return BadRequest(new { success = false, error = "DocEntry is required" });
        }

        if (!request.LineNum.HasValue || request.LineNum.Value < 0)
        {
            return BadRequest(new { success = false, error = "LineNum is required for BtnPUL" });
        }

        var runRequest = new PlanUnlimitedRunRequest
        {
            Btn = "BtnPUL",
            DocEntry = request.DocEntry,
            LineNum = request.LineNum.Value,
            ProdDocNum = -1,
            User = request.User
        };

        var result = await _runner.RunAsync(runRequest, cancellationToken);
        return BuildRunResponse(result);
    }

    /// <summary>
    /// Kører et generisk PlanUnlimited-kald baseret på knaptype og inputdata.
    /// </summary>
    [HttpPost("run")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(object), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> Run([FromBody] PlanUnlimitedRunRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.Btn) || !AllowedButtons.Contains(request.Btn))
        {
            return BadRequest(new
            {
                success = false,
                error = "Invalid btn. Allowed values: BtnPU, BtnPUL, BtnPR, BtnPl, BtnPlu"
            });
        }

        if (request.DocEntry <= 0 && request.ProdDocNum <= 0)
        {
            return BadRequest(new
            {
                success = false,
                error = "DocEntry (sales order) or ProdDocNum (production order) must be provided"
            });
        }

        var result = await _runner.RunAsync(request, cancellationToken);
        return BuildRunResponse(result);
    }

    private IActionResult BuildRunResponse(PlanUnlimitedRunResult result)
    {
        if (!result.Success)
        {
            return StatusCode(500, new
            {
                success = false,
                message = result.Message,
                result.ExitCode,
                result.Output,
                result.Error,
                result.ExePath,
                result.ExeLastWriteTime
            });
        }

        return Ok(new
        {
            success = true,
            message = result.Message,
            result.ExitCode,
            result.Output,
            result.Error,
            result.ExePath,
            result.ExeLastWriteTime
        });
    }
}
