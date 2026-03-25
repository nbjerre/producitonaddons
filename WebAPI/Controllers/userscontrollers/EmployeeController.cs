using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using WorksheetAPI.Interfaces;
using WorksheetAPI.Services;
using WorksheetAPI.Models;

namespace WebAPI.Controllers.userscontroller;

/// <summary>
/// Endpoints til læsning og vedligeholdelse af medarbejdere i SAP.
/// </summary>
[ApiController]
[Route("api/employees")]
public class EmployeesController : ControllerBase
{
    private readonly ISapConnectionService _sapConnection;
    private readonly IServiceLayerService _serviceLayer;
    private readonly ILogger<EmployeesController> _logger;

    public EmployeesController(
        ISapConnectionService sapConnection,
        IServiceLayerService serviceLayer,
        ILogger<EmployeesController> logger)
    {
        _sapConnection = sapConnection;
        _serviceLayer = serviceLayer;
        _logger = logger;
    }

    /// <summary>
    /// Henter medarbejdere fra EmployeesInfo.
    /// </summary>
    [HttpGet]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetAllEmployees([FromQuery] int? top = 100, [FromQuery] int? skip = null)
    {
        _logger.LogInformation("[EmployeesController] GET /api/employees kaldt");
        if (!_sapConnection.IsSapEnabled)
        {
            _logger.LogWarning("[EmployeesController] SAP integration is not enabled");
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });
        }
        var company = _sapConnection.GetMainCompany();
        var resource = "EmployeesInfo";
        _logger.LogInformation("[EmployeesController] Henter data fra SAP. Company: {Company}, Resource: {Resource}", company.Name, resource);

        // Pass top parameter directly, skip is not supported in interface
        var result = await _serviceLayer.GetStringAsync(company, resource, select: null, filter: null, top: top);
        if (!result.Success)
        {
            _logger.LogWarning("[EmployeesController] Fejl ved hentning af medarbejdere: {Error}", result.Error);
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get employees" });
        }
        return Content(BuildSuccessResponse(result.Data), "application/json");
    }

    /// <summary>
    /// Henter en medarbejder via EmployeeID.
    /// </summary>
    [HttpGet("{id}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetEmployee(string id)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });
        var company = _sapConnection.GetMainCompany();
        // EmployeesInfo uses integer key — filter by EmployeeID
        var result = await _serviceLayer.GetStringAsync(company, "EmployeesInfo", filter: $"EmployeeID eq {id}");
        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get employee" });
        return Content(BuildSuccessResponseFirst(result.Data), "application/json");
    }

    /// <summary>
    /// Opretter en medarbejder i SAP.
    /// </summary>
    [HttpPost]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> CreateEmployee([FromBody] JObject employee)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });
        var company = _sapConnection.GetMainCompany();
        // Fix: was "Users" — must be "EmployeesInfo"
        var result = await _serviceLayer.PostAsync<string>(company, "EmployeesInfo", employee);
        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to create employee" });
        return Content(BuildSuccessResponseFirst(result.Data), "application/json");
    }

    /// <summary>
    /// Opdaterer en medarbejder via EmployeeUpdateDto.
    /// </summary>
    [HttpPut("{id}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> UpdateEmployee(string id, [FromBody] EmployeeUpdateDto employee)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });
        var company = _sapConnection.GetMainCompany();
        var resource = $"EmployeesInfo({id})";
        // Convert DTO → JObject for the service layer
        var jObject = JObject.FromObject(employee);
        var result = await _serviceLayer.PatchAsync(company, resource, jObject);
        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to update employee" });
        return Ok(new { success = true });
    }

    // ─── Response helpers ────────────────────────────────────────────────────

    private static string BuildSuccessResponse(string? sapResponse)
    {
        if (string.IsNullOrEmpty(sapResponse))
            return "{\"success\":true,\"data\":[]}";
        try
        {
            var token = JToken.Parse(sapResponse);
            if (token is JObject obj && obj.TryGetValue("value", out var value))
                return $"{{\"success\":true,\"data\":{value}}}";
            return $"{{\"success\":true,\"data\":{token}}}";
        }
        catch { return "{\"success\":true,\"data\":[]}"; }
    }

    private static string BuildSuccessResponseFirst(string? sapResponse)
    {
        if (string.IsNullOrEmpty(sapResponse))
            return "{\"success\":true,\"data\":null}";
        try
        {
            var token = JToken.Parse(sapResponse);
            if (token is JObject obj && obj.TryGetValue("value", out var value) && value is JArray arr && arr.Count > 0)
                return $"{{\"success\":true,\"data\":{arr[0]}}}";
            return "{\"success\":true,\"data\":null}";
        }
        catch { return "{\"success\":true,\"data\":null}"; }
    }
}