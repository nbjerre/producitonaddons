using Microsoft.AspNetCore.Mvc;
using WorksheetAPI.Interfaces;
using Newtonsoft.Json.Linq;

namespace WebAPI.Controllers.userscontroller;

/// <summary>
/// Endpoints til læsning og vedligeholdelse af salgsmedarbejdere.
/// </summary>
[ApiController]
[Route("api/salespersons")]
public class SalesPersonsController : ControllerBase
{
    private readonly ISapConnectionService _sapConnection;
    private readonly IServiceLayerService _serviceLayer;
    private readonly ILogger<SalesPersonsController> _logger;

    public SalesPersonsController(
        ISapConnectionService sapConnection,
        IServiceLayerService serviceLayer,
        ILogger<SalesPersonsController> logger)
    {
        _sapConnection = sapConnection;
        _serviceLayer = serviceLayer;
        _logger = logger;
    }

    /// <summary>
    /// Henter aktive salgsmedarbejdere.
    /// </summary>
    [HttpGet]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetAll()
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();

        var result = await _serviceLayer.GetStringAsync(
            company,
            "SalesPersons",
            filter: "Active eq 'tYES'",
            top: 9999
        );

        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get sales persons" });

        return Content(BuildSuccessResponse(result.Data), "application/json");
    }

    /// <summary>
    /// Henter en enkelt salgsmedarbejder via SalesEmployeeCode.
    /// </summary>
    [HttpGet("{slpCode:int}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> GetOne(int slpCode)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();
        var result = await _serviceLayer.GetStringAsync(company, $"SalesPersons({slpCode})");
        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to get sales person" });

        return Content(BuildSuccessResponseSingle(result.Data), "application/json");
    }

    /// <summary>
    /// Opretter en ny salgsmedarbejder i SAP.
    /// </summary>
    [HttpPost]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> Create([FromBody] SalesPersonCreateDto dto)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        if (string.IsNullOrWhiteSpace(dto.SalesEmployeeName))
            return BadRequest(new { success = false, error = "SalesEmployeeName is required" });

        var company = _sapConnection.GetMainCompany();

        var jObject = new JObject
        {
            ["SalesEmployeeName"] = dto.SalesEmployeeName.Trim()
        };
        if (!string.IsNullOrEmpty(dto.Remarks))
            jObject["Remarks"] = dto.Remarks;

        jObject["Active"] = string.IsNullOrEmpty(dto.Active) ? "tYES" : dto.Active;

        _logger.LogInformation("[SalesPersonsController] Creating sales person: {Name}", dto.SalesEmployeeName);

        // Use JObject (not string) — SAP Service Layer returns a JSON object on POST,
        // and PostAsync<string> throws "Can not convert Object to String".
        var result = await _serviceLayer.PostAsync<JObject>(company, "SalesPersons", jObject);
        if (!result.Success)
        {
            _logger.LogWarning("[SalesPersonsController] Failed to create sales person: {Error}", result.Error);
            return BadRequest(new { success = false, error = result.Error ?? "Failed to create sales person" });
        }

        // result.Data is already a JObject — read the assigned SalesEmployeeCode directly
        var code = result.Data?["SalesEmployeeCode"]?.Value<int>();
        var name = result.Data?["SalesEmployeeName"]?.Value<string>();
        _logger.LogInformation("[SalesPersonsController] Sales person created: #{Code} {Name}", code, name);
        return Ok(new { success = true, data = new { SalesEmployeeCode = code, SalesEmployeeName = name } });
    }

    /// <summary>
    /// Opdaterer en eksisterende salgsmedarbejder.
    /// </summary>
    [HttpPatch("{slpCode:int}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> Update(int slpCode, [FromBody] SalesPersonCreateDto dto)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();
        var jObject = new JObject();
        if (!string.IsNullOrEmpty(dto.SalesEmployeeName)) jObject["SalesEmployeeName"] = dto.SalesEmployeeName;
        if (!string.IsNullOrEmpty(dto.Remarks)) jObject["Remarks"] = dto.Remarks;
        if (!string.IsNullOrEmpty(dto.Active)) jObject["Active"] = dto.Active;

        var result = await _serviceLayer.PatchAsync(company, $"SalesPersons({slpCode})", jObject);
        if (!result.Success)
            return BadRequest(new { success = false, error = result.Error ?? "Failed to update sales person" });

        return Ok(new { success = true });
    }

    /// <summary>
    /// Sletter en salgsmedarbejder.
    /// </summary>
    [HttpDelete("{slpCode:int}")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(object), StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> Delete(int slpCode)
    {
        if (!_sapConnection.IsSapEnabled)
            return BadRequest(new { success = false, error = "SAP integration is not enabled" });

        var company = _sapConnection.GetMainCompany();

        _logger.LogInformation("[SalesPersonsController] Deleting sales person: #{SlpCode}", slpCode);

        var result = await _serviceLayer.DeleteAsync(company, $"SalesPersons({slpCode})");
        if (!result.Success)
        {
            _logger.LogWarning("[SalesPersonsController] Failed to delete sales person #{SlpCode}: {Error}", slpCode, result.Error);
            return BadRequest(new { success = false, error = result.Error ?? "Failed to delete sales person" });
        }

        _logger.LogInformation("[SalesPersonsController] Sales person deleted: #{SlpCode}", slpCode);
        return Ok(new { success = true });
    }

    // ─── Response helpers ─────────────────────────────────────────────────────

    private static string BuildSuccessResponse(string? sapResponse)
    {
        if (string.IsNullOrEmpty(sapResponse)) return "{\"success\":true,\"data\":[]}";
        try
        {
            var token = JToken.Parse(sapResponse);
            if (token is JObject obj && obj.TryGetValue("value", out var value))
                return $"{{\"success\":true,\"data\":{value}}}";
            return $"{{\"success\":true,\"data\":{token}}}";
        }
        catch { return "{\"success\":true,\"data\":[]}"; }
    }

    private static string BuildSuccessResponseSingle(string? sapResponse)
    {
        if (string.IsNullOrEmpty(sapResponse)) return "{\"success\":true,\"data\":null}";
        try
        {
            var token = JToken.Parse(sapResponse);
            if (token is JObject obj && obj.TryGetValue("value", out var value) && value is JArray arr && arr.Count > 0)
                return $"{{\"success\":true,\"data\":{arr[0]}}}";
            return $"{{\"success\":true,\"data\":{token}}}";
        }
        catch { return "{\"success\":true,\"data\":null}"; }
    }
}

// ─── DTO ──────────────────────────────────────────────────────────────────────

/// <summary>
/// Requestmodel til oprettelse eller opdatering af salgsmedarbejdere.
/// </summary>
public class SalesPersonCreateDto
{
    /// <summary>
    /// Navn på salgsmedarbejderen.
    /// </summary>
    public string? SalesEmployeeName { get; set; }

    /// <summary>
    /// Frivillig note eller kommentar.
    /// </summary>
    public string? Remarks { get; set; }

    /// <summary>
    /// Angiver om posten er aktiv, typisk tYES eller tNO.
    /// </summary>
    public string? Active { get; set; }
}